import streamlit as st
import streamlit.components.v1 as components
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeResult
import google.generativeai as genai
from openai import OpenAI
import json
import time
import concurrent.futures
import pandas as pd
from thefuzz import fuzz
from collections import Counter
import re

# --- 1. 頁面設定 ---
st.set_page_config(page_title="交貨單稽核", page_icon="🏭", layout="centered")

# --- CSS 樣式 ---
st.markdown("""
<style>
/* 1. 標題大小控制 */
h1 {
    font-size: 1.7rem !important; 
    white-space: nowrap !important;
    overflow: hidden !important; 
    text-overflow: ellipsis !important;
}

/* 2. 主功能按鈕 (紅色 Primary) -> 變大、變高 */
/* 這會影響「開始分析」和「照片清除」 */
button[kind="primary"] {
    height: 60px;               
    font-size: 20px !important; 
    font-weight: bold !important;
    border-radius: 10px !important;
    margin-top: 0px !important;    
    margin-bottom: 5px !important; 
    width: 100%;                
}

/* 3. 次要按鈕 (灰色 Secondary) -> 保持原狀 */
/* 這會影響每一張照片下面的「X」按鈕，讓它維持小小的 */
button[kind="secondary"] {
    height: auto !important;
    font-weight: normal !important;
}
</style>
""", unsafe_allow_html=True)
# --- 2. 秘密金鑰讀取 ---
try:
    DOC_ENDPOINT = st.secrets["DOC_ENDPOINT"]
    DOC_KEY = st.secrets["DOC_KEY"]
    GEMINI_KEY = st.secrets["GEMINI_KEY"]
    OPENAI_KEY = st.secrets.get("OPENAI_KEY", "")
except:
    st.error("找不到金鑰！請在 Streamlit Cloud 設定 Secrets。")
    st.stop()

# --- 3. 初始化 Session State ---
if 'photo_gallery' not in st.session_state: st.session_state.photo_gallery = []
if 'uploader_key' not in st.session_state: st.session_state.uploader_key = 0
if 'auto_start_analysis' not in st.session_state: st.session_state.auto_start_analysis = False

# --- 側邊欄模型設定 (合併為單一選擇) ---
with st.sidebar:
    st.header("模型設定")
    
    # 這裡加入最新的 Gemini 模型
    model_options = {
        "Gemini 3 Flash preview": "gemini-3-flash-preview",
        "Gemini 2.5 Flash": "models/gemini-2.5-flash",
        "Gemini 2.5 Pro": "models/gemini-2.5-pro",
        #"GPT-5(無效)": "models/gpt-5",
        #"GPT-5 Mini(無效)": "models/gpt-5-mini",
    }
    options_list = list(model_options.keys())
    
    st.subheader("🤖 總稽核 Agent")
    model_selection = st.selectbox(
        "負責：規格、製程、數量、統計全包", 
        options=options_list, 
        index=0, 
        key="main_model"
    )
    main_model_name = model_options[model_selection]
    
    st.divider()
    
    default_auto = st.query_params.get("auto", "true") == "true"
    def update_url_param():
        current_state = "true" if st.session_state.enable_auto_analysis else "false"
        st.query_params["auto"] = current_state

    st.toggle(
        "⚡ 上傳後自動分析", 
        value=default_auto, 
        key="enable_auto_analysis", 
        on_change=update_url_param
    )

# --- Excel 規則讀取函數 (單一代理整合版) ---
@st.cache_data
def get_dynamic_rules(ocr_text, debug_mode=False):
    try:
        df = pd.read_excel("rules.xlsx")
        df.columns = [c.strip() for c in df.columns]
        ocr_text_clean = str(ocr_text).upper().replace(" ", "").replace("\n", "")
        specific_rules = []

        for index, row in df.iterrows():
            item_name = str(row.get('Item_Name', '')).strip()
            # 💡 跳過原本的「(通用)」項目，只抓特規
            if not item_name or "(通用)" in item_name: continue
            
            # 使用模糊匹配判斷是否為當前處理的項目
            score = fuzz.partial_ratio(item_name.upper().replace(" ", ""), ocr_text_clean)
            if score >= 85:
                # 提取特規資訊
                spec = str(row.get('Standard_Spec', ''))
                logic = str(row.get('Logic_Prompt', ''))
                u_local = str(row.get('Unit_Rule_Local', ''))
                u_agg = str(row.get('Unit_Rule_Agg', ''))
                u_freight = str(row.get('Unit_Rule_Freight', ''))
                
                desc = f"- **[特定項目規則] {item_name}**\n"
                if spec != 'nan' and spec: desc += f"  - [強制規格]: {spec}\n"
                if logic != 'nan' and logic: desc += f"  - [例外指令]: {logic}\n"
                if u_local != 'nan' and u_local: desc += f"  - [會計單項]: {u_local}\n"
                if u_agg != 'nan' and u_agg: desc += f"  - [會計聚合]: {u_agg}\n"
                if u_freight != 'nan' and u_freight: desc += f"  - [會計運費]: {u_freight}\n"
                specific_rules.append(desc)
        
        return "\n".join(specific_rules) if specific_rules else "無特定專案規則，請依照通用憲法執行。"
    except Exception as e:
        return f"讀取規則檔時發生錯誤: {e}"
        
# --- 4. 核心函數：Azure 神之眼 ---
def extract_layout_with_azure(file_obj, endpoint, key):
    client = DocumentIntelligenceClient(endpoint=endpoint, credential=AzureKeyCredential(key))
    file_content = file_obj.getvalue()
    
    poller = client.begin_analyze_document("prebuilt-layout", file_content, content_type="application/octet-stream")
    result: AnalyzeResult = poller.result()
    
    markdown_output = ""
    full_content_text = ""
    real_page_num = "Unknown"
    
    bottom_stop_keywords = ["注意事項", "中機品檢單位", "保存期限", "表單編號", "FORM NO", "簽章"]
    top_right_noise_keywords = [
        "檢驗類別", "尺寸檢驗", "依圖面標記", "材料檢驗", "成份分析", 
        "非破壞性", "正常化", "退火", "淬.回火", "表面硬化", "試車",
        "性能測試", "試壓試漏", "動.靜平衡試驗", ":selected:", ":unselected:",
        "抗拉", "硬度試驗", "UT", "PT", "MT"
    ]
    
    if result.tables:
        for idx, table in enumerate(result.tables):
            page_num = "Unknown"
            if table.bounding_regions: page_num = table.bounding_regions[0].page_number
            markdown_output += f"\n### Table {idx + 1} (Page {page_num}):\n"
            rows = {}
            stop_processing_table = False 
            
            for cell in table.cells:
                if stop_processing_table: break
                content = cell.content.replace("\n", " ").strip()
                
                for kw in bottom_stop_keywords:
                    if kw in content:
                        stop_processing_table = True
                        break
                if stop_processing_table: break
                
                is_noise = False
                for kw in top_right_noise_keywords:
                    if kw in content:
                        is_noise = True
                        break
                if is_noise: content = "" 

                r, c = cell.row_index, cell.column_index
                if r not in rows: rows[r] = {}
                rows[r][c] = content
            
            for r in sorted(rows.keys()):
                row_cells = []
                if rows[r]:
                    max_col = max(rows[r].keys())
                    for c in range(max_col + 1): 
                        row_cells.append(rows[r].get(c, ""))
                    markdown_output += "| " + " | ".join(row_cells) + " |\n"
    
    if result.content:
        match = re.search(r"(?:項次|Page|頁次|NO\.)[:\s]*(\d+)\s*[/／]\s*\d+", result.content, re.IGNORECASE)
        if match:
            real_page_num = match.group(1)

        cut_index = len(result.content)
        for keyword in bottom_stop_keywords:
            idx = result.content.find(keyword)
            if idx != -1 and idx < cut_index:
                cut_index = idx
        
        temp_text = result.content[:cut_index]
        for noise in top_right_noise_keywords:
            temp_text = temp_text.replace(noise, "")
            
        full_content_text = temp_text
        header_snippet = full_content_text[:800]
    else:
        full_content_text = ""
        header_snippet = ""

    return markdown_output, header_snippet, full_content_text, None, real_page_num

# --- Python 硬邏輯：表頭一致性檢查 (長度敏感版) ---
def python_header_check(photo_gallery):
    issues = []
    if not photo_gallery:
        return issues, []

    # 定義 Regex (針對 "去空白+去換行" 後的字串設計)
    patterns = {
        # 【修改點 1】工令 Regex 放寬：
        # 原本只抓 W 開頭，現在改抓 "編號" 後面接的 "任何英數字串"
        # 這樣就算它寫 WW363... 或是 12345... 都能整串抓出來比對
        "工令編號": r"[工土下][令冷今]編號[:\.]*([A-Za-z0-9\-\_]+)", 
        
        "預定交貨": r"[預预項頂][定交].*?(\d{2,4}[\.\-/]\d{1,2}[\.\-/]\d{1,2})",
        "實際交貨": r"[實真][際交].*?(\d{2,4}[\.\-/]\d{1,2}[\.\-/]\d{1,2})"
    }

    extracted_data = [] 
    all_values = {key: [] for key in patterns}

    for i, page in enumerate(photo_gallery):
        # 暴力清洗：去換行、去空格、轉大寫
        raw_text = page.get('header_text', '') + page.get('full_text', '')
        clean_text = raw_text.replace("\n", "").replace(" ", "").replace("\r", "").upper()
        
        # 【修改點 2】頁碼防呆：確保一定有值
        # 優先抓 real_page，抓不到就用 index
        r_page = page.get('real_page')
        if not r_page or r_page == "Unknown":
            page_label = f"P.{i + 1}"
        else:
            page_label = f"P.{r_page}"
            
        page_result = {"頁數": page_label}
        
        for key, pattern in patterns.items():
            match = re.search(pattern, clean_text)
            if match:
                val = match.group(1).strip()
                
                # 【修改點 3】針對工令的特殊處理 (如果太長可能就是重複打字)
                if key == "工令編號":
                    # 如果你確定工令只有 10 碼，但抓到了 11 碼以上 (如 WW...)
                    # 我們保留這個錯誤的值，讓後面的多數決去把它揪出來
                    pass 
                
                page_result[key] = val
                all_values[key].append(val)
            else:
                page_result[key] = "N/A"
        
        extracted_data.append(page_result)

    # 步驟 2: 決定「正確標準」 (使用多數決)
    standard_data = {}
    for key, values in all_values.items():
        if values:
            # 濾掉 N/A 後再投票
            valid_values = [v for v in values if v != "N/A"]
            if valid_values:
                most_common = Counter(valid_values).most_common(1)[0][0]
                standard_data[key] = most_common
            else:
                standard_data[key] = "N/A"
        else:
            standard_data[key] = "N/A"

    # 步驟 3: 比對每一頁
    for data in extracted_data:
        page_num = data['頁數']
        
        for key, standard_val in standard_data.items():
            current_val = data[key]
            
            if standard_val == "N/A": continue # 全卷都沒抓到就不比了

            # 開始比對 (字串不相等)
            if current_val != standard_val:
                
                # 判斷是否為長度異常 (針對工令)
                reason = "與全卷多數頁面不一致"
                if key == "工令編號" and len(current_val) != len(standard_val):
                    reason += f" (長度異常: {len(current_val)}碼 vs 標準{len(standard_val)}碼)"

                issue = {
                    "page": page_num.replace("P.", ""),
                    "item": f"表頭檢查-{key}",
                    "rule_used": "Python硬邏輯檢查",
                    "issue_type": "跨頁資訊不符",
                    "spec_logic": f"應為 {standard_val}",
                    "common_reason": reason,
                    "failures": [
                        {"id": "全卷基準", "val": standard_val, "calc": "多數決標準"},
                        {"id": f"本頁({page_num})", "val": current_val, "calc": "異常/漏抓"}
                    ],
                    "source": "🤖 系統自動"
                }
                issues.append(issue)
                
    return issues, extracted_data

    # --- 5. 總稽核 Agent (整合版 - 強邏輯優化) ---
def agent_unified_check(combined_input, full_text_for_search, api_key, model_name):
    dynamic_rules = get_dynamic_rules(full_text_for_search)

    system_prompt = f"""
    你是一位極度嚴謹的中鋼機械品管【數據抄錄員】。你必須像「電腦程式」一樣執行提取任務。
    
    {dynamic_rules}

    ---

    #### ⚔️ 模組 A：數據抄錄與分類 (AI 翻譯官)
    1. **規格抄錄 (std_spec)**：精確抄錄標題中含 `mm`、`±`、`+`、`-`、`至...再生` 的文字。
       - **🚫 禁令**：嚴禁執行加減法運算，保持 std_ranges 為空，將原始文字抄錄到 std_spec 即可。
    2. **數據抄錄 (ds)**：採用壓縮格式 `"ID:值|ID:值"`。
       - **字串保護**：實測值顯示 `349.90` 必寫 `"349.90"`。禁止簡化數字。
       - **壞軌標記 [!]**：若儲存格辨識不良（汙點/遮擋/黏連），嚴禁腦補，直接標記為 `[!]`。
    3. **分類識別 (category) 決策流**：
       - LEVEL 1：含「銲補」 -> `min_limit`。
       - LEVEL 2：含「未再生」。a.含「軸頸」-> `max_limit`；b.不含「軸頸」-> `un_regen`。
       - LEVEL 3：含「再生/研磨/精加工/車修/組裝/拆裝/真圓度」 -> `range`。

    #### 💰 模組 B：會計指標提取 (由 AI 抄錄傳票)
    1. **總表提取**：抄錄左上角統計表每一行的名稱與實交數量到 `summary_rows`。
    2. **指標提取**：提取運費項次到 `freight_target`，提取項目括號內的數字到 `item_pc_target`。
    3. **⚖️ 流程稽核**：檢查物理位階 `未再生 < 研磨 < 再生 < 銲補`。若跨頁面後段尺寸小於前段（銲補除外），報 `🛑流程異常`。

    ---
    
    ### 📝 輸出規範 (Output Format)
    必須回傳單一 JSON。統計不符時必須「逐行拆分」來源明細。

    {{
      "job_no": "工令",
      "summary_rows": [ {{ "title": "名", "target": 數字 }} ],
      "freight_target": 0,
      "issues": [ 
         {{ "page": "頁碼", "item": "項目", "issue_type": "統計不符 / 🛑流程異常", "common_reason": "原因", "failures": [] }}
      ],
      "dimension_data": [
         {{
           "page": 數字, "item_title": "標題", "category": "分類名稱", "item_pc_target": 0,
           "accounting_rules": {{ "local": "", "agg": "", "freight": "" }},
           "sl": {{ "lt": "分類標籤", "t": 0 }},
           "std_spec": "原始規格文字",
           "ds": "ID:值|ID:值" 
         }}
      ]
    }}
    """
    
    # 修改後的建議配置
    generation_config = {
    "temperature": 0.0,             # ⚡️ 設為 0：最快且最穩定，不讓 AI 多想
    "max_output_tokens": 4096,      # ⚡️ 先降回 4096：通常 4 頁資料這個長度就夠了，減少 AI 廢話
    # "response_mime_type": "application/json" # ⚡️ 暫時註解掉這行！
    }

    
    try:
        genai.configure(api_key=api_key)
        
        # 2. 設定 AI (開啟 JSON 模式以確保成功率)
        model = genai.GenerativeModel(
            model_name=model_name,
            generation_config={
                "temperature": 0.0,            # 最穩定
                "max_output_tokens": 8192,     # 給予足夠長度寫完大表
                "response_mime_type": "application/json" # ⚡️ 強制 JSON 模式 (避免解析失敗)
            },
            safety_settings={
                "HARM_CATEGORY_HARASSMENT": "BLOCK_NONE",
                "HARM_CATEGORY_HATE_SPEECH": "BLOCK_NONE",
                "HARM_CATEGORY_SEXUALLY_EXPLICIT": "BLOCK_NONE",
                "HARM_CATEGORY_DANGEROUS_CONTENT": "BLOCK_NONE",
            }
        )
        
        # 3. 呼叫 AI (這裡會跑 20-40 秒是正常的)
        with st.spinner('🤖 AI 正在全力抄寫數據中...'):
            response = model.generate_content([system_prompt, combined_input])
        
        # 4. 檢查是否有內容
        raw_content = response.text.strip()
        
        # 移除可能的多餘標記 (雙重保險)
        if raw_content.startswith("```json"):
            raw_content = raw_content[7:]
        if raw_content.endswith("```"):
            raw_content = raw_content[:-3]
        raw_content = raw_content.strip()

        # 5. 解析 JSON
        parsed_data = json.loads(raw_content)
        
        # 記錄 Token
        parsed_data["_token_usage"] = {
            "input": response.usage_metadata.prompt_token_count, 
            "output": response.usage_metadata.candidates_token_count
        }
        return parsed_data

    except json.JSONDecodeError as e:
        # 🚨 這裡就是抓出「為什麼跑了29秒卻失敗」的關鍵
        st.error("❌ JSON 解析失敗！請查看下方 AI 的原始回應：")
        with st.expander("👀 點擊查看 AI 到底回傳了什麼"):
            # 如果 AI 有回傳東西，印出來看
            if 'raw_content' in locals():
                st.code(raw_content)
            elif 'response' in locals():
                st.code(response.text)
        # 回傳錯誤結構，避免程式當機
        return {"job_no": "JSON Error", "issues": [], "dimension_data": []}

    except Exception as e:
        # 其他錯誤 (例如網路中斷、API 錯誤)
        st.error(f"❌ 系統發生錯誤: {str(e)}")
        return {"job_no": f"Error: {str(e)}", "issues": [], "dimension_data": []}
    
# --- 重點：Python 引擎獨立於 agent 函式之外 ---

def python_numerical_audit(dimension_data):
    grouped_errors = {}
    import re
    if not dimension_data: return []

    for item in dimension_data:
        ds = str(item.get("ds", ""))
        if not ds: continue
        raw_entries = [p.split(":") for p in ds.split("|") if ":" in p]
        
        title = str(item.get("item_title", "")).replace(" ", "").replace('"', "")
        raw_spec = str(item.get("std_spec", "")).replace('"', "")
        cat = str(item.get("category", "")).strip()
        page_num = item.get("page", "?")

        # 💡 [新增：Python 自動解析公差]
        s_ranges = []
        clean_part = raw_spec.replace(" ", "")
        pm = re.search(r"(\d+\.?\d*)?±(\d+\.?\d*)", clean_part)
        devs = re.findall(r"([+-]\d+\.?\d*)", clean_part)
        mm_match = re.findall(r"(\d+\.?\d*)mm", clean_part)
        clean_std = [float(n) for n in mm_match if float(n) > 5]

        if pm:
            b = float(pm.group(1)) if pm.group(1) else 0.0
            o = float(pm.group(2))
            s_ranges.append([round(b - o, 4), round(b + o, 4)])
        elif base_val := (clean_std[0] if clean_std else None):
            if len(devs) >= 2:
                calc_nums = [base_val + float(o) for o in devs]
                s_ranges.append([round(min(calc_nums), 4), round(max(calc_nums), 4)])

        for entry in raw_entries:
            if len(entry) < 2: continue
            rid, val_raw = entry[0].strip(), entry[1].strip()
            if not val_raw or val_raw in ["N/A", "nan"]: continue

            try:
                is_passed, reason, t_used, e_label = True, "", "N/A", "未知"
                
                # --- 7.1 壞軌偵測 ---
                if "[!]" in val_raw:
                    is_passed, reason, val_str, val = False, "🛑數據損壞(壞軌)", "[!]", -999.0
                else:
                    v_m = re.findall(r"\d+\.?\d*", val_raw)
                    val_str = v_m[0] if v_m else val_raw
                    val = float(val_str)

                # --- 7.2 格式判定 ---
                if val_str != "[!]":
                    is_two_dec = "." in val_str and len(val_str.split(".")[-1]) == 2
                    is_pure_int = "." not in val_str
                else: is_two_dec, is_pure_int = True, True

                # --- 7.3 判定邏輯 (銲補 > 未再生 > 精加工) ---
                if "min_limit" in cat or "銲補" in (cat + title):
                    e_label = "銲補"
                    t_used = min(clean_std) if clean_std else "N/A"
                    if not is_pure_int: is_passed, reason = False, "應為純整數"
                    elif t_used != "N/A" and val < t_used: is_passed, reason = False, "數值不足"
                
                elif "un_regen" in cat or "max_limit" in cat or "未再生" in (cat + title):
                    if "軸頸" in (cat + title):
                        e_label = "軸頸(上限)"
                        target = max(clean_std) if clean_std else 0
                        t_used = target
                        if target > 0 and val > target: is_passed, reason = False, f"超過上限 {target}"
                        if target > 0 and not is_pure_int: is_passed, reason = False, "應為純整數"
                    else:
                        e_label = "未再生(本體)"
                        candidates = [n for n in clean_std if n >= 120.0]
                        target = max(candidates) if candidates else 196.0
                        t_used = target
                        if val <= target and not is_pure_int: is_passed, reason = False, "應為整數"
                        elif val > target and not is_two_dec: is_passed, reason = False, "應填兩位小數"

                elif any(x in (cat + title) for x in ["再生", "精加工", "研磨", "車修", "組裝"]):
                    e_label = "精加工"
                    if not is_two_dec: is_passed, reason = False, "應填兩位小數"
                    elif s_ranges:
                        t_used = str(s_ranges)
                        if not any(r[0] <= val <= r[1] for r in s_ranges): is_passed, reason = False, "不在區間內"

                if not is_passed:
                    key = (page_num, title, reason)
                    if key not in grouped_errors:
                        grouped_errors[key] = {"page": page_num, "item": title, "issue_type": f"數值異常({e_label})", "common_reason": reason, "failures": []}
                    grouped_errors[key]["failures"].append({"id": rid, "val": val_str, "target": f"基準:{t_used}"})
            except: continue
    return list(grouped_errors.values())
    
def python_accounting_audit(dimension_data, res_main):
    """
    Python 會計官：【自動查表完全體】
    整合功能：單項核對(去重)、軸頸限次、KG重量累加、A/B總表模式、運費動態解析。
    """
    accounting_issues = []
    from thefuzz import fuzz
    from collections import Counter
    import re
    import pandas as pd

    # 💡 [新增：自動查表準備] 讀取全域 Excel 規則檔
    try:
        df_rules = pd.read_excel("rules.xlsx")
        df_rules.columns = [c.strip() for c in df_rules.columns]
    except:
        df_rules = None

    # 💡 [輔助工具：安全轉型數字] 
    def safe_float(value):
        if value is None or str(value).upper() == 'NULL': return 0.0
        val_str = str(value).strip()
        if "[!]" in val_str: return "BAD_DATA" 
        cleaned = "".join(re.findall(r"[\d\.]+", val_str.replace(',', '')))
        try: return float(cleaned) if cleaned else 0.0
        except: return 0.0

    # 1. 取得對帳基準 (來自左上角統計表)
    summary_rows = res_main.get("summary_rows", [])
    global_sum_tracker = {
        s['title']: {"target": safe_float(s['target']), "actual": 0, "details": []} 
        for s in summary_rows if s.get('title')
    }
    
    freight_target = safe_float(res_main.get("freight_target", 0))
    freight_actual_sum = 0
    freight_details = []

    # 2. 開始逐項遍歷
    for item in dimension_data:
        title = item.get("item_title", "")
        page = item.get("page", "?")
        target_pc = safe_float(item.get("item_pc_target", 0)) 
        
        # 💡 [關鍵功能：Python 自動查表補位]
        # 不再依賴 AI 抄錄，直接從標題匹配 Excel 裡的三個會計欄位
        matched_rule = {"local": "", "agg": "", "freight": ""}
        if df_rules is not None:
            for _, row in df_rules.iterrows():
                # 使用標題模糊匹配 Excel 裡的 Item_Name
                if fuzz.partial_ratio(str(row.get('Item_Name', '')), title) >= 85:
                    matched_rule = {
                        "local": str(row.get('Unit_Rule_Local', '')),
                        "agg": str(row.get('Unit_Rule_Agg', '')),
                        "freight": str(row.get('Unit_Rule_Freight', ''))
                    }
                    break
        
        # 💡 解開數據字串 ds
        ds = str(item.get("ds", ""))
        data_list = [pair.split(":") for pair in ds.split("|") if ":" in pair]
        if not data_list: continue
        
        ids = [str(e[0]).strip() for e in data_list if len(e) > 0]
        id_counts = Counter(ids)

        # --- 2.1 單項 PC 數核對 (含壞軌相容邏輯) ---
        u_local = matched_rule["local"]
        is_body = "本體" in title
        is_journal = any(k in title for k in ["軸頸", "內孔", "Journal"])
        
        # 判斷是否為重量計件 (KG)
        is_weight_mode = "KG" in title.upper() or target_pc > 100

        if is_weight_mode:
            current_sum = 0
            has_bad_sector = False
            for e in data_list:
                temp_val = safe_float(e[1])
                if temp_val == "BAD_DATA": has_bad_sector = True
                else: current_sum += temp_val
            actual_item_qty = current_sum
            if has_bad_sector:
                accounting_issues.append({
                    "page": page, "item": title, "issue_type": "⚠️數據損毀",
                    "common_reason": "包含無法辨識的重量數據，總重可能不準",
                    "failures": [{"id": "警告", "val": "[!]", "calc": "數據損毀"}]
                })
        else:
            # 數量模式：1SET換算、本體去重、其餘計行
            if "1SET=4PCS" in u_local: actual_item_qty = len(data_list) / 4
            elif "1SET=2PCS" in u_local: actual_item_qty = len(data_list) / 2
            elif is_body or "PC=PC" in u_local: actual_item_qty = len(set(ids)) 
            else: actual_item_qty = len(data_list)

        # 單項數量比對
        if not is_weight_mode and actual_item_qty != target_pc and target_pc > 0:
            accounting_issues.append({
                "page": page, "item": title, "issue_type": "統計不符(單項)",
                "common_reason": f"要求 {target_pc}PC，內文核算為 {actual_item_qty}",
                "failures": [{"id": "標題目標", "val": target_pc}, {"id": "內文實際", "val": actual_item_qty}]
            })

        # --- 2.2 軸頸重複性檢查 ---
        if is_journal:
            for rid, count in id_counts.items():
                if count >= 3:
                    accounting_issues.append({
                        "page": page, "item": title, "issue_type": "🛑編號重複異常",
                        "common_reason": f"編號 {rid} 出現 {count} 次，軸頸限 2 次",
                        "failures": [{"id": rid, "val": count, "calc": "禁止超過2次"}]
                    })

        # --- 2.3 總表對帳 (A聚合/B一般) ---
        u_agg_raw = matched_rule["agg"]
        agg_parts = [p.strip() for p in u_agg_raw.split(",")]
        is_exempt_from_basket = "豁免" in agg_parts
        
        agg_multiplier = 1.0
        for p in agg_parts:
            conv = re.search(r"(\d+)SET=1PC", p)
            if conv: agg_multiplier = 1.0 / float(conv.group(1))

        for s_title, data in global_sum_tracker.items():
            is_rep = any(k in s_title for k in ["ROLL車修", "再生"])
            is_weld = "銲補" in s_title
            is_assem = any(k in s_title for k in ["拆裝", "組裝"])
            
            match = False
            # A 模式 (聚合籃子)
            if (is_rep or is_weld or is_assem) and not is_exempt_from_basket:
                if is_rep and any(k in title for k in ["未再生", "再生", "研磨", "車修"]): match = True
                elif is_weld and "銲補" in title: match = True
                elif is_assem and any(k in title for k in ["拆裝", "組裝", "真圓度"]): match = True
            # B 模式 (名字對帳)
            if not match and fuzz.partial_ratio(s_title.upper(), title.upper()) > 90: match = True

            if match:
                val_for_agg = actual_item_qty * agg_multiplier
                data["actual"] += val_for_agg
                data["details"].append({"id": f"{title} (P.{page})", "val": val_for_agg, "calc": "計入總帳"})

        # --- 2.4 運費核對 (動態解析 XPC=1) ---
        u_fr = matched_rule["freight"]
        if ("計入" in u_fr or (is_body and "未再生" in title)) and "豁免" not in u_fr:
            fr_divisor = 1.0
            fr_match = re.search(r"(\d+)PC=1", u_fr)
            if fr_match: fr_divisor = float(fr_match.group(1))
            
            val_for_fr = actual_item_qty / fr_divisor
            freight_actual_sum += val_for_fr
            freight_details.append({"id": f"{title} (P.{page})", "val": val_for_fr, "calc": "計入運費"})

    # 3. 結算異常報告
    for s_title, data in global_sum_tracker.items():
        if abs(data["actual"] - data["target"]) > 0.01 and data["target"] > 0:
            icon = "🚚" if "運費" in s_title else "🔍"
            accounting_issues.append({
                "page": "總表", "item": s_title, "issue_type": "統計不符",
                "common_reason": f"標註 {data['target']} != 加總 {data['actual']}",
                "failures": [{"id": f"{icon} 基準", "val": data["target"]}] + data["details"] + [{"id": "🧮 總計", "val": data["actual"]}]
            })

    if abs(freight_actual_sum - freight_target) > 0.01 and freight_target > 0:
        accounting_issues.append({
            "page": "總表", "item": "運費核對", "issue_type": "統計不符(運費)",
            "common_reason": f"基準 {freight_target} != 加總 {freight_actual_sum}",
            "failures": [{"id": "🚚 基準", "val": freight_target}] + freight_details + [{"id": "🧮 總計", "val": freight_actual_sum}]
        })
        
    return accounting_issues
    
def python_process_audit(dimension_data):
    """
    Python 流程稽核員：跨頁面檢查每一支編號的尺寸演進是否符合物理規律
    """
    process_issues = []
    roll_history = {} 
    import re
    if not dimension_data: return []

    # 1. 建立「工件履歷資料庫」
    for item in dimension_data:
        p_num = item.get("page", "?")
        ds = str(item.get("ds", ""))
        cat = str(item.get("category", "")).strip()
        title = str(item.get("item_title", ""))
        
        # 解析壓縮字串
        pairs = [p.split(":") for p in ds.split("|") if ":" in p]
        for rid, val_str in pairs:
            # 💡 壞軌偵測：如果數值看不清，不列入位階比對，避免誤判
            if "[!]" in val_str: continue 
            
            try:
                # 提取純數字
                val_match = re.findall(r"\d+\.?\d*", val_str)
                val = float(val_match[0]) if val_match else None
                if val is None: continue
                
                rid_clean = rid.strip()
                if rid_clean not in roll_history: roll_history[rid_clean] = []
                
                # 將這筆紀錄存進該編號的履歷中
                roll_history[rid_clean].append({
                    "process": cat, 
                    "val": val, 
                    "page": p_num, 
                    "title": title
                })
            except: continue

    # 2. 定義物理位階權重 (數字越大代表製程越後段)
    # 權重規則：未再生(1) < 研磨(2) < 再生(3) < 銲補(4)
    weights = {
        "未再生本體": 1, 
        "軸頸未再生": 1, 
        "精加工再生": 3, 
        "銲補": 4
    }

    # 3. 執行「跨製程比對」
    for rid, records in roll_history.items():
        if len(records) < 2: continue # 只有一筆紀錄無法比對
        
        # 按頁碼排序，模擬加工先後順序
        records.sort(key=lambda x: str(x['page']))
        
        for i in range(len(records) - 1):
            curr = records[i] # 前一個製程
            nxt = records[i+1] # 後一個製程
            
            # 💡 [細節校正]：如果標題含「研磨」，位階設為 2
            w_curr = 2 if "研磨" in curr['title'] else weights.get(curr['process'], 3)
            w_nxt = 2 if "研磨" in nxt['title'] else weights.get(nxt['process'], 3)
            
            # 💡 核心判定：如果後一個製程的位階比較高，尺寸「不應」變小
            # (例如：再生車修後的尺寸理論上應大於未再生時的尺寸門檻)
            if w_nxt > w_curr and nxt['val'] < curr['val']:
                process_issues.append({
                    "page": nxt['page'], 
                    "item": f"編號 {rid} 跨製程位階檢查",
                    "issue_type": "🛑流程異常(位階衝突)",
                    "common_reason": f"後段製程尺寸({nxt['val']})小於前段({curr['val']})",
                    "failures": [{
                        "id": rid, 
                        "val": f"後段:{nxt['val']} < 前段:{curr['val']}", 
                        "calc": "不符物理演進邏輯"
                    }],
                    "source": "🐍 系統判定"
                })
    return process_issues
    
# --- 6. 手機版 UI 與 核心執行邏輯 ---
st.title("🏭 交貨單稽核")

data_source = st.radio(
    "請選擇資料來源：", 
    ["📸 上傳照片", "📂 上傳 JSON 檔", "📊 上傳 Excel 檔"], 
    horizontal=True
)

with st.container(border=True):
    # --- 情況 A: 上傳照片 ---
    if data_source == "📸 上傳照片":
        if st.session_state.get('source_mode') == 'json' or st.session_state.get('source_mode') == 'excel':
            st.session_state.photo_gallery = []
            st.session_state.source_mode = 'image'

        uploaded_files = st.file_uploader(
            "請選擇 JPG/PNG 照片...", 
            type=['jpg', 'png', 'jpeg'], 
            accept_multiple_files=True, 
            key=f"uploader_{st.session_state.uploader_key}"
        )
        
        if uploaded_files:
            for f in uploaded_files: 
                if not any(x['file'].name == f.name for x in st.session_state.photo_gallery if x['file']):
                    st.session_state.photo_gallery.append({
                        'file': f, 
                        'table_md': None, 
                        'header_text': None,
                        'full_text': None,
                        'raw_json': None
                    })
            st.session_state.uploader_key += 1
            if st.session_state.enable_auto_analysis:
                st.session_state.auto_start_analysis = True
            components.html("""<script>window.parent.document.body.scrollTo(0, window.parent.document.body.scrollHeight);</script>""", height=0)
            st.rerun()

    # --- 情況 B: 上傳 JSON ---
    elif data_source == "📂 上傳 JSON 檔":
        st.info("💡 請點擊下方按鈕，從你的資料夾選擇之前下載的 `.json` 檔。")
        uploaded_json = st.file_uploader("上傳JSON檔", type=['json'], key="json_uploader")
        
        if uploaded_json:
            try:
                current_file_name = uploaded_json.name
                if st.session_state.get('last_loaded_json_name') != current_file_name:
                    json_data = json.load(uploaded_json)
                    st.session_state.photo_gallery = []
                    st.session_state.source_mode = 'json'
                    st.session_state.last_loaded_json_name = current_file_name
                    
                    import re
                    for page in json_data:
                        real_page = "Unknown"
                        full_text = page.get('full_text', '')
                        if full_text:
                            match = re.search(r"(?:項次|Page|頁次|NO\.)[:\s]*(\d+)\s*[/／]\s*\d+", full_text, re.IGNORECASE)
                            if match:
                                real_page = match.group(1)
                        
                        st.session_state.photo_gallery.append({
                            'file': None,
                            'table_md': page.get('table_md'),
                            'header_text': page.get('header_text'),
                            'full_text': full_text,
                            'raw_json': page.get('raw_json'),
                            'real_page': real_page
                        })
                    
                    st.toast(f"✅ 成功載入 JSON: {current_file_name}", icon="📂")
                    if st.session_state.enable_auto_analysis:
                        st.session_state.auto_start_analysis = True
                    st.rerun()
                else:
                    st.success(f"📂 目前載入 JSON：**{uploaded_json.name}**")
            except Exception as e:
                st.error(f"JSON 檔案格式錯誤: {e}")

    # --- 情況 C: 上傳 Excel (新增的放在這) ---
    elif data_source == "📊 上傳 Excel 檔":
        st.info("💡 上傳 Excel 檔後，系統會將表格內容轉換為文字供 AI 稽核。")
        uploaded_xlsx = st.file_uploader("上傳 Excel 檔", type=['xlsx', 'xls'], key="xlsx_uploader")
        
        if uploaded_xlsx:
            try:
                current_file_name = uploaded_xlsx.name
                if st.session_state.get('last_loaded_xlsx_name') != current_file_name:
                    df_dict = pd.read_excel(uploaded_xlsx, sheet_name=None)
                    st.session_state.photo_gallery = []
                    st.session_state.source_mode = 'excel'
                    st.session_state.last_loaded_xlsx_name = current_file_name
                    
                    for sheet_name, df in df_dict.items():
                        df = df.fillna("")
                        md_table = df.to_markdown(index=False)
                        st.session_state.photo_gallery.append({
                            'file': None,
                            'table_md': md_table,
                            'header_text': f"來源分頁: {sheet_name}",
                            'full_text': f"Excel 內容 - 分頁 {sheet_name}\n" + md_table,
                            'raw_json': None,
                            'real_page': sheet_name
                        })
                    st.toast(f"✅ 成功載入 Excel: {current_file_name}", icon="📊")
                    if st.session_state.enable_auto_analysis:
                        st.session_state.auto_start_analysis = True
                    st.rerun()
                else:
                    st.success(f"📊 目前載入 Excel：**{uploaded_xlsx.name}**")
            except Exception as e:
                st.error(f"Excel 讀取失敗: {e}")

if st.session_state.photo_gallery:
    st.caption(f"已累積 {len(st.session_state.photo_gallery)} 頁文件")
    col_btn1, col_btn2 = st.columns([1, 1], gap="small")
    with col_btn1: start_btn = st.button("🚀 開始分析", type="primary", use_container_width=True)
    with col_btn2: 
        clear_btn = st.button("🗑️照片清除", help="清除", use_container_width=True)

    if clear_btn:
        st.session_state.photo_gallery = []
        st.session_state.analysis_result_cache = None
        if 'last_loaded_json_name' in st.session_state:
            del st.session_state.last_loaded_json_name 
        st.rerun()

    is_auto_start = st.session_state.auto_start_analysis
    if is_auto_start:
        st.session_state.auto_start_analysis = False

    if 'analysis_result_cache' not in st.session_state:
        st.session_state.analysis_result_cache = None

    trigger_analysis = start_btn or is_auto_start

    if trigger_analysis:
        st.session_state.auto_start_analysis = False
        total_start = time.time()
        
        # 1. 執行分析區塊
        with st.status("總稽核官正在進行全方位分析...", expanded=True) as status_box:
            status_text = st.empty()
            progress_bar = st.progress(0)
            total_imgs = len(st.session_state.photo_gallery)
            ocr_start = time.time()
            
            def process_task(index, item):
                if item.get('full_text'):
                    return index, item.get('header_text',''), item['full_text'], None
                try:
                    item['file'].seek(0)
                    _, h, f, _, _ = extract_layout_with_azure(item['file'], DOC_ENDPOINT, DOC_KEY)
                    return index, h, f, None
                except Exception as e:
                    return index, None, None, str(e)

            # 數據收集
            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                futures = [executor.submit(process_task, i, item) for i, item in enumerate(st.session_state.photo_gallery)]
                for future in concurrent.futures.as_completed(futures):
                    idx, h_txt, f_txt, err = future.result()
                    if not err:
                        st.session_state.photo_gallery[idx].update({'header_text': h_txt, 'full_text': f_txt, 'file': None})
                    progress_bar.progress((idx + 1) / total_imgs)

            ocr_duration = time.time() - ocr_start
            combined_input = ""
            for i, p in enumerate(st.session_state.photo_gallery):
                combined_input += f"\n=== Page {i+1} ===\n{p.get('full_text','')}\n"

            res_main = agent_unified_check(combined_input, combined_input, GEMINI_KEY, main_model_name)
            st.write("DEBUG - AI 回傳內容:", res_main) # ⚡️ 讓錯誤現形
            dim_data = res_main.get("dimension_data", [])
            python_numeric_issues = python_numerical_audit(dim_data)
            python_accounting_issues = python_accounting_audit(dim_data, res_main)
            python_process_issues = python_process_audit(dim_data)
            python_header_issues, python_debug_data = python_header_check(st.session_state.photo_gallery)

            ai_filtered_issues = []
            ai_raw_issues = res_main.get("issues", [])
            if isinstance(ai_raw_issues, list):
                for i in ai_raw_issues:
                    if isinstance(i, dict):
                        i['source'] = '🤖 總稽核 AI'
                        if any(k in i.get("issue_type", "") for k in ["流程", "規格提取失敗", "未匹配"]):
                            ai_filtered_issues.append(i)

            # 最終合併所有籃子
            all_issues = ai_filtered_issues + python_numeric_issues + python_accounting_issues + python_process_issues + python_header_issues
            
            # 存入快取
            usage = res_main.get("_token_usage", {"input": 0, "output": 0})
            st.session_state.analysis_result_cache = {
                "job_no": res_main.get("job_no", "Unknown"),
                "all_issues": all_issues,
                "total_duration": time.time() - total_start,
                "cost_twd": (usage.get("input", 0)*0.5 + usage.get("output", 0)*3.0) / 1000000 * 32.5,
                "total_in": usage.get("input", 0),
                "total_out": usage.get("output", 0),
                "ocr_duration": ocr_duration,
                "time_eng": time.time() - total_start - ocr_duration,
                "ai_extracted_data": dim_data,
                "python_debug_data": python_debug_data,
                "full_text_for_search": combined_input, # 補回這行以免報錯
                "combined_input": combined_input  # ✅ 確保這一行一定要在！
            }
            status_box.update(label="✅ 分析完成！", state="complete", expanded=False)
            st.rerun()

    # --- 💡 [重大修正] 顯示結果區塊：必須與 if trigger_analysis 平級 ---
    if st.session_state.analysis_result_cache:
        cache = st.session_state.analysis_result_cache
        all_issues = cache.get('all_issues', [])
        
        st.success(f"工令: {cache['job_no']} | ⏱️ {cache['total_duration']:.1f}s")
        st.info(f"💰 本次成本: NT$ {cache['cost_twd']:.2f} (In: {cache['total_in']:,} / Out: {cache['total_out']:,})")
        st.caption(f"細節耗時: Azure OCR {cache['ocr_duration']:.1f}s | AI 分析 {cache['time_eng']:.1f}s")
        
        # 展開頁面
        with st.expander("🔍 查看 AI 讀取到的 Excel 規則 (Debug)"):
            rules_text = get_dynamic_rules(cache.get('full_text_for_search',''), debug_mode=True)
            st.markdown(rules_text)
                
        with st.expander("🔬 查看 AI 抄錄原始數據", expanded=False):
            st.json(cache.get("ai_extracted_data", []))

        with st.expander("🐍 查看 Python 硬邏輯偵測結果 (Debug)", expanded=False):
            if cache.get('python_debug_data'):
                st.dataframe(cache['python_debug_data'], use_container_width=True, hide_index=True)
            else:
                st.caption("無偵測資料")

        # 判定結論顯示
        real_errors = [i for i in all_issues if "未匹配" not in i.get('issue_type', '')]
        if not all_issues:
            st.balloons()
            st.success("✅ 全數合格！")
        elif not real_errors:
            st.success(f"✅ 數值合格！ (但有 {len(all_issues)} 個項目未匹配規則)")
        else:
            st.error(f"發現 {len(real_errors)} 類異常")

        # 卡片循環顯示
        for item in all_issues:
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                source_label = item.get('source', '')
                issue_type = item.get('issue_type', '異常')
                c1.markdown(f"**P.{item.get('page', '?')} | {item.get('item')}**  `{source_label}`")
                
                if any(kw in issue_type for kw in ["統計", "數量", "流程"]):
                    c2.error(f"🛑 {issue_type}")
                else:
                    c2.warning(f"⚠️ {issue_type}")
                
                st.caption(f"原因: {item.get('common_reason', '')}")
                
                failures = item.get('failures', [])
                if failures:
                    table_data = []
                    for f in failures:
                        if isinstance(f, dict):
                            table_data.append({
                                "項目/編號": f.get('id', '未知'), 
                                "實測/計數": f.get('val', 'N/A'),
                                "標準/備註": f.get('target', ''),
                                "狀態": f.get('calc', '')
                            })
                    st.dataframe(table_data, use_container_width=True, hide_index=True)
        
        st.divider()
        # 下載按鈕與原文展開
        # ... (這裡接你原本剩下的代碼即可，也要記得縮排往左移)
        current_job_no = cache.get('job_no', 'Unknown')
        safe_job_no = current_job_no.replace("/", "_").replace("\\", "_").strip()
        file_name_str = f"{safe_job_no}_cleaned.json"

        # 準備匯出資料
        export_data = []
        for item in st.session_state.photo_gallery:
            export_data.append({
                "table_md": item.get('table_md'),
                "header_text": item.get('header_text'),
                "full_text": item.get('full_text'),
                "raw_json": item.get('raw_json')
            })
        json_str = json.dumps(export_data, indent=2, ensure_ascii=False)

        st.subheader("💾 測試資料存檔")
        st.caption(f"已識別工令：**{current_job_no}**。下載後可供下次測試使用。")
        
        st.download_button(
            label=f"⬇️ 下載測試資料 ({file_name_str})",
            data=json_str,
            file_name=file_name_str,
            mime="application/json",
            type="primary"
        )

        # 💡 使用 .get() 可以防止因為找不到標籤而直接報錯當機
        with st.expander("👀 查看傳給 AI 的最終文字 (Prompt Input)"):
            st.caption("這才是 AI 真正讀到的內容 (已過濾雜訊)：")
            st.code(cache.get('combined_input', '無資料'), language='markdown')
    
    if st.session_state.photo_gallery and st.session_state.get('source_mode') != 'json':
        st.caption("已拍攝照片：")
        cols = st.columns(4)
        for idx, item in enumerate(st.session_state.photo_gallery):
            with cols[idx % 4]:
                if item.get('file'):
                    st.image(item['file'], caption=f"P.{idx+1}", use_container_width=True)
                if st.button("❌", key=f"del_{idx}"):
                    st.session_state.photo_gallery.pop(idx)
                    st.session_state.analysis_result_cache = None
                    st.rerun()
else:
    st.info("👆 請點擊上方按鈕開始新增照片")
