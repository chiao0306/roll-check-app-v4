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
        "Gemini 2.5 Flash Lite": "gemini-2.5-flash-lite",
        "Gemini 2.5 Pro": "models/gemini-2.5-pro",
        #"GPT-5(無效)": "models/gpt-5",
        #"GPT-5 Mini(無效)": "models/gpt-5-mini",
    }
    options_list = list(model_options.keys())
    
    st.subheader("🤖 總稽核 Agent")
    model_selection = st.selectbox(
        "負責：規格、製程、數量、統計全包", 
        options=options_list, 
        index=1, 
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

def agent_unified_check(combined_input, full_text_for_search, api_key, model_name):
    import re
    # 讀取 Excel 規則 (供 Python 後端查表使用)
    dynamic_rules = get_dynamic_rules(full_text_for_search)

    # 1. 整合您的【工程級精密 Prompt】 - 🔇 靜音版 (移除 AI 判斷功能)
    system_prompt = f"""
    你是一位極度嚴謹的中鋼機械品管【數據抄錄員】。你必須像「電腦程式」一樣執行任務。
    
    {dynamic_rules}

    ---

    #### ⚔️ 模組 A：工程尺寸數據提取 (AI 任務：純抄錄)
    1. **規格抄錄 (std_spec)**：精確抄錄標題中含 `mm`、`±`、`+`、`-` 的原始文字。
    2. **數據抄錄 (ds)**：格式為 `"ID:值|ID:值"`。
       - **⚠️ 絕對完整原則 (Anti-Deduplication)**：表格裡有幾行數據，就必須輸出幾組 `ID:值`。
       - **🚫 嚴禁合併重複 ID**：一支輥輪通常有兩個軸頸，若表格顯示兩次 `Y5612001`，你必須輸出兩次！
         - 錯誤範例：`"Y5612001:98"` (只寫一次)
         - 正確範例：`"Y5612001:98|Y5612001:98"` (完整保留)
       - **字串保護**：禁止簡化數字。`349.90` 必寫 `"349.90"`。
       - **壞軌標記 [!]**：若儲存格辨識不良（汙點/字跡黏連/反光），嚴禁猜測，直接標記為 `[!]`。
    
    3. **項目分類決策流程 (由上至下執行，命中即停止)**：
        - **LEVEL 1：銲補與裝配判定 (最高優先)**
          * 標題含「銲補」、「銲接」 -> `min_limit`。
          * 標題含「組裝」、「拆裝」、「裝配」、「真圓度」 -> `range`。
        - **LEVEL 2：未再生判定 (含車修)**
          * 標題含「未再生」三字時：
            a. 含「軸頸」 -> `max_limit`。
            b. 不含「軸頸」(本體) -> `un_regen`。
          * (💡 注意：此類項目即便包含「車修」字眼，也必須鎖定在 LEVEL 2)。
        - **LEVEL 3：精加工判定**
          * 標題不含「未再生」，且包含「再生」、「研磨」、「精加工」、「車修加工」、「KEYWAY」 -> `range`。

    #### 💰 模組 B：會計指標提取 (AI 任務：純抄錄)
    1. **統計表**：抄錄左上角統計表每一行名稱與實交數量到 `summary_rows`。
    2. **指標提取**：提取運費項次與標題括號內的 PC 數。

    ---
    #### 📝 輸出規範 (極簡 JSON Format)
    必須回傳單一合法 JSON。
    ⚠️ 絕對禁止回傳 accounting_rules, sl 以及 issues 欄位。
    
    格式如下：
    {{
      "job_no": "工令",
      "summary_rows": [ {{ "title": "名稱", "target": 數字 }} ],
      "freight_target": 數字,
      "dimension_data": [
         {{
           "page": 數字, "item_title": "標題", "category": "分類名稱", 
           "item_pc_target": 數字, "std_spec": "規格文字", "ds": "ID:值|ID:值" 
         }}
      ]
    }}
    """
    
    try:
        genai.configure(api_key=api_key)
        
        model = genai.GenerativeModel(
            model_name=model_name,
            generation_config={
                "response_mime_type": "application/json",
                "temperature": 0.0,
                "max_output_tokens": 16384
            }
        )
        
        with st.spinner('🤖 總稽核 Agent 正在進行數據轉錄 (強制完整模式)...'):
            response = model.generate_content([system_prompt, combined_input])
        
        raw_content = response.text.strip()
        
        # 🛡️ 強化解析
        json_match = re.search(r"\{.*\}", raw_content, re.DOTALL)
        if json_match:
            raw_content = json_match.group()
            
        parsed_data = json.loads(raw_content)
        
        # 記錄消耗 Token
        parsed_data["_token_usage"] = {
            "input": response.usage_metadata.prompt_token_count, 
            "output": response.usage_metadata.candidates_token_count
        }
        return parsed_data

    except json.JSONDecodeError as e:
        st.error(f"❌ JSON 解析失敗！")
        with st.expander("👀 查看導致錯誤的 AI 原始回應"):
            if 'raw_content' in locals():
                st.code(raw_content)
            elif 'response' in locals():
                st.code(response.text)
        return {"job_no": "JSON Error", "issues": [], "dimension_data": []}

    except Exception as e:
        st.error(f"❌ 系統發生錯誤: {str(e)}")
        return {"job_no": f"Error: {str(e)}", "issues": [], "dimension_data": []}

# --- 重點：Python 引擎獨立於 agent 函式之外 ---

def python_numerical_audit(dimension_data):
    grouped_errors = {}
    import re
    if not dimension_data: return []

    for item in dimension_data:
        # 1. 取得數據
        ds = str(item.get("ds", ""))
        if not ds: continue
        raw_entries = [p.split(":") for p in ds.split("|") if ":" in p]
        
        # 🧽 強制清洗標題與分類
        title = str(item.get("item_title", "")).replace(" ", "").replace("\n", "").replace('"', "")
        cat = str(item.get("category", "")).replace(" ", "").strip()
        
        page_num = item.get("page", "?")
        raw_spec = str(item.get("std_spec", "")).replace('"', "")
        
        # 2. 🛡️ 數據清洗
        all_nums = [float(n) for n in re.findall(r"[-+]?\d+\.?\d*", raw_spec.replace(" ", ""))]
        noise = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 0.0] 
        clean_std = [n for n in all_nums if (n not in noise and n > 10)]

        # 3. 💡 多重區間自動預算
        s_ranges = []
        spec_parts = re.split(r"[一二三四五六]|[;；]", raw_spec)
        
        for part in spec_parts:
            # ⚡️ 修正點：移除 "mm" 與 "MM"，讓 "135mm~129mm" 變成 "135~129"
            clean_part = part.replace(" ", "").replace("\n", "").replace("mm", "").replace("MM", "").strip()
            if not clean_part: continue
            
            # 邏輯 A：優先處理 ± (如 300±0.1)
            pm_match = re.search(r"(\d+\.?\d*)?±(\d+\.?\d*)", clean_part)
            if pm_match:
                b = float(pm_match.group(1)) if pm_match.group(1) else 0.0
                o = float(pm_match.group(2))
                s_ranges.append([round(b - o, 4), round(b + o, 4)])
                continue

            # 邏輯 B：處理波浪號區間 (如 135~129)
            # 現在移除了 mm，這裡就能成功抓到 [129, 135] 了！
            tilde_match = re.search(r"(\d+\.?\d*)[~～-](\d+\.?\d*)", clean_part)
            if tilde_match:
                n1, n2 = float(tilde_match.group(1)), float(tilde_match.group(2))
                # 防呆：避免把 160-0.01 (公差) 誤判為 160~0.01 (區間)
                if abs(n1 - n2) < n1 * 0.5: 
                    s_ranges.append([round(min(n1, n2), 4), round(max(n1, n2), 4)])
                    continue

            # 邏輯 C：智慧配對 (解決 140 -0.01, -0.03)
            all_tokens = re.findall(r"[-+]?\d+\.?\d*", clean_part)
            if not all_tokens: continue

            bases = []
            offsets = []
            for token in all_tokens:
                val = float(token)
                if val > 10.0: bases.append(val)
                elif abs(val) < 10.0: offsets.append(val)
            
            if bases:
                for b in bases:
                    if offsets:
                        endpoints = [round(b + o, 4) for o in offsets]
                        if len(endpoints) == 1: endpoints.append(b)
                        s_ranges.append([min(endpoints), max(endpoints)])
                    else:
                        s_ranges.append([b, b])

        # 4. 💡 預算基準
        s_threshold = 0
        un_regen_target = None
        if cat in ["un_regen", "未再生"] or ("未再生" in (cat + title) and "軸頸" not in (cat + title)):
            cands = [n for n in clean_std if n >= 120.0]
            if cands: un_regen_target = max(cands)

        # --- 5. 開始逐一判定 ---
        for entry in raw_entries:
            if len(entry) < 2: continue
            rid = str(entry[0]).strip().replace(" ", "")
            val_raw = str(entry[1]).strip().replace(" ", "")
            
            if not val_raw or val_raw in ["N/A", "nan", "M10"]: continue

            try:
                is_passed, reason, t_used, engine_label = True, "", "N/A", "未知"

                if "[!]" in val_raw:
                    is_passed = False
                    reason = "🛑數據損壞(壞軌)"
                    val_str = "[!]"
                    val = -999.0
                else:
                    v_m = re.findall(r"\d+\.?\d*", val_raw)
                    val_str = v_m[0] if v_m else val_raw
                    val = float(val_str)

                if val_str != "[!]":
                    is_two_dec = "." in val_str and len(val_str.split(".")[-1]) == 2
                    is_pure_int = "." not in val_str
                else:
                    is_two_dec, is_pure_int = True, True 

                if "min_limit" in cat or "銲補" in (cat + title):
                    engine_label = "銲補"
                    if not is_pure_int: is_passed, reason = False, "應為純整數"
                    elif clean_std:
                        t_used = min(clean_std, key=lambda x: abs(x - val))
                        if val < t_used: is_passed, reason = False, "數值不足"
                
                elif un_regen_target is not None:
                    engine_label = "未再生"
                    t_used = un_regen_target
                    if val <= t_used:
                        if not is_pure_int: is_passed, reason = False, "應為整數"
                    elif not is_two_dec: 
                        is_passed, reason = False, "應填兩位小數"

                elif "max_limit" in cat or (("軸頸" in (cat + title)) and ("未再生" in (cat + title))):
                    engine_label = "軸頸(上限)"
                    candidates = clean_std
                    target = max(candidates) if candidates else 0
                    t_used = target
                    if target > 0:
                        if not is_pure_int: is_passed, reason = False, "應為純整數"
                        elif val > target: is_passed, reason = False, f"超過上限 {target}"

                elif any(x in (cat + title) for x in ["再生", "精加工", "研磨", "車修", "組裝", "拆裝", "真圓度"]) and "未再生" not in (cat + title):
                    engine_label = "精加工"
                    if not is_two_dec:
                        is_passed, reason = False, "應填兩位小數"
                    elif s_ranges:
                        t_used = str(s_ranges)
                        # 💡 核心：只要符合任何一個解析出的區間就算合格
                        if not any(r[0] <= val <= r[1] for r in s_ranges): 
                            is_passed, reason = False, "不在區間內"

                if not is_passed:
                    key = (page_num, title, reason)
                    if key not in grouped_errors:
                        grouped_errors[key] = {
                            "page": page_num, "item": title, 
                            "issue_type": f"異常({engine_label})", 
                            "common_reason": reason, "failures": [],
                            "source": "🐍 工程引擎"
                        }
                    grouped_errors[key]["failures"].append({"id": rid, "val": val_str, "target": f"基準:{t_used}"})
            except: continue
                
    return list(grouped_errors.values())

def python_accounting_audit(dimension_data, res_main):
    """
    Python 會計官：運費邏輯全面接管版
    1. 修正：不再依賴 freight_target > 0 開關，強制計算每筆項目的運費值。
    2. 注入：若總表籃子名稱含「運費」，直接使用計算出的運費值，不再進行模糊比對。
    """
    accounting_issues = []
    from thefuzz import fuzz
    from collections import Counter
    import re
    import pandas as pd 

    # 🧽 真空清洗工具
    def clean_text(text):
        return str(text).replace(" ", "").replace("\n", "").replace("\r", "").replace('"', '').replace("'", "").strip()

    # 安全轉型工具
    def safe_float(value):
        if value is None or str(value).upper() == 'NULL': return 0.0
        if "[!]" in str(value): return "BAD_DATA" 
        cleaned = "".join(re.findall(r"[\d\.]+", str(value).replace(',', '')))
        try: return float(cleaned) if cleaned else 0.0
        except: return 0.0

    # 0. 預載 Excel 規則
    rules_dict = {}
    try:
        df = pd.read_excel("rules.xlsx")
        df.columns = [c.strip() for c in df.columns]
        for _, row in df.iterrows():
            iname = str(row.get('Item_Name', '')).strip()
            u_fr = str(row.get('Unit_Rule_Freight', '')).strip()
            if iname: rules_dict[clean_text(iname)] = u_fr
    except:
        pass 

    # 1. 取得對帳基準
    summary_rows = res_main.get("summary_rows", [])
    global_sum_tracker = {
        s['title']: {"target": safe_float(s['target']), "actual": 0, "details": []} 
        for s in summary_rows if s.get('title')
    }
    
    freight_target = safe_float(res_main.get("freight_target", 0))
    freight_actual_sum = 0
    freight_details = []

    # 2. 逐項過帳
    for item in dimension_data:
        raw_title = item.get("item_title", "")
        title_clean = clean_text(raw_title) 
        page = item.get("page", "?")
        target_pc = safe_float(item.get("item_pc_target", 0)) 
        
        ds = str(item.get("ds", ""))
        data_list = [pair.split(":") for pair in ds.split("|") if ":" in pair]
        if not data_list: continue
        
        ids = [str(e[0]).strip() for e in data_list if len(e) > 0]
        id_counts = Counter(ids)

        # --- 2.1 單項數量計算 ---
        is_weight_mode = "KG" in title_clean.upper() or target_pc > 100
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
                    "page": page, "item": raw_title, "issue_type": "⚠️數據損毀",
                    "common_reason": "含無法辨識重量",
                    "failures": [{"id": "警告", "val": "[!]", "calc": "數據損毀"}],
                    "source": "🐍 會計引擎"
                })
        else:
            actual_item_qty = len(data_list) 

        if actual_item_qty != target_pc and target_pc > 0:
            accounting_issues.append({
                "page": page, "item": raw_title, "issue_type": "統計不符(單項)",
                "common_reason": f"標題 {target_pc}PC != 內文 {actual_item_qty}",
                "failures": [
                    {"id": "目標", "val": target_pc, "calc": "標題"},
                    {"id": "實際", "val": actual_item_qty, "calc": "內文計數"}
                ],
                "source": "🐍 會計引擎"
            })

        # --- 2.2 編號重複性示警 ---
        if "本體" in title_clean:
             for rid, count in id_counts.items():
                if count > 1:
                     accounting_issues.append({
                        "page": page, "item": raw_title, "issue_type": "⚠️編號重複警示(本體)",
                        "common_reason": f"本體編號 {rid} 重複 {count} 次",
                        "failures": [{"id": rid, "val": count, "calc": "建議檢查"}],
                        "source": "🐍 會計引擎"
                     })
        elif any(k in title_clean for k in ["軸頸", "內孔", "JOURNAL"]):
             for rid, count in id_counts.items():
                if count > 2:
                     accounting_issues.append({
                        "page": page, "item": raw_title, "issue_type": "⚠️編號重複警示(軸頸)",
                        "common_reason": f"軸頸編號 {rid} 出現 {count} 次",
                        "failures": [{"id": rid, "val": count, "calc": "建議檢查"}],
                        "source": "🐍 會計引擎"
                     })

        # --- ⚡️ 插入：預先計算此項目的「智慧運費值」 ---
        # 即使 freight_target 為 0，我們也要算，因為總表籃子可能會用到
        
        # Step A: 查找 Excel 規則
        u_fr = rules_dict.get(title_clean, "")
        if not u_fr and rules_dict:
            best_score = 0
            for k, v in rules_dict.items():
                score = fuzz.ratio(k, title_clean)
                if score > 95 and score > best_score:
                    best_score = score
                    u_fr = v
        
        # Step B: 判斷是否計入
        is_exempt = "豁免" in str(u_fr)
        conv_match = re.search(r"(\d+)\s*(?:PC|SET|PCS)?\s*=\s*1", str(u_fr), re.IGNORECASE)
        # 預設底線：全卷「本體」且「未再生」
        is_default_target = "本體" in title_clean and "未再生" in title_clean

        freight_val_for_item = 0.0
        freight_note = ""

        if is_exempt:
            freight_val_for_item = 0.0
        elif conv_match:
            divisor = float(conv_match.group(1))
            freight_val_for_item = actual_item_qty / divisor
            freight_note = f"計入 (/{int(divisor)})"
        elif is_default_target:
            freight_val_for_item = actual_item_qty
            freight_note = "計入運費"
            
        # 累積到獨立變數 (如果有用到的話)
        if freight_val_for_item > 0:
            freight_actual_sum += freight_val_for_item
            freight_details.append({"id": f"{raw_title}", "val": freight_val_for_item, "calc": freight_note})

        # --- 2.3 總表對帳 (含運費注入邏輯) ---
        for s_title, data in global_sum_tracker.items():
            match = False
            s_title_clean = clean_text(s_title)
            
            # 💡 檢查：這是不是一個「運費籃子」？
            is_freight_basket = "運費" in s_title_clean
            
            if is_freight_basket:
                # ⭐️ 運費籃子專用通道：直接注入剛剛算好的 freight_val_for_item
                if freight_val_for_item > 0:
                    data["actual"] += freight_val_for_item
                    data["details"].append({"id": f"{raw_title} (P.{page})", "val": freight_val_for_item, "calc": freight_note})
                continue # 處理完直接換下一個籃子，不走下面的模糊比對
            
            # === 以下為非運費籃子的常規邏輯 ===
            
            # 門禁特徵
            req_body = "本體" in s_title_clean
            req_journal = any(k in s_title_clean for k in ["軸頸", "內孔", "JOURNAL"])
            req_unregen = "未再生" in s_title_clean
            req_regen_only = "再生" in s_title_clean and not req_unregen
            
            # 項目特徵
            is_item_body = "本體" in title_clean
            is_item_journal = any(k in title_clean for k in ["軸頸", "內孔", "JOURNAL"])
            is_item_unregen = "未再生" in title_clean
            
            # 優先級一：三大天王
            is_main_disassembly = "ROLL拆裝" in s_title_clean 
            is_main_machining = "ROLL車修" in s_title_clean   
            is_main_welding = "ROLL銲補" in s_title_clean     

            if is_main_disassembly:
                if "組裝" in title_clean or "拆裝" in title_clean: match = True
            elif is_main_machining:
                has_part = "軸頸" in title_clean or "本體" in title_clean
                has_action = "再生" in title_clean or "未再生" in title_clean
                if has_part and has_action: match = True
            elif is_main_welding:
                has_part = "軸頸" in title_clean or "本體" in title_clean
                if has_part and "銲補" in title_clean: match = True
            else:
                # 優先級二：普通籃子
                if fuzz.partial_ratio(s_title_clean, title_clean) > 98:
                    match = True
                    if req_body and not is_item_body: match = False
                    elif req_journal and not is_item_journal: match = False
                    if req_unregen and not is_item_unregen: match = False
                    elif req_regen_only and is_item_unregen: match = False

            if match:
                data["actual"] += actual_item_qty
                data["details"].append({"id": f"{raw_title} (P.{page})", "val": actual_item_qty, "calc": "計入"})

    # 3. 結算異常
    for s_title, data in global_sum_tracker.items():
        if abs(data["actual"] - data["target"]) > 0.01 and data["target"] > 0:
            accounting_issues.append({
                "page": "總表", "item": s_title, "issue_type": "統計不符(總帳)",
                "common_reason": f"標註 {data['target']} != 實際 {data['actual']}",
                "failures": [{"id": "🔍 基準", "val": data["target"]}] + data["details"] + [{"id": "🧮 實際", "val": data["actual"]}],
                "source": "🐍 會計引擎"
            })

    # 運費獨立檢查 (如果 AI 有抓到獨立變數的話，也檢查一下)
    if abs(freight_actual_sum - freight_target) > 0.01 and freight_target > 0:
        accounting_issues.append({
            "page": "總表", "item": "運費核對(獨立)", "issue_type": "統計不符(運費)",
            "common_reason": f"基準 {freight_target} != 實際 {freight_actual_sum}",
            "failures": [{"id": "🚚 基準", "val": freight_target}] + freight_details + [{"id": "🧮 實際", "val": freight_actual_sum}],
            "source": "🐍 會計引擎"
        })
        
    return accounting_issues
    
def python_process_audit(dimension_data):
    process_issues = []
    roll_history = {} # { "ID": [{"p": "cat", "v": 190, "page": 1}, ...] }
    import re
    if not dimension_data: return []

    for item in dimension_data:
        p_num = item.get("page", "?")
        ds = str(item.get("ds", ""))
        cat = str(item.get("category", "")).strip()
        
        # 1. 先用 | 切分不同數據
        raw_segments = ds.split("|")
        
        for seg in raw_segments:
            # 2. 基本過濾：必須包含冒號
            if ":" not in seg: continue
            
            # 3. 🛡️ 安全切分：防止 "ID:值:備註" 這種多冒號導致崩潰
            parts = seg.split(":")
            
            # 如果切出來少於 2 段 (例如 "ID:")，跳過
            if len(parts) < 2: continue
            
            # 強制只取前兩段，無視後面多餘的冒號
            rid = str(parts[0]).strip()
            val_str = str(parts[1]).strip()
            
            try:
                # 簡單清洗取出數字
                # 這裡加個保護，萬一 val_str 裡沒有數字 (例如 "N/A") 也不要報錯
                found_nums = re.findall(r"\d+\.?\d*", val_str)
                if not found_nums: continue
                
                val = float(found_nums[0])
                
                if rid not in roll_history: roll_history[rid] = []
                roll_history[rid].append({
                    "p": cat, 
                    "v": val, 
                    "page": p_num, 
                    "title": item.get("item_title", "")
                })
            except: 
                continue

    # --- 流程邏輯判定 ---
    weights = {"un_regen": 1, "max_limit": 1, "range": 3, "min_limit": 4}
    
    for rid, records in roll_history.items():
        if len(records) < 2: continue
        
        # 依照頁碼排序
        records.sort(key=lambda x: str(x['page']))
        
        for i in range(len(records) - 1):
            curr, nxt = records[i], records[i+1]
            
            # 取得權重 (預設 2)
            w_curr = weights.get(curr['p'], 2)
            if "研磨" in str(curr['title']): w_curr = 2
            
            w_nxt = weights.get(nxt['p'], 2)
            if "研磨" in str(nxt['title']): w_nxt = 2
            
            # 💡 關鍵判定：後段位階大(如銲補)，數值就不應該變小
            # 例如：先「車修(1)」後「銲補(4)」，尺寸變小是合理的 (車掉一層) -> Pass
            # 例如：先「銲補(4)」後「車修(1)」，尺寸變小是合理的 -> Pass
            # 等等... 這裡的邏輯是「位階檢查」，您的原意應該是：
            # 如果從「低位階」(如車修) 到了 「高位階」(如精加工)，理論上是把東西做小了？
            # 或者是檢查「不合邏輯的尺寸跳變」？
            # 依照原程式碼邏輯保留：
            
            if w_nxt > w_curr and nxt['v'] < curr['v']:
                process_issues.append({
                    "page": nxt['page'], "item": f"編號 {rid} 尺寸位階檢查",
                    "issue_type": "🛑流程異常(尺寸倒置)",
                    "common_reason": f"後段{nxt['p']}尺寸小於前段{curr['p']}",
                    "failures": [{"id": rid, "val": f"後:{nxt['v']} < 前:{curr['v']}", "calc": "尺寸不符位階邏輯"}],
                    "source": "🐍 流程引擎"
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
        # ⚡️ 新增這行：強制清除上一筆的結果
        st.session_state.analysis_result_cache = None 
        
        st.session_state.auto_start_analysis = False
        total_start = time.time()
        
        with st.status("總稽核官正在進行全方位分析...", expanded=True) as status_box:
            progress_bar = st.progress(0)
            
            # 1. OCR (這段保留，速度很快)
            status_box.write("👀 正在進行 OCR 文字識別...")
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

            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                futures = [executor.submit(process_task, i, item) for i, item in enumerate(st.session_state.photo_gallery)]
                for future in concurrent.futures.as_completed(futures):
                    idx, h_txt, f_txt, err = future.result()
                    if not err:
                        st.session_state.photo_gallery[idx].update({'header_text': h_txt, 'full_text': f_txt, 'file': None})
                    progress_bar.progress(0.4 * ((idx + 1) / len(st.session_state.photo_gallery)))

            ocr_duration = time.time() - ocr_start
            
            # 2. 組合所有文字 (關鍵：一次丟進去)
            combined_input = ""
            for i, p in enumerate(st.session_state.photo_gallery):
                combined_input += f"\n=== Page {i+1} ===\n{p.get('full_text','')}\n"

            # 3. 呼叫 AI (這裡只會跑一次，約 20-30 秒)
            status_box.write("🤖 AI 正在全卷分析...")
            res_main = agent_unified_check(combined_input, combined_input, GEMINI_KEY, main_model_name)
            progress_bar.progress(0.8)
            
            # 4. Python 邏輯檢查
            status_box.write("🐍 Python 正在進行邏輯比對...")
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
                        if not any(k in i.get("issue_type", "") for k in ["流程", "規格提取失敗", "未匹配"]):
                            ai_filtered_issues.append(i)

            all_issues = ai_filtered_issues + python_numeric_issues + python_accounting_issues + python_process_issues + python_header_issues
            
            # 5. 存檔與完成
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
                "full_text_for_search": combined_input,
                "combined_input": combined_input
            }
            
            progress_bar.progress(1.0)
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
