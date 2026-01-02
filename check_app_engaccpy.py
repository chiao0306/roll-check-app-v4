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
    你是一位極度嚴謹的中鋼機械品管【數據抄錄員】。你必須像「電腦程式」一樣執行任務。
    
    {dynamic_rules}

    ---

    #### ⚔️ 模組 A：工程尺寸數據提取 (AI 任務：抄錄)
    1. **規格抄錄 (std_spec)**：精確抄錄標題中含 `mm`、`±`、`+`、`-` 的原始文字。
    2. **數據抄錄 (ds)**：格式 `"ID:值|ID:值"`。禁止簡化，`349.90` 必寫 `"349.90"`。
    3. **項目分類決策流程 (由上至下執行，命中即停止)**：
        - **LEVEL 1：銲補與裝配判定 (最高優先)**
          * 標題含「銲補」、「銲接」 -> `min_limit`。
          * 標題含「組裝」、「拆裝」、「裝配」、「真圓度」 -> `range`。
         
        - **LEVEL 2：未再生判定 (含車修)**
          * 標題含「未再生」三字時：
            a. 含「軸頸」 -> `max_limit`。
            b. 不含「軸頸」(本體) -> `un_regen`。
          * (💡 注意：此類項目即使包含「車修」字眼，也必須鎖定在 LEVEL 2，嚴禁進入下一個層級)。
         
        - **LEVEL 3：精加工判定**
          * 標題不含「未再生」，且包含「再生」、「研磨」、「精加工」、「車修加工」、「KEYWAY」 -> `range`。

    4. **數據抄錄 (字串保護模式)**：
       - **禁止簡化**：實測值若顯示 `349.90`，必須輸出 `"349.90"`。
       - **格式**：所有實測值必須包裹成雙引號字串。`["RollID", "實測值字串"]`。

    #### 💰 模組 B：會計指標提取 (AI 任務：抄錄)
    1. **統計表**：抄錄統計表每一行名稱與實交數量到 `summary_rows`。
    2. **運費與指標**：提取運費項次與標題括號內的 PC 數。你不需抄錄規則文字。

    #### ⚖️ 模組 C：流程稽核 (AI 任務：判定)
    1. **位階檢查**：`未再生 < 研磨 < 再生 < 銲補`。若跨頁面後段尺寸小於前段（銲補除外），報 `🛑流程異常`。

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
    
    generation_config = {"response_mime_type": "application/json", "temperature": 0.0}
    
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name)
        response = model.generate_content([system_prompt, combined_input], generation_config=generation_config)
        
        raw_content = response.text
        # 🛡️ 超級解析器：防止 AI 輸出帶有 Markdown 標籤或廢話
        import re
        json_match = re.search(r"\{.*\}", raw_content, re.DOTALL)
        if json_match:
            raw_content = json_match.group()
            
        parsed_data = json.loads(raw_content)
        parsed_data["_token_usage"] = {
            "input": response.usage_metadata.prompt_token_count, 
            "output": response.usage_metadata.candidates_token_count
        }
        return parsed_data

    except Exception as e:
        return {"job_no": f"JSON Error: {str(e)}", "issues": [], "dimension_data": []}
        
# --- 重點：Python 引擎獨立於 agent 函式之外 ---

def python_process_audit(dimension_data):
    process_issues = []
    roll_history = {} # { "ID": [{"p": "cat", "v": 190, "page": 1}, ...] }
    if not dimension_data: return []

    for item in dimension_data:
        p_num, ds, cat = item.get("page", "?"), item.get("ds", ""), str(item.get("category", "")).strip()
        pairs = [p.split(":") for p in ds.split("|") if ":" in p]
        for rid, val_str in pairs:
            try:
                val = float(re.findall(r"\d+\.?\d*", val_str)[0])
                rid_clean = rid.strip()
                if rid_clean not in roll_history: roll_history[rid_clean] = []
                roll_history[rid_clean].append({"p": cat, "v": val, "page": p_num, "title": item.get("item_title")})
            except: continue

    weights = {"un_regen": 1, "max_limit": 1, "range": 3, "min_limit": 4}
    for rid, records in roll_history.items():
        if len(records) < 2: continue
        records.sort(key=lambda x: str(x['page']))
        for i in range(len(records) - 1):
            curr, nxt = records[i], records[i+1]
            w_curr = weights.get(curr['p'], 2)
            if "研磨" in curr['title']: w_curr = 2
            w_nxt = weights.get(nxt['p'], 2)
            if "研磨" in nxt['title']: w_nxt = 2
            
            # 💡 關鍵判定：後段位階大，數值就不應該變小
            if w_nxt > w_curr and nxt['v'] < curr['v']:
                process_issues.append({
                    "page": nxt['page'], "item": f"編號 {rid} 尺寸位階檢查",
                    "issue_type": "🛑流程異常(尺寸倒置)",
                    "common_reason": f"後段{nxt['p']}尺寸小於前段{curr['p']}",
                    "failures": [{"id": rid, "val": f"後:{nxt['v']} < 前:{curr['v']}", "calc": "尺寸不符位階邏輯"}]
                })
    return process_issues
    
def python_numerical_audit(dimension_data):
    grouped_errors = {}
    import re
    if not dimension_data: return []

    for item in dimension_data:
        # 1. 取得數據 (DS 格式)
        ds = str(item.get("ds", ""))
        if not ds: continue
        raw_entries = [p.split(":") for p in ds.split("|") if ":" in p]
        
        title = str(item.get("item_title", "")).replace(" ", "").replace('"', "")
        cat = str(item.get("category", "")).strip()
        page_num = item.get("page", "?")
        raw_spec = str(item.get("std_spec", "")).replace('"', "")
        
        # 2. 🛡️ 數據清洗與 mm 定位 (每個項目只算一次，提速關鍵)
        mm_nums = [float(n) for n in re.findall(r"(\d+\.?\d*)\s*mm", raw_spec)]
        all_nums = [float(n) for n in re.findall(r"(\d+\.?\d*)", raw_spec)]
        noise = [350.0, 300.0, 200.0, 145.0, 130.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0]
        # 免死金牌：緊貼 mm 的數字不准過濾
        clean_std = [n for n in all_nums if (n in mm_nums) or (n not in noise and n > 5)]

        # 3. 💡 公差自動預算 (解決 AI 算數學慢的問題)
        s_ranges = []
        pm_match = re.search(r"(\d+\.?\d*)\s*[±]\s*(\d+\.?\d*)", raw_spec)
        dev_match = re.search(r"(\d+\.?\d*)\s*[\+]\s*(\d+\.?\d*)\s*,\s*[\-]\s*(\d+\.?\d*)", raw_spec)
        if pm_match:
            b, o = float(pm_match.group(1)), float(pm_match.group(2))
            s_ranges.append([b - o, b + o])
        elif dev_match:
            b, p, m = float(dev_match.group(1)), float(dev_match.group(2)), float(dev_match.group(3))
            s_ranges.append([b - m, b + p])

        # 4. 💡 預算基準 (移出循環，提升 40 倍速度)
        logic = item.get("sl", {})
        l_type = logic.get("lt", "")
        s_threshold = logic.get("t", 0)
        
        un_regen_target = None
        if "un_regen" in l_type or ("未再生" in (cat + title) and "軸頸" not in (cat + title)):
            cands = [n for n in clean_std if n >= 120.0]
            if s_threshold and float(s_threshold) >= 120.0: cands.append(float(s_threshold))
            if cands: un_regen_target = max(cands)

        # 5. 開始逐一判定
        for entry in raw_entries:
            if len(entry) < 2: continue
            rid, val_raw = entry[0].strip(), entry[1].strip()
            # 只取第一個數字過濾手寫
            v_m = re.findall(r"\d+\.?\d*", val_raw)
            val_str = v_m[0] if v_m else val_raw
            
            try:
                val = float(val_str)
                is_two_dec = "." in val_str and len(val_str.split(".")[-1]) == 2
                is_pure_int = "." not in val_str
                is_passed, reason, t_used, engine_label = True, "", "N/A", "未知"

                # A. 銲補
                if "min_limit" in l_type or "銲補" in (cat + title):
                    engine_label = "銲補"
                    if not is_pure_int: is_passed, reason = False, "應為純整數"
                    elif clean_std:
                        t_used = min(clean_std, key=lambda x: abs(x - val))
                        if val < t_used: is_passed, reason = False, "數值不足"
                # B. 未再生本體
                elif un_regen_target is not None:
                    engine_label = "未再生"
                    t_used = un_regen_target
                    if val <= t_used:
                        if not is_pure_int: is_passed, reason = False, "應為整數"
                    elif not is_two_dec: is_passed, reason = False, "應填兩位小數"
                # C. 精加工/區間
                elif any(x in (cat + title) for x in ["再生", "精加工", "研磨", "車修", "組裝", "拆裝", "真圓度"]):
                    engine_label = "精加工"
                    if not is_two_dec: is_passed, reason = False, "應填兩位小數"
                    elif s_ranges:
                        t_used = str(s_ranges)
                        if not any(r[0] <= val <= r[1] for r in s_ranges): is_passed, reason = False, "不在區間內"

                if not is_passed:
                    key = (page_num, title, reason)
                    if key not in grouped_errors:
                        grouped_errors[key] = {"page": page_num, "item": title, "issue_type": f"異常({engine_label})", "common_reason": reason, "failures": []}
                    grouped_errors[key]["failures"].append({"id": rid, "val": val_str, "target": f"基準:{t_used}"})
            except: continue
    return list(grouped_errors.values())
    
def python_accounting_audit(dimension_data, res_main):
    """
    Python 會計官：負責所有數量的精確對帳與運費計算
    """
    accounting_issues = []
    from thefuzz import fuzz
    from collections import Counter
    import re

    # 1. 取得對帳基準 (來自左上角統計表)
    summary_rows = res_main.get("summary_rows", [])
    global_sum_tracker = {s['title']: {"target": s['target'], "actual": 0, "details": []} for s in summary_rows if s.get('title')}
    
    # 💡 取得運費目標 (左上角)
    freight_target = res_main.get("freight_target", 0)
    freight_actual_sum = 0
    freight_details = []

    # 2. 開始逐項過帳
    for item in dimension_data:
        title = item.get("item_title", "")
        page = item.get("page", "?")
        target_pc = item.get("item_pc_target", 0)
        rules = item.get("accounting_rules", {})
        
        # 💡 解開壓縮數據 ds
        ds = str(item.get("ds", ""))
        data_list = [pair.split(":") for pair in ds.split("|") if ":" in pair]
        if not data_list: continue
        
        # 準備編號清單用於計數與重複檢查
        ids = [str(e[0]).strip() for e in data_list if len(e) > 0]
        id_counts = Counter(ids)

        # --- 2.1 單項 PC 數核對 (Local Rule) ---
        u_local = str(rules.get("local", ""))
        is_body = "本體" in title
        is_journal = any(k in title for k in ["軸頸", "內孔", "Journal"])
        
        if "1SET=4PCS" in u_local: actual_item_qty = len(data_list) / 4
        elif "1SET=2PCS" in u_local: actual_item_qty = len(data_list) / 2
        elif is_body or "PC=PC" in u_local: actual_item_qty = len(set(ids)) # 本體去重
        else: actual_item_qty = len(data_list)

        if actual_item_qty != target_pc and target_pc > 0:
            accounting_issues.append({
                "page": page, "item": title, "issue_type": "統計不符(單項)",
                "common_reason": f"要求 {target_pc}PC，內文數到 {actual_item_qty}",
                "failures": [{"id": "標題目標", "val": target_pc}, {"id": "內文計數", "val": actual_item_qty}]
            })

        # --- 2.2 軸頸重複性檢查 (限 2 次) ---
        if is_journal:
            for rid, count in id_counts.items():
                if count >= 3:
                    accounting_issues.append({
                        "page": page, "item": title, "issue_type": "🛑編號重複異常",
                        "common_reason": f"編號 {rid} 出現 {count} 次，軸頸限 2 次",
                        "failures": [{"id": rid, "val": count, "calc": "禁止超過2次"}]
                    })

        # --- 2.3 總表對帳 (A聚合/B一般) ---
        u_agg_raw = str(rules.get("agg", ""))
        agg_multiplier = 1.0
        # 解析單位換算
        conv = re.search(r"(\d+)SET=1PC", u_agg_raw)
        if conv: agg_multiplier = 1.0 / float(conv.group(1))

        for s_title, data in global_sum_tracker.items():
            is_rep = any(k in s_title for k in ["ROLL車修"])
            is_weld = "ROLL車修" in s_title
            is_assem = any(k in s_title for k in ["ROLL拆裝"])
            
            match = False
            # A 模式：聚合籃子
            if (is_rep or is_weld or is_assem) and "豁免" not in u_agg_raw:
                if is_rep and any(k in title for k in ["未再生", "再生"]): match = True
                elif is_weld and "銲補" in title: match = True
                elif is_assem and any(k in title for k in ["拆裝", "組裝", "真圓度"]): match = True
            # B 模式：名字直接對帳
            if not match and fuzz.partial_ratio(s_title, title) > 85: match = True

            if match:
                val_for_agg = actual_item_qty * agg_multiplier
                data["actual"] += val_for_agg
                data["details"].append({"id": f"{title} (P.{page})", "val": val_for_agg, "calc": "計入總帳"})

        # --- 2.4 💡 運費核對 (Freight Check - 補回) ---
        u_fr = str(rules.get("freight", ""))
        # 規則：全卷「本體」+「未再生」之項目
        if "計入" in u_fr or ("本體" in title and "未再生" in title):
            if "豁免" not in u_fr:
                freight_actual_sum += actual_item_qty
                freight_details.append({"id": f"{title} (P.{page})", "val": actual_item_qty, "calc": "計入運費"})

    # 3. 結算最終結果報告
    # 總表報告
    for s_title, data in global_sum_tracker.items():
        if abs(data["actual"] - data["target"]) > 0.01 and data["target"] > 0:
            accounting_issues.append({
                "page": "總表", "item": s_title, "issue_type": "統計不符(總帳)",
                "common_reason": f"標註 {data['target']} != 實際 {data['actual']}",
                "failures": [{"id": "🔍 統計基準", "val": data["target"]}] + data["details"] + [{"id": "🧮 實際總計", "val": data["actual"]}]
            })

    # 運費報告
    if abs(freight_actual_sum - freight_target) > 0.01 and freight_target > 0:
        accounting_issues.append({
            "page": "總表", "item": "運費核對", "issue_type": "統計不符(運費)",
            "common_reason": f"運費基準 {freight_target} != 實際 {freight_actual_sum}",
            "failures": [{"id": "🚚 運費基準", "val": freight_target}] + freight_details + [{"id": "🧮 運費總計", "val": freight_actual_sum}]
        })
        
    return accounting_issues
    
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
        total_start = time.time()
        status = st.empty()
        progress_bar = st.progress(0)
            
        extracted_data_list = [None] * len(st.session_state.photo_gallery)
        full_text_for_search = ""
        total_imgs = len(st.session_state.photo_gallery)
            
        ocr_start = time.time()

        def process_image_task(index, item):
            index = int(index)
            # 如果已經有資料了就不重複掃描
            if item.get('table_md') and item.get('header_text') and item.get('full_text'):
                real_page = item.get('real_page', str(index + 1))
                return index, item['table_md'], item['header_text'], item['full_text'], None, real_page, None
    
            try:
                if item.get('file') is None:
                    return index, None, None, None, None, None, "無圖片檔案"
                
                item['file'].seek(0)
                # 這裡會接到我們剛才修改後回傳的 None
                table_md, header, full, _, real_page = extract_layout_with_azure(item['file'], DOC_ENDPOINT, DOC_KEY)
                return index, table_md, header, full, None, real_page, None
            except Exception as e:
                return index, None, None, None, None, None, f"OCR失敗: {str(e)}"

        status.text(f"Azure 正在平行掃描 {total_imgs} 頁文件...")

        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            futures = []
            for i, item in enumerate(st.session_state.photo_gallery):
                futures.append(executor.submit(process_image_task, i, item))
            
            completed_count = 0
            for future in concurrent.futures.as_completed(futures):
                idx, t_md, h_txt, f_txt, raw_j, r_page, err = future.result()
                idx = int(idx)
                
                if err:
                    st.error(f"第 {idx+1} 頁讀取失敗: {err}")
                    extracted_data_list[idx] = None
                else:
                    st.session_state.photo_gallery[idx]['table_md'] = t_md
                    st.session_state.photo_gallery[idx]['header_text'] = h_txt
                    st.session_state.photo_gallery[idx]['full_text'] = f_txt
                    st.session_state.photo_gallery[idx]['raw_json'] = raw_j
                    st.session_state.photo_gallery[idx]['real_page'] = r_page
                    st.session_state.photo_gallery[idx]['file'] = None
                    
                    extracted_data_list[idx] = {
                        "page": r_page,
                        "table": t_md or "", 
                        "header_text": h_txt or ""
                    }
                
                completed_count += 1
                progress_bar.progress(completed_count / (total_imgs + 1))
        
        for i, data in enumerate(extracted_data_list):
            if data and isinstance(data, dict):
                page_idx = i
                if 0 <= page_idx < len(st.session_state.photo_gallery):
                    full_text_for_search += st.session_state.photo_gallery[page_idx].get('full_text', '')

        ocr_end = time.time()
        ocr_duration = ocr_end - ocr_start

        combined_input = "以下是各頁資料：\n"
        for i, data in enumerate(extracted_data_list):
            if data is None: continue
            page_num = data.get('page', i+1)
            table_text = data.get('table', '')
            header_text = data.get('header_text', '')
            combined_input += f"\n=== Page {page_num} ===\n【頁首】:\n{header_text}\n【表格】:\n{table_text}\n"
            
        status.text("總稽核 Agent 正在進行全方位分析...")
        
        # 1. 執行 AI 
        res_main = agent_unified_check(combined_input, combined_input, GEMINI_KEY, main_model_name)
        
        # 💡 [重大修正]：從 AI 回傳中抓取維度數據
        dim_data = res_main.get("dimension_data", [])
        
        # 2. 執行三個 Python 引擎 (數值、會計、流程)
        python_numeric_issues = python_numerical_audit(dim_data)
        python_accounting_issues = python_accounting_audit(dim_data, res_main)
        
        # 💡 [新增]：啟動 Python 流程稽核引擎
        python_process_issues = python_process_audit(dim_data)
        
        # 3. 合併結果 (帶有防呆檢查，並確保權力徹底移交) ---
        ai_raw_issues = res_main.get("issues", [])
        ai_filtered_issues = []

        if isinstance(ai_raw_issues, list):
            for i in ai_raw_issues:
                if isinstance(i, dict):
                    i['source'] = '🤖 總稽核 AI'
                    i_type = str(i.get("issue_type", ""))
                    
                    # 💡 [關鍵修正]：
                    # 我們只保留 AI 發現的：規格提取失敗、未匹配規則、還有表頭資訊不符。
                    # 「流程」和「統計」已經完全交給 Python 引擎了，所以這裡絕對不留 AI 報的。
                    ai_tasks_to_keep = ["規格提取失敗", "未匹配", "表頭"]
                    if any(k in i_type for k in ai_tasks_to_keep):
                        ai_filtered_issues.append(i)
                else:
                    # 如果 AI 回傳格式崩潰，至少保留原始文字供檢查
                    ai_filtered_issues.append({
                        "page": "?", "item": "AI 回傳解析異常", "issue_type": "⚠️格式錯誤",
                        "common_reason": f"原始內容: {str(i)}", "source": "🤖 總稽核 AI"
                    })

        # 4. 取得 Python 表頭檢查 (日期、工令等)
        python_header_issues, python_debug_data = python_header_check(st.session_state.photo_gallery)
        
        # 最終合併：AI(提取警告) + Python(數值) + Python(會計) + Python(流程) + Python(表頭)
        all_issues = ai_filtered_issues + python_numeric_issues + python_accounting_issues + python_process_issues + python_header_issues
        
        # 5. 存入快取 (這是 Debug 頁面能顯示數據的唯一關鍵)
        st.session_state.analysis_result_cache = {
            "job_no": res_main.get("job_no", "Unknown"),
            "all_issues": all_issues,
            "total_duration": time.time() - total_start,
            "cost_twd": (res_main.get("_token_usage",{}).get("input",0)*0.5 + res_main.get("_token_usage",{}).get("output",0)*3.0)/1000000*32.5,
            "total_in": res_main.get("_token_usage",{}).get("input", 0),
            "total_out": res_main.get("_token_usage",{}).get("output", 0),
            "ocr_duration": ocr_duration,
            "time_eng": time.time() - total_start - ocr_duration,
            "full_text_for_search": combined_input,
            "combined_input": combined_input,
            "python_debug_data": python_debug_data,
            # ✅ 這行沒加，Debug 頁面就是空的！
            "ai_extracted_data": dim_data 
        }
        
    if st.session_state.analysis_result_cache:
        cache = st.session_state.analysis_result_cache
        all_issues = cache['all_issues']
        
        st.success(f"工令: {cache['job_no']} | ⏱️ {cache['total_duration']:.1f}s")
        st.info(f"💰 本次成本: NT$ {cache['cost_twd']:.2f} (In: {cache['total_in']:,} / Out: {cache['total_out']:,})")
        st.caption(f"細節耗時: Azure OCR {cache['ocr_duration']:.1f}s | AI 分析 {cache['time_eng']:.1f}s")
        
        with st.expander("🔍 查看 AI 讀取到的 Excel 規則 (Debug)"):
            rules_text = get_dynamic_rules(cache['full_text_for_search'], debug_mode=True)
            if "無特定規則" in rules_text:
                st.caption("無匹配規則")
            else:
                st.markdown(rules_text)
                
        # --- 新增的 Debug 展開頁 ---
        with st.expander("🔬 查看 AI 抄錄給 Python 的原始數據 (檢查手寫過濾)", expanded=False):
            raw_dim_data = cache.get("ai_extracted_data", [])
            if raw_dim_data:
                st.write("這是 AI 抄錄並翻譯後的 JSON（包含格式是否正確、數字是否被簡化）：")
                st.json(raw_dim_data)
            else:
                st.caption("無數據提取資料。")

        with st.expander("🐍 查看 Python 硬邏輯偵測結果 (Debug)", expanded=False):
            if cache.get('python_debug_data'):
                p_data = cache['python_debug_data']
                standard_data = {}
                all_values = {"工令編號": [], "預定交貨": [], "實際交貨": []}
                for page in p_data:
                    for k in all_values.keys():
                        if page.get(k) and page[k] != "N/A":
                            all_values[k].append(page[k])
                
                standard_row = {"頁碼": "🏆 判定標準"}
                for k, v in all_values.items():
                    if v:
                        standard_row[k] = Counter(v).most_common(1)[0][0]
                    else:
                        standard_row[k] = "N/A"
                
                final_df_data = [standard_row] + p_data
                st.dataframe(final_df_data, use_container_width=True, hide_index=True)
                st.info("💡 「判定標準」是依據多數決產生的。")
            else:
                st.caption("無偵測資料")

        real_errors = [i for i in all_issues if "未匹配" not in i.get('issue_type', '')]
        
        if not real_errors:
            st.balloons()
            if not all_issues:
                st.success("✅ 全數合格！")
            else:
                st.success(f"✅ 數值全數合格！ (但有 {len(all_issues)} 個項目未匹配規則，請檢查)")
        else:
            st.error(f"發現 {len(real_errors)} 類數值異常，另有 {len(all_issues) - len(real_errors)} 個項目未匹配規則")

        for item in all_issues:
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                
                source_label = item.get('source', '')
                issue_type = item.get('issue_type', '異常')
                
                c1.markdown(f"**P.{item.get('page', '?')} | {item.get('item')}**  `{source_label}`")
                
                # 顏色控制：會計統計類用紅色，規格類用黃色
                if "統計" in issue_type or "數量" in issue_type or "流程" in issue_type:
                    c2.error(f"🛑 {issue_type}")
                else:
                    c2.warning(f"⚠️ {issue_type}")
                
                st.caption(f"原因: {item.get('common_reason', '')}")
                
                # --- 渲染表格 (會計對帳單) ---
                failures = item.get('failures', [])
                if failures:
                    table_data = []
                    for f in failures:
                        if isinstance(f, dict):
                            # 我們統一使用這四個欄位標題，會計與工程共用
                            row = {
                                "項目/滾輪編號": f.get('id', '未知'), 
                                "實測/計數": f.get('val', 'N/A'),
                                "標準/備註": f.get('target', ''), # 工程用
                                "判定算式/狀態": f.get('calc', '') # 會計用
                            }
                            # 如果是會計模式，把 target 留空，資訊主要在 id 和 val
                            table_data.append(row)
                    
                    if table_data:
                        st.dataframe(table_data, use_container_width=True, hide_index=True)
                else:
                    # 如果沒有 failures，至少顯示一個數據提示
                    st.info(f"詳細數據見上述原因說明")
        
        st.divider()

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

        with st.expander("👀 查看傳給 AI 的最終文字 (Prompt Input)"):
            st.caption("這才是 AI 真正讀到的內容 (已過濾雜訊)：")
            st.code(cache['combined_input'], language='markdown')
    
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
