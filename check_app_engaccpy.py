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
        "GPT-5 Mini": "models/gpt-5-mini-2025-08-07",
        "GPT-5 Nano": "models/gpt-5-nano-2025-08-07",
        
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

# --- Excel 規則讀取函數 (專業極簡版) ---
@st.cache_data
def get_dynamic_rules(ocr_text, debug_mode=False):
    try:
        df = pd.read_excel("rules.xlsx")
        df.columns = [c.strip() for c in df.columns]
        ocr_text_clean = str(ocr_text).upper().replace(" ", "").replace("\n", "")
        
        ai_prompt_list = []    # 給 AI 的 (純文字)
        debug_view_list = []   # 給人看的 (排版清潔)

        for index, row in df.iterrows():
            item_name = str(row.get('Item_Name', '')).strip()
            if not item_name or "(通用)" in item_name: continue
            
            score = fuzz.partial_ratio(item_name.upper().replace(" ", ""), ocr_text_clean)
            if score >= 85:
                # 取值與清洗
                def clean(v): return str(v).strip() if v and str(v) != 'nan' else None
                
                spec = clean(row.get('Standard_Spec', ''))
                logic = clean(row.get('Logic_Prompt', ''))
                u_fr = clean(row.get('Unit_Rule_Freight', ''))
                u_loc = clean(row.get('Unit_Rule_Local', ''))
                u_agg = clean(row.get('Unit_Rule_Agg', ''))

                # --- A. 建構 AI Prompt (維持不變) ---
                if not debug_mode:
                    if spec or logic:
                        desc = f"- [參考資訊] {item_name}\n"
                        if spec: desc += f"  - 標準規格: {spec}\n"
                        if logic: desc += f"  - 注意事項: {logic}\n"
                        ai_prompt_list.append(desc)
                
                # --- B. 建構 Debug 顯示 (去除圖案，改用表格感排版) ---
                else:
                    # 使用 Markdown 的引用區塊 (>) 來做層級區分，看起來很乾淨
                    block = f"#### ■ {item_name} (匹配度 {score}%)\n"
                    
                    # AI 區塊
                    block += "**[ AI Prompt 輸入 ]**\n"
                    if spec or logic:
                        if spec: block += f"- 規格標準 : `{spec}`\n"
                        if logic: block += f"- 注意事項 : `{logic}`\n"
                    else:
                        block += "- (無特定輸入)\n"

                    # Python 區塊
                    block += "\n**[ Python 硬邏輯設定 ]**\n"
                    has_py = False
                    if u_fr: 
                        block += f"- 運費邏輯 : `{u_fr}`\n"
                        has_py = True
                    if u_loc:
                        block += f"- 單項規則 : `{u_loc}`\n"
                        has_py = True
                    if u_agg:
                        block += f"- 聚合規則 : `{u_agg}`\n"
                        has_py = True
                    
                    if not has_py:
                        block += "- (使用預設邏輯)\n"
                    
                    block += "\n---\n"
                    debug_view_list.append(block)

        if debug_mode:
            if not debug_view_list: return "無特定規則命中。"
            return "\n".join(debug_view_list)
        else:
            return "\n".join(ai_prompt_list) if ai_prompt_list else ""

    except Exception as e:
        return f"讀取錯誤: {e}"

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
            # 1. 取得頁碼 (保留原邏輯)
            page_num = "Unknown"
            if table.bounding_regions: page_num = table.bounding_regions[0].page_number
            
            # =========================================================
            # 🔍 [新增] 智慧標籤偵測：在處理表格前，先判斷它是誰
            # =========================================================
            table_tag = "未知表格"
            
            # 技巧：抓取表格「第一列 (row_index=0)」的所有文字來判斷
            # 這樣不用讀完整張表，只要看表頭就知道它是總表還是明細
            first_cells = [c.content for c in table.cells if c.row_index == 0]
            first_row_text = "".join(first_cells)
            
            # 定義關鍵字 (您可以根據實際表格微調)
            summary_keywords = ["實交", "申請", "名稱及規範", "完成交貨日期", "存放位置"]
            detail_keywords = ["規範標準", "檢驗紀錄", "實測", "編號", "尺寸", "W3 #", "公差"]

            if any(k in first_row_text for k in summary_keywords):
                table_tag = "SUMMARY_TABLE (總表)"
            elif any(k in first_row_text for k in detail_keywords):
                table_tag = "DETAIL_TABLE (明細表)"
            
            # 📝 [修改] 輸出標頭：這裡不再只寫 Table X，而是加上我們判斷的標籤
            # 加上 "===" 是為了讓 Prompt 裡的「注意範圍」指令能精準鎖定
            markdown_output += f"\n\n=== [{table_tag} | Page {page_num}] ===\n"
            # =========================================================

            rows = {}
            stop_processing_table = False 
            
            # --- 以下保留您原本的 Cell 處理邏輯，完全不用動 ---
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
    
def assign_category_by_python(item_title):
    """
    Python 分類官 (新增關鍵字：粗車、精車)
    """
    # 🧽 預處理
    t = str(item_title).upper().replace(" ", "").replace("\n", "").replace('"', "")
    
    # --- LEVEL 1：銲補與裝配 (最高優先) ---
    if any(k in t for k in ["銲補", "銲接", "WELD"]):
        return "min_limit"
    
    if any(k in t for k in ["組裝", "拆裝", "裝配", "真圓度", "ASSY"]):
        return "range"

    # --- LEVEL 2：未再生判定 (含粗車) ---
    # ⚡️ [新增] 關鍵字：粗車
    if any(k in t for k in ["未再生", "UN_REGEN", "粗車"]):
        # a. 含「軸頸」 -> max_limit
        if any(k in t for k in ["軸頸", "內孔", "JOURNAL"]):
            return "max_limit"
        # b. 不含「軸頸」(即本體) -> un_regen
        else:
            return "un_regen"

    # --- LEVEL 3：精加工判定 (含精車) ---
    # ⚡️ [新增] 關鍵字：精車
    if any(k in t for k in ["再生", "研磨", "精加工", "車修", "KEYWAY", "GRIND", "MACHIN", "精車"]):
        return "range"

    return "unknown"
    
def consolidate_issues(issues):
    """
    🗂️ 異常合併器：將「項目」、「錯誤類型」、「原因」完全相同的異常合併成一張卡片
    """
    grouped = {}
    
    for i in issues:
        # 1. 產生合併鑰匙 (Key)：項目 + 類型 + 原因
        # 這樣確保只有真正一樣的問題才會被並在一起
        key = (i.get('item', ''), i.get('issue_type', ''), i.get('common_reason', ''))
        
        if key not in grouped:
            # 初始化：複製第一筆資料
            grouped[key] = i.copy()
            # 把頁碼轉成 Set 集合 (避免重複)
            grouped[key]['pages_set'] = {str(i.get('page', '?'))}
            # 確保 failures 是獨立的 list
            grouped[key]['failures'] = i.get('failures', []).copy()
        else:
            # 合併：把新的頁碼加進去
            grouped[key]['pages_set'].add(str(i.get('page', '?')))
            # 合併：把新的證據 (failures) 加到表格裡
            grouped[key]['failures'].extend(i.get('failures', []))
            
    # 2. 轉回 List 並整理頁碼格式
    result = []
    for key, val in grouped.items():
        # 頁碼排序：讓它顯示 P.1, P.3, P.5 而不是亂跳
        sorted_pages = sorted(list(val['pages_set']), key=lambda x: int(x) if x.isdigit() else 999)
        val['page'] = ", ".join(sorted_pages) # 變成字串 "1, 3, 5"
        
        # 移除暫存的 set
        del val['pages_set']
        result.append(val)
        
    return result

# --- 5. 總稽核 Agent (雙核心引擎版：Gemini + OpenAI) ---
def agent_unified_check(combined_input, full_text_for_search, api_key, model_name):
    # 1. 準備 Prompt (規則與指令)
    dynamic_rules = get_dynamic_rules(full_text_for_search)

    system_prompt = f"""
    你是一位極度嚴謹的中鋼機械品管【數據抄錄員】。你必須像「電腦程式」一樣執行任務。
    
    {dynamic_rules}

    ---

    #### ⚔️ 模組 A：工程尺寸數據提取 (AI 任務：純抄錄)
    ⚠️ **注意範圍**：你只能從標記為 `=== [DETAIL_TABLE (明細表)] ===` 的區域提取數據。
    
    1. **規格抄錄 (std_spec)**：精確抄錄標題中含 `mm`、`±`、`+`、`-` 的原始文字。
    
    2. **標題抄錄 (item_title)**：⚠️ 極度重要！必須完整抄錄項目標題，**嚴禁遺漏**「未再生」、「銲補」、「車修」、「軸頸」等關鍵字。
    
    3. **目標數量提取 (item_pc_target)**：
       - 請從標題中提取括號內的數量要求（例如標題含 `(4SET)` 則提取 `4`，`(10PC)` 則提取 `10`）。
       - 若無括號標註數量，請填 `0`。

    4. **分類 (category)**：**請直接回傳 `null`**。由後端程式判定。

    5. **數據抄錄 (ds) 與 字串保護規範**：
       - **格式**：輸出為 `"ID:值|ID:值"` 的字串格式。
       - **禁止簡化**：實測值若顯示 `349.90`，必須輸出 `"349.90"`，保留尾數 0。
       - **🚫 遇到干擾不鑽牛角尖**：若儲存格內的數值因手寫塗改、圓圈遮擋、污點、字跡黏連或光線反光，導致你無法「100% 確定」原始打印數字時，**嚴禁腦補或猜測**。
       - **壞軌標記 [BAD]**：請將該筆數值直接標記為 `[!]`。
       - **範例**：若 ID 清楚但數值模糊 -> `"V100:[!]"`；若整個儲存格都看不清 -> `"[!] : [!]"`。
       - **跳過策略**：一旦標記為 `[!]`，請立即跳到下一格，不要浪費 Token 描述雜訊。

    #### 💰 模組 B：會計指標提取 (AI 任務：抄錄)
    ⚠️ **注意範圍**：你只能從標記為 `=== [SUMMARY_TABLE (總表)] ===` 的區域提取數據。
    
    1. **統計表**：請鎖定 `實交數量` 欄位。抄錄每一行的「名稱」與「實交數量」到 `summary_rows`。
    (無需額外提取運費或特殊指標，只要完整抄錄表格行項目即可)

    ---

    ### 📝 輸出規範 (Output Format)
    必須回傳單一 JSON。注意：AI 不需回傳流程異常，僅需回傳原始數據。

    {{
      "job_no": "工令",
      "summary_rows": [ {{ "title": "名", "target": 數字 }} ],
      "freight_target": 0, 
      "issues": [], 
      "dimension_data": [
         {{
           "page": 數字, "item_title": "標題", "category": null, 
           "item_pc_target": 0,
           "accounting_rules": {{ "local": "", "agg": "", "freight": "" }},
           "sl": {{ "lt": "null", "t": 0 }},
           "std_spec": "原始規格文字",
           "ds": "ID:值|ID:值" 
         }}
      ]
    }}
    """

    # 2. 判斷要使用哪一顆引擎
    raw_content = ""
    
    # --- 引擎 A: OpenAI GPT 系列 ---
    if "gpt" in model_name.lower():
        try:
            # 必須使用全域變數 OPENAI_KEY，因為傳入的 api_key 參數通常是 GEMINI_KEY
            openai_key = st.secrets.get("OPENAI_KEY", "")
            if not openai_key:
                return {"job_no": "Error: 缺少 OPENAI_KEY", "issues": [], "dimension_data": []}
                
            client = OpenAI(api_key=openai_key)
            
            response = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": combined_input}
                ],
                temperature=0.0,
                response_format={"type": "json_object"} # GPT-4o 支援強制 JSON 模式
            )
            raw_content = response.choices[0].message.content
            
            # 模擬 Token 用量 (OpenAI 格式不同，這裡做個簡單轉換以便統一顯示)
            input_tokens = response.usage.prompt_tokens
            output_tokens = response.usage.completion_tokens
            
        except Exception as e:
            return {"job_no": f"OpenAI Error: {str(e)}", "issues": [], "dimension_data": []}

    # --- 引擎 B: Google Gemini 系列 ---
    else:
        try:
            genai.configure(api_key=api_key) # 這裡用傳入的 GEMINI_KEY
            generation_config = {"response_mime_type": "application/json", "temperature": 0.0}
            model = genai.GenerativeModel(model_name)
            
            # Gemini 2.0 可能需要不同的呼叫方式，這裡保持通用接口
            response = model.generate_content([system_prompt, combined_input], generation_config=generation_config)
            raw_content = response.text
            
            input_tokens = response.usage_metadata.prompt_token_count
            output_tokens = response.usage_metadata.candidates_token_count
            
        except Exception as e:
            return {"job_no": f"Gemini Error: {str(e)}", "issues": [], "dimension_data": []}

    # 3. 統一解析與回傳
    try:
        # 🛡️ 超級解析器：防止 AI 輸出帶有 Markdown 標籤或廢話
        import re
        json_match = re.search(r"\{.*\}", raw_content, re.DOTALL)
        if json_match:
            raw_content = json_match.group()
            
        parsed_data = json.loads(raw_content)
        
        # 統一 Token 用量格式
        parsed_data["_token_usage"] = {
            "input": input_tokens, 
            "output": output_tokens
        }
        return parsed_data

    except Exception as e:
        return {"job_no": f"JSON Parsing Error: {str(e)}", "issues": [], "dimension_data": []}

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
    Python 會計官 (智慧匹配修復版)
    1. 規則查找升級：若精準匹配失敗，自動嘗試「脫殼」(移除標題括號) 與「部分匹配」。
    2. 規則字串清洗：保留防彈級的字元正規化 (全形轉半形)。
    """
    accounting_issues = []
    from thefuzz import fuzz
    from collections import Counter
    import re
    import pandas as pd 

    # 🧽 基礎清洗工具
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
    rules_map = {}
    try:
        df = pd.read_excel("rules.xlsx")
        df.columns = [c.strip() for c in df.columns]
        for _, row in df.iterrows():
            iname = str(row.get('Item_Name', '')).strip()
            if iname: 
                rules_map[clean_text(iname)] = {
                    "u_local": str(row.get('Unit_Rule_Local', '')).strip(),
                    "u_fr": str(row.get('Unit_Rule_Freight', '')).strip(),
                    "u_agg": str(row.get('Unit_Rule_Agg', '')).strip()
                }
    except: pass 

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
        
        # --- 🔍 查找 Excel 規則 (⚡️ 邏輯升級區) ---
        rule_set = rules_map.get(title_clean)
        
        # 策略 A: 如果直接沒找到，嘗試「脫殼」：把括號 (2SET) 拿掉再找一次
        if not rule_set:
            # 移除 (xxx) 或 （xxx） 的內容
            title_no_suffix = re.sub(r"[\(（].*?[\)）]", "", title_clean)
            rule_set = rules_map.get(title_no_suffix)

        # 策略 B: 如果還是沒找到，改用 Partial Ratio (像 Debug 看板一樣寬容)
        # 只要規則名稱完整出現在標題裡 (e.g. "車修" 在 "車修(2SET)" 裡面)，就算命中
        if not rule_set and rules_map:
            best_score = 0
            for k, v in rules_map.items():
                # 改用 partial_ratio，並將門檻設為 90
                score = fuzz.partial_ratio(k, title_clean)
                if score > 90 and score > best_score:
                    best_score = score
                    rule_set = v
        
        u_local = rule_set.get("u_local", "") if rule_set else ""
        u_fr = rule_set.get("u_fr", "") if rule_set else ""

        # --- 以下邏輯保持不變 (字串清洗與計算) ---
        u_local_norm = u_local.upper().replace(" ", "").replace("　", "").replace("＝", "=").replace("：", "=").replace(":", "=")
        u_fr_norm = u_fr.upper().replace(" ", "").replace("　", "").replace("＝", "=").replace("：", "=").replace(":", "=")

        ds = str(item.get("ds", ""))
        data_list = [pair.split(":") for pair in ds.split("|") if ":" in pair]
        if not data_list: continue
        ids = [str(e[0]).strip() for e in data_list if len(e) > 0]
        id_counts = Counter(ids)

        # 2.1 單項數量計算
        is_local_exempt = "豁免" in u_local
        is_weight_mode = "KG" in title_clean.upper() or target_pc > 100
        
        if is_weight_mode:
            current_sum = 0
            has_bad_sector = False
            for e in data_list:
                temp_val = safe_float(e[1])
                if temp_val == "BAD_DATA": has_bad_sector = True
                else: current_sum += temp_val
            actual_item_qty = current_sum
            if has_bad_sector and not is_local_exempt:
                accounting_issues.append({
                    "page": page, "item": raw_title, "issue_type": "⚠️數據損毀",
                    "common_reason": "含無法辨識重量",
                    "failures": [{"id": "警告", "val": "[!]", "calc": "數據損毀"}]
                })
        else:
            # 🔢 數量模式
            conv_match = re.search(r"1SET=(\d+\.?\d*)", u_local_norm)
            
            if conv_match:
                divisor = float(conv_match.group(1))
                if divisor == 0: divisor = 1 
                actual_item_qty = len(data_list) / divisor
            elif "PC=PC" in u_local_norm or "本體" in title_clean:
                actual_item_qty = len(set(ids))
            else:
                actual_item_qty = len(data_list)

        if not is_local_exempt and abs(actual_item_qty - target_pc) > 0.01 and target_pc > 0:
            accounting_issues.append({
                "page": page, "item": raw_title, "issue_type": "統計不符(單項)",
                "common_reason": f"標題 {target_pc}PC != 內文 {actual_item_qty} (規則:{u_local if u_local else '無'})",
                "failures": [{"id": "目標", "val": target_pc}, {"id": "實際", "val": actual_item_qty}],
                "source": "🐍 會計引擎"
            })

        # 2.2 重複性示警
        if "本體" in title_clean:
             for rid, count in id_counts.items():
                if count > 1:
                     accounting_issues.append({"page": page, "item": raw_title, "issue_type": "⚠️編號重複警示(本體)", "common_reason": f"本體 {rid} 重複 {count}次", "failures": []})
        elif any(k in title_clean for k in ["軸頸", "內孔", "JOURNAL"]):
             for rid, count in id_counts.items():
                if count > 2:
                     accounting_issues.append({"page": page, "item": raw_title, "issue_type": "⚠️編號重複警示(軸頸)", "common_reason": f"軸頸 {rid} 重複 {count}次", "failures": []})

        # 2.3 運費計算
        is_fr_exempt = "豁免" in u_fr
        fr_conv_match = re.search(r"(\d+)[:=]1", u_fr_norm)
        
        is_default_target = "本體" in title_clean and ("未再生" in title_clean or "粗車" in title_clean)

        freight_val_for_item = 0.0
        freight_note = ""

        if is_fr_exempt: freight_val_for_item = 0.0
        elif fr_conv_match:
            divisor = float(fr_conv_match.group(1))
            freight_val_for_item = actual_item_qty / divisor
            freight_note = f"計入 (/{int(divisor)})"
        elif is_default_target:
            freight_val_for_item = actual_item_qty
            freight_note = "計入運費"
            
        if freight_val_for_item > 0:
            freight_actual_sum += freight_val_for_item
            freight_details.append({"id": f"{raw_title}", "val": freight_val_for_item, "calc": freight_note})

        # 2.4 總表對帳
        for s_title, data in global_sum_tracker.items():
            match = False
            s_title_clean = clean_text(s_title)
            
            if "運費" in s_title_clean:
                if freight_val_for_item > 0:
                    data["actual"] += freight_val_for_item
                    data["details"].append({"id": f"{raw_title}", "val": freight_val_for_item, "calc": freight_note})
                continue 
            
            req_body = "本體" in s_title_clean
            req_journal = any(k in s_title_clean for k in ["軸頸", "內孔", "JOURNAL"])
            req_unregen = "未再生" in s_title_clean or "粗車" in s_title_clean
            req_regen_only = ("再生" in s_title_clean or "精車" in s_title_clean) and not req_unregen
            
            is_item_body = "本體" in title_clean
            is_item_journal = any(k in title_clean for k in ["軸頸", "內孔", "JOURNAL"])
            is_item_unregen = "未再生" in title_clean or "粗車" in title_clean
            
            is_main_disassembly = "ROLL拆裝" in s_title_clean 
            is_main_machining = "ROLL車修" in s_title_clean   
            is_main_welding = "ROLL銲補" in s_title_clean     

            if is_main_disassembly:
                if "組裝" in title_clean or "拆裝" in title_clean: match = True
            elif is_main_machining:
                has_part = "軸頸" in title_clean or "本體" in title_clean
                has_action = any(k in title_clean for k in ["再生", "精車", "未再生", "粗車"])
                if has_part and has_action: match = True
            elif is_main_welding:
                has_part = "軸頸" in title_clean or "本體" in title_clean
                if has_part and "銲補" in title_clean: match = True
            else:
                if fuzz.partial_ratio(s_title_clean, title_clean) > 90:
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

    if abs(freight_actual_sum - freight_target) > 0.01 and freight_target > 0:
        accounting_issues.append({
            "page": "總表", "item": "運費核對", "issue_type": "統計不符(運費)",
            "common_reason": f"基準 {freight_target} != 實際 {freight_actual_sum}",
            "failures": [{"id": "🚚 基準", "val": freight_target}] + freight_details + [{"id": "🧮 實際", "val": freight_actual_sum}],
            "source": "🐍 會計引擎"
        })
        
    return accounting_issues

def python_process_audit(dimension_data):
    """
    Python 流程引擎 (通用原因合併版)
    1. 粗車 = 未再生 (Stage 1)
    2. 精車 = 再生 (Stage 3)
    3. 修改：common_reason 不再包含 ID，以便前端卡片合併。
    """
    process_issues = []
    import re
    
    # 定義工序與名稱
    STAGE_MAP = {
        1: "未再生/粗車",
        2: "銲補",
        3: "再生/精車",
        4: "研磨"
    }

    history = {} 

    if not dimension_data: return []

    for item in dimension_data:
        p_num = item.get("page", "?")
        title = str(item.get("item_title", "")).strip()
        ds = str(item.get("ds", ""))
        
        # --- A. 軌道判斷 ---
        track = "Unknown"
        if "本體" in title:
            track = "本體"
        elif any(k in title for k in ["軸頸", "內孔", "JOURNAL"]):
            track = "軸頸"
        else:
            continue 

        # --- B. 工序判斷 ---
        stage = 0
        if "研磨" in title:
            stage = 4
        elif "銲補" in title or "銲接" in title:
            stage = 2
        elif "未再生" in title or "粗車" in title:
            stage = 1
        elif "再生" in title or "精車" in title: 
            stage = 3
        
        if stage == 0: continue 

        # --- C. 數據解析 ---
        segments = ds.split("|")
        for seg in segments:
            parts = seg.split(":")
            if len(parts) < 2: continue
            
            rid = parts[0].strip()
            val_str = parts[1].strip()
            
            nums = re.findall(r"\d+\.?\d*", val_str)
            if not nums: continue
            val = float(nums[0])
            
            key = (rid, track)
            if key not in history: history[key] = {}
            history[key][stage] = {
                "val": val,
                "page": p_num,
                "title": title
            }

    # 2. 執行核心邏輯檢查
    for (rid, track), stages_data in history.items():
        present_stages = sorted(stages_data.keys())
        if not present_stages: continue
        max_stage = present_stages[-1]
        
        # === 邏輯一：溯源檢查 ===
        missing_stages = []
        for req_s in range(1, max_stage):
            if req_s not in stages_data:
                missing_stages.append(STAGE_MAP[req_s])
        
        if missing_stages:
            last_info = stages_data[max_stage]
            # ⚡️ [修改點] common_reason 移除 {rid}，改成通用描述
            # 舊: f"[{track}] {rid} 進度至..." -> 不能合併
            # 新: f"[{track}] 進度至..." -> 可以合併！
            process_issues.append({
                "page": last_info['page'],
                "item": f"{last_info['title']}", # 保留標題，如果標題不同還是會分開，這通常是好事
                "issue_type": "🛑溯源異常(缺漏工序)",
                "common_reason": f"[{track}] 進度至【{STAGE_MAP[max_stage]}】，缺前置：{', '.join(missing_stages)}",
                "failures": [{"id": rid, "val": "缺漏", "calc": "履歷不完整"}],
                "source": "🐍 流程引擎"
            })

        # === 邏輯二：尺寸檢查 ===
        size_rank = { 1: 10, 4: 20, 3: 30, 2: 40 }
        
        for i in range(len(present_stages)):
            for j in range(i + 1, len(present_stages)):
                s_a = present_stages[i]
                s_b = present_stages[j]
                info_a = stages_data[s_a]
                info_b = stages_data[s_b]
                
                expect_a_smaller = size_rank[s_a] < size_rank[s_b]
                is_violation = False
                if expect_a_smaller:
                    if info_a['val'] >= info_b['val']: is_violation = True
                else:
                    if info_a['val'] <= info_b['val']: is_violation = True
                    
                if is_violation:
                    sign = "<" if expect_a_smaller else ">"
                    # ⚡️ [修改點] 同樣移除 common_reason 裡的 ID
                    process_issues.append({
                        "page": info_b['page'],
                        "item": f"[{track}] 尺寸邏輯檢查", # 這裡把 item 也改通用一點，確保跨頁合併
                        "issue_type": "🛑流程異常(尺寸倒置)",
                        "common_reason": f"尺寸邏輯錯誤：{STAGE_MAP[s_a]} 應 {sign} {STAGE_MAP[s_b]}",
                        "failures": [
                            {"id": f"{rid} ({STAGE_MAP[s_a]})", "val": info_a['val'], "calc": "前工序"},
                            {"id": f"{rid} ({STAGE_MAP[s_b]})", "val": info_b['val'], "calc": "後工序"}
                        ],
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
            
            # ⚡️ [插入點]：Python 奪權！強制覆寫 AI 的分類
            for item in dim_data:
                # 即使 AI 有填 category，我們也用 Python 的邏輯覆蓋它，保證 100% 一致性
                # 或者，如果 AI 沒填，這裡就是補填的關鍵
                new_cat = assign_category_by_python(item.get("item_title", ""))
                item["category"] = new_cat
                # 順便把 category 寫進 rules 供前端顯示 (選用)
                if "sl" not in item: item["sl"] = {}
                item["sl"]["lt"] = new_cat
            
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
            
            # ⭐️ [關鍵修正] 這裡必須把 freight_target 和 summary_rows 存進去，不然顯示時會抓不到！
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
                
                # 👇 這裡是我幫您補上的，為了新的看板功能
                "freight_target": res_main.get("freight_target", 0),
                "summary_rows": res_main.get("summary_rows", []),
                
                "full_text_for_search": combined_input,
                "combined_input": combined_input
            }
            
            progress_bar.progress(1.0)
            status_box.update(label="✅ 分析完成！", state="complete", expanded=False)
            st.rerun()

            # --- 💡 [顯示結果區塊] 數量同步修正版 ---
    if st.session_state.analysis_result_cache:
        cache = st.session_state.analysis_result_cache
        all_issues = cache.get('all_issues', [])
        
        # 1. 頂部狀態條
        st.success(f"工令: {cache['job_no']} | ⏱️ {cache['total_duration']:.1f}s")
        st.info(f"💰 本次成本: NT$ {cache['cost_twd']:.2f} (In: {cache['total_in']:,} / Out: {cache['total_out']:,})")
        st.caption(f"細節耗時: Azure OCR {cache['ocr_duration']:.1f}s | AI 分析 {cache['time_eng']:.1f}s")

        # 2. 規則檢視
        with st.expander("🔍 檢視 Excel 規則與邏輯參數", expanded=False):
            rules_text = get_dynamic_rules(cache.get('full_text_for_search',''), debug_mode=True)
            st.markdown(rules_text)
                
        # 3. 原始數據檢視
        with st.expander("📊 檢視 AI 抄錄原始數據", expanded=False):
            st.markdown("**1. 核心指標摘要**")
            f_target = cache.get('freight_target', 0)
            sum_rows_len = len(cache.get("summary_rows", []))
            summary_df = pd.DataFrame([{
                "工令單號": cache.get("job_no", "N/A"),
                "運費 Target (PC)": f_target,
                "運費偵測狀態": "有抓到" if f_target > 0 else "未偵測",
                "總表行數": sum_rows_len,
                "總表狀態": "正常" if sum_rows_len > 0 else "空值"
            }])
            st.dataframe(summary_df, hide_index=True, use_container_width=True)
            st.divider()

            st.markdown("**2. 左上角統計表 (Summary Rows)**")
            sum_rows = cache.get("summary_rows", [])
            if sum_rows:
                df_sum = pd.DataFrame(sum_rows)
                df_sum.rename(columns={"title": "項目名稱", "target": "實交數量"}, inplace=True)
                st.dataframe(df_sum, hide_index=True, use_container_width=True)
            else:
                st.caption("無數據 (變數 summary_rows 為空)")
            st.divider()

            st.markdown("**3. 全卷詳細抄錄數據 (JSON)**")
            st.json(cache.get("ai_extracted_data", []), expanded=True)

        # 4. Python Debug
        with st.expander("🐍 Python 硬邏輯偵測結果", expanded=False):
            if cache.get('python_debug_data'):
                st.dataframe(cache['python_debug_data'], use_container_width=True, hide_index=True)
            else:
                st.caption("無偵測資料")

        # ========================================================
        # ⚡️ [修正重點]：先進行合併，再根據合併後的清單來計算數量
        # ========================================================
        
        # 1. 執行合併 (把 51 個異常壓縮成 N 類)
        consolidated_list = consolidate_issues(all_issues)

        # 2. 過濾出「真正的錯誤」 (排除僅是未匹配規則的警告)
        # 注意：我們是在 consolidated_list 上做篩選，這樣數量才會對
        real_errors_consolidated = [i for i in consolidated_list if "未匹配" not in i.get('issue_type', '')]

        # 3. 顯示結論 (使用合併後的數量)
        if not all_issues:
            st.balloons()
            st.success("✅ 全數合格！")
        elif not real_errors_consolidated:
            # 這裡用 len(consolidated_list) 代表還有幾個黃色警告
            st.success(f"✅ 數值合格！ (但有 {len(consolidated_list)} 類項目未匹配規則)")
        else:
            # 這裡顯示紅色的異常「類別」數量
            st.error(f"發現 {len(real_errors_consolidated)} 類異常")

        # 4. 卡片循環顯示 (使用合併後的清單)
        for item in consolidated_list:
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                source_label = item.get('source', '')
                issue_type = item.get('issue_type', '異常')
                
                # 頁碼顯示優化
                page_str = item.get('page', '?')
                if "," in str(page_str):
                    page_display = f"Pages: {page_str}"
                else:
                    page_display = f"P.{page_str}"

                c1.markdown(f"**{page_display} | {item.get('item')}** `{source_label}`")
                
                if any(kw in issue_type for kw in ["統計", "數量", "流程", "溯源"]):
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
