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

GLOBAL_FUZZ_THRESHOLD = 70


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
    
def python_engineering_audit(dimension_data):
    """
    Python 工程引擎 (新增：負責 Excel 強制分類與數值檢查)
    1. 這是原本我們要修改的邏輯，現在獨立出來，不與表頭檢查衝突。
    2. 負責執行：Range(再生), Un_regen(本體), Max, Min, Exempt(豁免)。
    """
    issues = []
    import re

    # 輔助：數值提取
    def get_val(val_str):
        clean_v = "".join(re.findall(r"[\d\.\-]+", str(val_str)))
        try: return float(clean_v)
        except: return None

    # 核心檢查迴圈
    for item in dimension_data:
        p_num = item.get("page", "?")
        title = item.get("item_title", "Unknown")
        ds_str = item.get("ds", "")
        
        # 1. 取得分類 (這裡會去呼叫我們等下要更新的 assign_category_by_python)
        # 這一步最關鍵！它會去讀 Excel 看有沒有強制規則
        final_category = assign_category_by_python(title)
        
        # 2. ⚡️ 豁免機制：若 Excel 設定為「豁免」，直接跳過
        if final_category == "exempt":
            continue

        # 3. 執行各類別檢查
        
        # A. Un_regen (本體未再生 - 強制整數檢查)
        if final_category == "un_regen":
            for pair in ds_str.split("|"):
                if ":" not in pair: continue
                rid, val_s = pair.split(":")[:2]
                val = get_val(val_s)
                
                if val is not None:
                    # 檢查是否為整數 (允許 0.05 誤差)
                    if abs(val - round(val)) > 0.05:
                         issues.append({
                            "page": p_num,
                            "item": title,
                            "issue_type": "⚠️異常(未再生)",
                            "common_reason": "應為整數 (Excel規則:本體未再生)",
                            "failures": [{"id": rid, "val": val, "calc": "非整數"}],
                            "source": "🐍 工程引擎"
                        })

        # B. Range (再生車修 - 區間檢查)
        elif final_category == "range":
            # 這裡您可以呼叫原本寫好的 check_range 邏輯
            # 或者暫時留空，至少它不會誤判成 "未再生"
            pass 

        # C. Max/Min Limit (軸頸/銲補)
        elif final_category == "max_limit" or final_category == "min_limit":
             # 這裡呼叫原本的 check_limit 邏輯
             pass 

    return issues

def assign_category_by_python(item_title):
    """
    Python 分類官 (v11: 熱處理/動平衡豁免版)
    1. [豁免]: 動平衡、熱處理 -> 直接 Exempt (不驗尺寸)。
    2. [既有功能]: 軸位/軸頸 Max Limit、SKIP 判斷等。
    """
    import pandas as pd
    from thefuzz import fuzz
    import re

    def clean_text(text):
        return str(text).replace(" ", "").replace("\n", "").replace("\r", "").replace('"', '').replace("'", "").strip()

    title_clean = clean_text(item_title)
    t = str(item_title).upper().replace(" ", "").replace("\n", "").replace('"', "")

    # ⚡️ [新增] 動平衡、熱處理直接豁免 (不驗尺寸，但會計照常)
    if any(k in t for k in ["動平衡", "BALANCING", "熱處理", "HEAT", "TREATING"]):
        return "exempt"

    try:
        df = pd.read_excel("rules.xlsx")
        df.columns = [c.strip() for c in df.columns]
        
        best_score = 0
        forced_rule = None
        
        for _, row in df.iterrows():
            rule_val = str(row.get('Category_Rule', '')).strip()
            if not rule_val or rule_val.lower() == 'nan': continue
            
            iname = str(row.get('Item_Name', '')).strip()
            iname_clean = clean_text(iname)
            
            score = fuzz.partial_ratio(iname_clean, title_clean)
            if score < 95: 
                 t_no = re.sub(r"[\(（].*?[\)）]", "", title_clean)
                 sc_no = fuzz.partial_ratio(iname_clean, t_no)
                 if sc_no > score: score = sc_no
            
            if score > 85: 
                if score > best_score:
                    best_score = score
                    forced_rule = rule_val
                elif score == best_score:
                    if len(rule_val) > len(forced_rule if forced_rule else ""):
                        forced_rule = rule_val

        if forced_rule:
            fr = forced_rule.upper()
            if "豁免" in fr or "EXEMPT" in fr or "SKIP" in fr: return "exempt"
            
            if "再生" in fr or "精車" in fr or "RANGE" in fr: return "range"
            if "銲" in fr or "焊" in fr or "MIN" in fr: return "min_limit"
            if "軸頸" in fr or "軸頭" in fr or "軸位" in fr or "MAX" in fr: return "max_limit"
            if "本體" in fr or "UN_REGEN" in fr: return "un_regen"
            
    except Exception: pass

    has_weld = any(k in t for k in ["銲補", "銲接", "焊", "WELD", "鉀"])
    has_unregen = any(k in t for k in ["未再生", "UN_REGEN", "粗車"])
    has_regen = any(k in t for k in ["再生", "研磨", "精加工", "車修", "KEYWAY", "GRIND", "MACHIN", "精車", "組裝", "拆裝", "裝配", "ASSY"])
    
    if has_weld: return "min_limit"
    if has_unregen:
        if any(k in t for k in ["軸頸", "軸頭", "軸位", "內孔", "JOURNAL"]): return "max_limit"
        return "un_regen"
    if has_regen: return "range"

    return "unknown"

def consolidate_issues(issues):
    """
    🗂️ 異常合併器：將「項目」、「錯誤類型」、「原因」完全相同的異常合併成一張卡片
    """
    grouped = {}
    for i in issues:
        key = (i.get('item', ''), i.get('issue_type', ''), i.get('common_reason', ''))
        if key not in grouped:
            grouped[key] = i.copy()
            grouped[key]['pages_set'] = {str(i.get('page', '?'))}
            grouped[key]['failures'] = i.get('failures', []).copy()
        else:
            grouped[key]['pages_set'].add(str(i.get('page', '?')))
            grouped[key]['failures'].extend(i.get('failures', []))
            
    result = []
    for key, val in grouped.items():
        sorted_pages = sorted(list(val['pages_set']), key=lambda x: int(x) if x.isdigit() else 999)
        val['page'] = ", ".join(sorted_pages)
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
    
    4. **特殊批量總數提取 (batch_total_qty)：
       - 若標題包含「熱處理」、「研磨」、「動平衡」且內文第一欄為合併儲存格顯示總量 (如 2425KG, 8293.80 IN2)：
       - 請將該數值提取至 JSON 的 "batch_total_qty" 欄位 (純數字)。
         (注意：研磨與動平衡若後續還有個別 ID 與尺寸，請照常抄錄到 "ds"。)

    5. **分類 (category)**：**請直接回傳 `null`**。由後端程式判定。

    6. **數據抄錄 (ds) 與 字串保護規範**：
       - **格式**：輸出為 `"ID:值|ID:值"` 的字串格式。
       - **禁止簡化**：實測值若顯示 `349.90`，必須輸出 `"349.90"`，保留尾數 0。
       - **🚫 遇到干擾不鑽牛角尖**：若儲存格內的數值因手寫塗改、圓圈遮擋、污點、字跡黏連或光線反光，導致你無法「100% 確定」原始打印數字時，**嚴禁腦補或猜測**。
       - **壞軌標記 [BAD]**：請將該筆數值直接標記為 `[!]`。
       - **範例**：若 ID 清楚但數值模糊 -> `"V100:[!]"`；若整個儲存格都看不清 -> `"[!] : [!]"`。
       - **跳過策略**：一旦標記為 `[!]`，請立即跳到下一格，不要浪費 Token 描述雜訊。

    #### 💰 模組 B：會計指標提取 (AI 任務：抄錄)
    ⚠️ **注意範圍**：你只能從標記為 `=== [SUMMARY_TABLE (總表)] ===` 的區域提取數據。
    1. **統計表**：請提取每一行的以下三個欄位：
       - **項目名稱 (title)**
       - **申請數量 (apply_qty)**：通常在左側。
       - **實交數量 (delivery_qty)**：通常在右側 (這是會計核對的基準)。
       
    2. **頁碼標註**：請務必在每個 `summary_rows` 物件中記錄該行所在的頁碼 (`page`)。

    #### 📋 模組 C：表頭資訊 (Header Info)
    ⚠️ **注意範圍**：你只能從標記為 `=== [SUMMARY_TABLE (總表)] ===` 的區域提取數據。
    1. **工令單號 (job_no)**：通常是 10 碼，由英文字母 (W, R, O, Y) 開頭，並且不會含超過3個英文字母。
    2. **預定交貨日 (scheduled_date)**：請將日期統一格式化為 "YYYY/MM/DD"。
    3. **實際交貨日 (actual_date)**：請將日期統一格式化為 "YYYY/MM/DD"。

    ### 📝 輸出規範 (Output Format)
    必須回傳單一 JSON。

    {{
      "header_info": {{
          "job_no": "Wxxxxxxxxx",
          "scheduled_date": "YYYY/MM/DD",
          "actual_date": "YYYY/MM/DD"
      }},
      "summary_rows": [ 
          {{ "page": 頁碼, "title": "名", "apply_qty": 數字, "delivery_qty": 數字 }} 
      ], 
      "issues": [], 
      "dimension_data": [
         {{
           "page": 數字, "item_title": "標題", "batch_total_qty": 0, "category": null, 
           "item_pc_target": 0,
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
    """
    Python 工程引擎 (v29: 全域統一特規版)
    升級內容：
    1. [統一配對]: 引入與會計同級的配對邏輯 (GLOBAL_FUZZ_THRESHOLD + fuzz.ratio)。
    2. [規則優先]: 若 Excel 特規配對成功且設定為 SKIP/EXEMPT，直接豁免。
    3. [原有邏輯]: 保留熱處理/動平衡豁免，以及各種數值檢查邏輯。
    """
    grouped_errors = {}
    import re
    import pandas as pd
    from thefuzz import fuzz

    # 🔥 1. 讀取全域門檻 (與會計同步)
    CURRENT_THRESHOLD = globals().get('GLOBAL_FUZZ_THRESHOLD', 95)

    if not dimension_data: return []

    # 🔥 2. 預先載入規則 (只載入一次)
    rules_map = {}
    try:
        df = pd.read_excel("rules.xlsx")
        df.columns = [c.strip() for c in df.columns]
        for _, row in df.iterrows():
            iname = str(row.get('Item_Name', '')).strip()
            if iname: 
                # 工程主要看 Local 規則 (是否豁免)
                rules_map[str(iname).replace(" ", "").replace("\n", "").strip()] = {
                    "u_local": str(row.get('Unit_Rule_Local', '')).strip()
                }
    except: pass

    for item in dimension_data:
        ds = str(item.get("ds", ""))
        if not ds: continue
        raw_entries = [p.split(":") for p in ds.split("|") if ":" in p]
        
        title = str(item.get("item_title", "")).replace(" ", "").replace('"', "")
        cat = str(item.get("category", "")).strip()
        page_num = item.get("page", "?")
        raw_spec = str(item.get("std_spec", "")).replace('"', "")
        
        # =========================================================
        # 🔥 3. 執行特規配對 (統一邏輯)
        # =========================================================
        title_clean = title.strip()
        rule_set = None
        
        # A. 完全匹配
        if title_clean in rules_map:
            rule_set = rules_map[title_clean]
        
        # B. 去括號匹配
        if not rule_set:
            t_no = re.sub(r"[\(（].*?[\)）]", "", title_clean)
            if t_no in rules_map:
                rule_set = rules_map[t_no]
        
        # C. 模糊匹配 (使用全域門檻 + 嚴格比對)
        if not rule_set and rules_map:
            best_score = 0
            for k, v in rules_map.items():
                sc = fuzz.token_sort_ratio(k, title_clean) # 嚴格比對
                if sc > CURRENT_THRESHOLD and sc > best_score:
                    best_score = sc
                    rule_set = v
        # =========================================================

        # ⚡️ [既有豁免] 動平衡、熱處理直接跳過 (關鍵字優先)
        t_upper = title.upper()
        if any(k in t_upper for k in ["動平衡", "BALANCING", "熱處理", "HEAT"]):
            continue
            
        # ⚡️ [規則豁免] 如果 Excel 規則說要 SKIP，就跳過
        if rule_set:
            u_local = rule_set.get("u_local", "").upper()
            if "SKIP" in u_local or "EXEMPT" in u_local or "豁免" in u_local:
                continue

        # --- 以下為數值提取與檢查邏輯 (保持 v28 原貌) ---
        
        mm_nums = [float(n) for n in re.findall(r"(\d+\.?\d*)\s*mm", raw_spec)]
        all_nums = [float(n) for n in re.findall(r"(\d+\.?\d*)", raw_spec)]
        noise = [350.0, 300.0, 200.0, 145.0, 130.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0]
        clean_std = [n for n in all_nums if (n in mm_nums) or (n not in noise and n > 5)]

        s_ranges = []
        spec_parts = re.split(r"[\n\r]|[一二三四五六]|[（(]\d+[)）]|[;；]", raw_spec)
        
        for part in spec_parts:
            clean_part = part.replace(" ", "").replace("\n", "").replace("mm", "").replace("MM", "").strip()
            if not clean_part: continue
            
            pm_matches = list(re.finditer(r"(\d+\.?\d*)?±(\d+\.?\d*)", clean_part))
            if pm_matches:
                for match in pm_matches:
                    base_str, offset_str = match.group(1), match.group(2)
                    b = float(base_str) if base_str else 0.0
                    o = float(offset_str)
                    s_ranges.append([round(b - o, 4), round(b + o, 4)])
                continue 

            tilde_matches = list(re.finditer(r"(\d+\.?\d*)[~～-](\d+\.?\d*)", clean_part))
            has_valid_tilde = False
            if tilde_matches:
                for match in tilde_matches:
                    n1, n2 = float(match.group(1)), float(match.group(2))
                    if abs(n1 - n2) < n1 * 0.5:
                        s_ranges.append([round(min(n1, n2), 4), round(max(n1, n2), 4)])
                        has_valid_tilde = True
            
            if has_valid_tilde: continue

            all_numbers = re.findall(r"[-+]?\d+\.?\d*", clean_part)
            if not all_numbers: continue

            try:
                bases = []
                offsets = []
                for token in all_numbers:
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
            except: continue
                    
        logic = item.get("sl", {})
        l_type = logic.get("lt", "")
        
        # 4. 預算基準
        if "SKIP" in l_type.upper() or "EXEMPT" in l_type.upper() or "豁免" in l_type:
            un_regen_target = None
            
        elif l_type in ["range", "max_limit", "min_limit"]:
            un_regen_target = None
            
        else:
            s_threshold = logic.get("t", 0)
            un_regen_target = None
            if l_type in ["un_regen", "未再生"] or ("未再生" in (cat + title) and not any(k in (cat + title) for k in ["軸頸", "軸頭", "軸位"])):
                cands = [n for n in clean_std if n >= 120.0]
                if s_threshold and float(s_threshold) >= 120.0: cands.append(float(s_threshold))
                if cands: un_regen_target = max(cands)

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

                if "SKIP" in l_type.upper() or "EXEMPT" in l_type.upper():
                    continue

                elif "min_limit" in l_type or "銲補" in (cat + title):
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

                elif l_type == "max_limit" or (any(k in (cat + title) for k in ["軸頸", "軸頭", "軸位"]) and ("未再生" in (cat + title))):
                    engine_label = "軸頸(上限)"
                    candidates = clean_std
                    target = max(candidates) if candidates else 0
                    t_used = target
                    if target > 0:
                        if not is_pure_int: is_passed, reason = False, "應為純整數"
                        elif val > target: is_passed, reason = False, f"超過上限 {target}"

                elif l_type == "range" or (any(x in (cat + title) for x in ["再生", "精加工", "研磨", "車修", "組裝", "拆裝", "真圓度"]) and "未再生" not in (cat + title)):
                    engine_label = "精加工"
                    if not is_two_dec:
                        is_passed, reason = False, "應填兩位小數"
                    elif s_ranges:
                        t_used = str(s_ranges)
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
    Python 會計官 (v53: 全域特規模糊比對版)
    修改重點：
    1. [全域連動]: 不再使用寫死的 FUZZ_THRESHOLD。
       - 改為讀取 globals().get('GLOBAL_FUZZ_THRESHOLD', 90)。
       - 讓會計、工程、流程能統一使用外部設定的門檻。
    2. [功能保留]: 
       - 執行高門檻比對 (預設使用全域設定)。
       - 將命中紀錄打包回傳 (HIDDEN_DATA)。
       - 核心籃子邏輯維持不變。
    """
    accounting_issues = []
    from thefuzz import fuzz
    from collections import Counter
    import re
    import pandas as pd 

    # --- 0. 設定 (改為讀取全域變數) ---
    # 嘗試讀取全域設定，如果沒設定則預設為 90 (依您提供的代碼預設值)
    CURRENT_THRESHOLD = globals().get('GLOBAL_FUZZ_THRESHOLD', 90)

    def clean_text(text):
        return str(text).replace(" ", "").replace("\n", "").replace("\r", "").replace('"', '').replace("'", "").strip()

    def safe_float(value):
        if value is None or str(value).upper() == 'NULL': return 0.0
        if "[!]" in str(value): return "BAD_DATA" 
        cleaned = "".join(re.findall(r"[\d\.]+", str(value).replace(',', '')))
        try: return float(cleaned) if cleaned else 0.0
        except: return 0.0

    def parse_ratio(rule_str):
        if not rule_str: return 1.0
        match = re.search(r"(\d+)\s*/\s*(\d+)", str(rule_str))
        if match:
            n, d = float(match.group(1)), float(match.group(2))
            if d != 0: return n / d
        return 1.0

    # --- 1. 載入規則 ---
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

    summary_rows = res_main.get("summary_rows", [])
    
    # 🔥 特規命中紀錄器
    rule_hits_log = {} 

    # =================================================
    # 🕵️‍♂️ 第一關：總表內戰
    # =================================================
    global_sum_tracker = {}
    for s in summary_rows:
        s_title = s.get('title', 'Unknown')
        q_apply = safe_float(s.get('apply_qty', 0))      
        q_deliver = safe_float(s.get('delivery_qty', 0)) 
        if q_deliver == 0 and 'target' in s: q_deliver = safe_float(s.get('target', 0))

        if abs(q_apply - q_deliver) > 0.01:
             accounting_issues.append({
                "page": s.get('page', "總表"), 
                "item": f"{s_title}", 
                "issue_type": "🚨 總表數量異常", 
                "common_reason": f"申請({q_apply}) != 實交({q_deliver})", 
                "failures": [
                    {"頁碼": "總表", "項目名稱": "📝 申請數量", "數量": q_apply, "備註": "原始值"},
                    {"頁碼": "總表", "項目名稱": "🚛 實交數量", "數量": q_deliver, "備註": "核對值"}
                ], 
                "source": "🐍 會計引擎"
            })
        global_sum_tracker[s_title] = {
            "target": q_deliver, "actual": 0, "details": [], "page": s.get('page', "總表")
        }

    # =================================================
    # 🕵️‍♂️ 第二關：逐項掃描
    # =================================================
    for item in dimension_data:
        raw_title = item.get("item_title", "")
        title_clean = clean_text(raw_title) 
        page = item.get("page", "?")
        target_pc = safe_float(item.get("item_pc_target", 0)) 
        batch_qty = safe_float(item.get("batch_total_qty", 0))
        
        # 2.1 規則匹配 (紀錄邏輯)
        rule_set = None
        matched_rule_name = None
        match_type = ""
        match_score = 0

        # A. 完全匹配
        if title_clean in rules_map:
            rule_set = rules_map[title_clean]
            matched_rule_name = title_clean
            match_type = "完全匹配"
            match_score = 100
        
        # B. 去括號匹配
        if not rule_set:
            t_no = re.sub(r"[\(（].*?[\)）]", "", title_clean)
            if t_no in rules_map:
                rule_set = rules_map[t_no]
                matched_rule_name = t_no
                match_type = "去括號匹配"
                match_score = 100

        # C. 模糊匹配 (使用全域變數 CURRENT_THRESHOLD)
        if not rule_set and rules_map:
            best_score = 0
            best_rule = None
            for k, v in rules_map.items():
                sc = fuzz.token_sort_ratio(k, title_clean) 
                # 🔥 改用 CURRENT_THRESHOLD
                if sc > CURRENT_THRESHOLD and sc > best_score:
                    best_score = sc
                    rule_set = v
                    best_rule = k
            
            if rule_set:
                matched_rule_name = best_rule
                match_type = "模糊匹配"
                match_score = best_score
        
        # 記錄命中
        if matched_rule_name:
            if matched_rule_name not in rule_hits_log:
                rule_hits_log[matched_rule_name] = []
            
            rule_hits_log[matched_rule_name].append({
                "明細名稱": raw_title,
                "匹配類型": match_type,
                "分數": match_score,
                "頁碼": page
            })

        # --- 以下為既有邏輯 ---
        u_local = rule_set.get("u_local", "") if rule_set else ""
        u_fr = rule_set.get("u_fr", "") if rule_set else ""
        u_agg = rule_set.get("u_agg", "") if rule_set else ""
        
        ds = str(item.get("ds", ""))
        data_list = [pair.split(":") for pair in ds.split("|") if ":" in pair]
        raw_count = len(data_list) if data_list else 0
        id_counts = Counter([str(e[0]).strip() for e in data_list if len(e)>0])

        # A. 單項檢查
        is_local_exempt = "豁免" in str(u_local) or "SKIP" in str(u_local).upper() or "EXEMPT" in str(u_local).upper()
        actual_item_qty = raw_count if batch_qty > 0 else raw_count * parse_ratio(u_local)
        if not is_local_exempt and abs(actual_item_qty - target_pc) > 0.01 and target_pc > 0:
             accounting_issues.append({
                 "page": page, "item": raw_title, "issue_type": "🛑 統計不符(單項)", 
                 "common_reason": f"標題 {target_pc} != 內文 {actual_item_qty}", 
                 "failures": [], "source": "🐍 會計引擎"
             })

        # B. 重複檢查
        journal_family = ["軸頸", "軸頭", "軸位", "內孔", "JOURNAL"]
        if "本體" in title_clean:
             for rid, count in id_counts.items():
                if count > 1: accounting_issues.append({"page": page, "item": raw_title, "issue_type": "⚠️編號重複(本體)", "common_reason": f"{rid} 重複 {count}次", "failures": []})
        elif any(k in title_clean for k in journal_family):
             for rid, count in id_counts.items():
                if count > 2: accounting_issues.append({"page": page, "item": raw_title, "issue_type": "⚠️編號重複(軸頸)", "common_reason": f"{rid} 重複 {count}次", "failures": []})

        # C. 運費 & 歸戶
        fr_multiplier = parse_ratio(u_fr)
        freight_val = 0.0
        f_note = ""
        u_fr_upper = str(u_fr).upper()
        is_fr_exempt = "豁免" in u_fr_upper or "SKIP" in u_fr_upper
        is_forced_include = "計入" in str(u_fr) or "INCLUDED" in u_fr_upper
        is_default_target = ("本體" in title_clean and "未再生" in title_clean) or ("新品組裝" in title_clean)
        
        if not is_fr_exempt and (is_default_target or is_forced_include or fr_multiplier != 1.0):
            freight_val = actual_item_qty * fr_multiplier
            f_note = f"x{fr_multiplier}" if fr_multiplier != 1.0 else ""

        # 確定 Agg Mode
        agg_mode = "B" 
        if u_agg:
            p_clean = str(u_agg).upper().replace(" ", "")
            if "EXEMPT" in p_clean or "SKIP" in p_clean: agg_mode = "EXEMPT"
            elif "AB" in p_clean: agg_mode = "AB"
            elif "A" in p_clean: agg_mode = "A"

        agg_multiplier = parse_ratio(u_agg)
        qty_agg = batch_qty if batch_qty > 0 else actual_item_qty * agg_multiplier

        if agg_mode != "EXEMPT":
            for s_title, data in global_sum_tracker.items():
                s_clean = clean_text(s_title)
                
                if (fuzz.partial_ratio("輥輪拆裝.車修或銲補運費", s_clean) > 70) or ("運費" in s_clean):
                    if freight_val > 0:
                        data["actual"] += freight_val
                        data["details"].append({"page": page, "title": raw_title, "val": freight_val, "note": f"運費 {f_note}"})
                    continue

                # =========================================================
                # 🧺 步驟 1: 籃子撈人 (v52)
                # =========================================================
                match_A = (fuzz.partial_ratio(s_clean, title_clean) > 90)
                match_B = False
                
                s_upper_check = s_clean.upper() 

                is_dis = fuzz.partial_ratio("ROLL拆裝", s_upper_check) > 80
                is_mac = fuzz.partial_ratio("ROLL車修", s_upper_check) > 80
                is_weld = (fuzz.partial_ratio("ROLL銲補", s_upper_check) > 80) or \
                          ("焊" in s_upper_check) or \
                          ("鉀" in s_upper_check)
                
                has_part_body = "本體" in title_clean
                has_part_journal = any(k in title_clean for k in journal_family)
                
                # 白名單還原: 只保留嚴格動作
                has_act_mac = any(k in title_clean for k in ["再生", "精車", "未再生", "粗車"])
                
                has_act_weld = ("銲補" in title_clean or "焊" in title_clean or "鉀" in title_clean)
                is_assy = ("組裝" in title_clean or "拆裝" in title_clean)
                
                if is_dis and is_assy: match_B = True
                elif is_mac and (has_part_body or has_part_journal) and has_act_mac: match_B = True
                elif is_weld and (has_part_body or has_part_journal) and has_act_weld: match_B = True
                
                if agg_mode == "A": match = match_A
                elif agg_mode == "AB": match = match_A or match_B
                else: match = match_B if match_B else match_A

                # =========================================================
                # 🛑 步驟 2: 攔截者
                # =========================================================
                if match:
                    s_upper = s_clean.upper()
                    t_upper = title_clean.upper()
                    
                    s_is_unregen = "未再生" in s_clean or "粗車" in s_clean
                    t_is_unregen = "未再生" in title_clean or "粗車" in title_clean
                    s_is_regen = ("再生" in s_clean or "精車" in s_clean) and not s_is_unregen
                    t_is_regen = ("再生" in title_clean or "精車" in title_clean or "車修" in title_clean) and not t_is_unregen
                    
                    s_is_body = "本體" in s_clean
                    t_is_body = "本體" in title_clean
                    s_is_journal = any(k in s_clean for k in journal_family)
                    t_is_journal = any(k in title_clean for k in journal_family)

                    if s_is_regen and t_is_unregen: match = False
                    if s_is_unregen and t_is_regen: match = False
                    if s_is_body and not s_is_journal and t_is_journal: match = False
                    if s_is_journal and not s_is_body and t_is_body: match = False
                    if "TOP" in s_upper and "BOTTOM" in t_upper: match = False
                    if "BOTTOM" in s_upper and "TOP" in t_upper: match = False

                if match:
                    data["actual"] += qty_agg
                    c_msg = f"x{agg_multiplier}" if agg_multiplier != 1.0 else ""
                    data["details"].append({"page": page, "title": raw_title, "val": qty_agg, "note": c_msg})

    # =================================================
    # 🕵️‍♂️ 第三關：明細總結算
    # =================================================
    for s_title, data in global_sum_tracker.items():
        if abs(data["actual"] - data["target"]) > 0.01: 
            fail_table = []
            fail_table.append({"頁碼": "總表", "項目名稱": f"🎯 目標 (實交)", "數量": data["target"], "備註": "基準"})
            for d in data["details"]:
                fail_table.append({"頁碼": f"P.{d['page']}", "項目名稱": d['title'], "數量": d['val'], "備註": d['note']})
            fail_table.append({"頁碼": "∑", "項目名稱": "加總結果", "數量": data["actual"], "備註": "總計"})

            accounting_issues.append({
                "page": data["page"], "item": s_title, 
                "issue_type": "🛑 明細匯總不符", 
                "common_reason": f"實交({data['target']}) != 明細加總({data['actual']})", 
                "failures": fail_table, "source": "🐍 會計引擎"
            })
            
    # 🔥🔥🔥 [關鍵]: 將命中資料當作一個隱藏的 ISSUE 回傳 (TYPE=HIDDEN_DATA)
    if rule_hits_log:
        accounting_issues.append({
            "issue_type": "HIDDEN_DATA",
            "rule_hits": rule_hits_log,
            "fuzz_threshold": CURRENT_THRESHOLD # 🔥 顯示目前實際使用的門檻
        })
            
    return accounting_issues

def python_process_audit(dimension_data):
    """
    Python 流程引擎 (v24: 全域統一特規版)
    升級內容：
    1. [統一配對]: 引入與會計/工程同級的配對邏輯 (GLOBAL_FUZZ_THRESHOLD + fuzz.ratio)。
       - 徹底解決 "規則劫持" 導致的錯誤工序判定。
    2. [規則優先]: 若 Excel 特規配對成功且設定為 SKIP/EXEMPT，直接跳過檢查。
    3. [既有功能]: 保留熱處理/動平衡關鍵字排除、工序溯源、尺寸邏輯。
    """
    process_issues = []
    import re
    import pandas as pd
    from thefuzz import fuzz

    # 🔥 1. 讀取全域門檻 (與會計/工程同步)
    CURRENT_THRESHOLD = globals().get('GLOBAL_FUZZ_THRESHOLD', 95)

    def clean_text(text):
        return str(text).replace(" ", "").replace("\n", "").replace("\r", "").replace('"', '').replace("'", "").strip()

    # 2. 載入規則
    rules_map = {}
    try:
        df = pd.read_excel("rules.xlsx")
        df.columns = [c.strip() for c in df.columns]
        for _, row in df.iterrows():
            iname = str(row.get('Item_Name', '')).strip()
            p_rule = str(row.get('Process_Rule', '')).strip()
            # 流程引擎主要看 Process_Rule
            if iname and p_rule and p_rule.lower() != 'nan':
                rules_map[clean_text(iname)] = p_rule
    except: pass

    STAGE_MAP = { 1: "未再生/粗車", 2: "銲補/焊補", 3: "再生/精車", 4: "研磨" }
    history = {} 

    if not dimension_data: return []

    for item in dimension_data:
        p_num = item.get("page", "?")
        title = str(item.get("item_title", "")).strip()
        title_clean = clean_text(title)
        ds = str(item.get("ds", ""))
        
        # ⚡️ [既有豁免] 動平衡、熱處理直接跳過流程檢查 (關鍵字優先)
        t_upper = title_clean.upper()
        if any(k in t_upper for k in ["動平衡", "BALANCING", "熱處理", "HEAT"]):
            continue

        # =========================================================
        # 🔥 3. 執行特規配對 (統一邏輯)
        # =========================================================
        forced_rule = None
        
        # A. 完全匹配
        if title_clean in rules_map:
            forced_rule = rules_map[title_clean]
        
        # B. 去括號匹配
        if not forced_rule:
            t_no = re.sub(r"[\(（].*?[\)）]", "", title_clean)
            if t_no in rules_map:
                forced_rule = rules_map[t_no]

        # C. 模糊匹配 (使用全域門檻 + 嚴格比對)
        if not forced_rule and rules_map:
            best_score = 0
            for k, v in rules_map.items():
                sc = fuzz.token_sort_ratio(k, title_clean) # 嚴格比對 (原為 partial_ratio)
                if sc > CURRENT_THRESHOLD and sc > best_score:
                    best_score = sc
                    forced_rule = v
        # =========================================================

        track = "Unknown"
        stage = 0
        
        # 如果配對到規則，解析規則內容
        if forced_rule:
            fr = forced_rule.upper()
            # ⚡️ [規則豁免] 如果規則說 SKIP，跳過
            if "豁免" in fr or "EXEMPT" in fr or "SKIP" in fr: 
                continue 
            
            if "本體" in fr: track = "本體"
            elif "軸頸" in fr or "軸頭" in fr or "軸位" in fr: track = "軸頸"
            
            if "未再生" in fr or "粗車" in fr: stage = 1
            elif "銲" in fr or "焊" in fr or "鉀" in fr: stage = 2
            elif "再生" in fr or "精車" in fr: stage = 3
            elif "研磨" in fr: stage = 4

        # 如果規則沒指定(或沒配到)，使用預設關鍵字判斷
        if stage == 0:
            if "研磨" in title: stage = 4
            elif any(k in title for k in ["銲補", "銲接", "焊", "鉀"]): stage = 2
            elif "未再生" in title or "粗車" in title: stage = 1
            elif "再生" in title or "精車" in title: stage = 3

        if track == "Unknown":
            if "本體" in title: track = "本體"
            elif any(k in title for k in ["軸頸", "軸頭", "軸位", "內孔", "JOURNAL"]): track = "軸頸"
        
        if track == "Unknown" or stage == 0: continue 

        # --- 以下為數值收集邏輯 (保持不變) ---
        segments = ds.split("|")
        for seg in segments:
            parts = seg.split(":")
            if len(parts) < 2: continue
            rid = parts[0].strip().upper()
            val_str = parts[1].strip()
            nums = re.findall(r"\d+\.?\d*", val_str)
            if not nums: continue
            val = float(nums[0])
            
            key = (rid, track)
            if key not in history: history[key] = {}
            history[key][stage] = {
                "val": val, "page": p_num, "title": title
            }

    # --- 以下為檢查邏輯 (缺漏工序 + 尺寸倒置) 保持不變 ---
    for (rid, track), stages_data in history.items():
        present_stages = sorted(stages_data.keys())
        if not present_stages: continue
        max_stage = present_stages[-1]
        
        missing_stages = []
        for req_s in range(1, max_stage):
            if req_s not in stages_data: missing_stages.append(STAGE_MAP[req_s])
        
        if missing_stages:
            last_info = stages_data[max_stage]
            process_issues.append({
                "page": last_info['page'],
                "item": f"{last_info['title']}",
                "issue_type": "🛑溯源異常(缺漏工序)",
                "common_reason": f"[{track}] 進度至【{STAGE_MAP[max_stage]}】，缺前置：{', '.join(missing_stages)}",
                "failures": [{"id": rid, "val": "缺漏", "calc": "履歷不完整"}],
                "source": "🐍 流程引擎"
            })

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
                    process_issues.append({
                        "page": info_b['page'],
                        "item": f"[{track}] 尺寸邏輯",
                        "issue_type": "🛑流程異常(尺寸倒置)",
                        "common_reason": f"尺寸邏輯錯誤：{STAGE_MAP[s_a]} 應 {sign} {STAGE_MAP[s_b]}",
                        "failures": [{"id": STAGE_MAP[s_a], "val": info_a['val'], "calc": "前"}, {"id": STAGE_MAP[s_b], "val": info_b['val'], "calc": "後"}],
                        "source": "🐍 流程引擎"
                    })

    return process_issues
    
def python_header_audit_batch(photo_gallery, ai_res_json):
    """
    Python 表頭稽核官 (Batch 架構適配版 v30)
    1. [Raw Text] 掃描每一頁 OCR 文字，檢查工令是否混單 (Regex)。
    2. [AI JSON] 檢查 AI 讀出的工令格式 (10碼)。
    3. [AI JSON] 檢查日期邏輯 (實際 <= 預定)。
    """
    header_issues = []
    import re
    from datetime import datetime

    # --- 1. 混單檢查 (利用 OCR 原始文字) ---
    # 策略：直接用 Regex 在每一頁的文字裡撈 W/R/O/Y 開頭的字串
    job_pattern = r"([WROY][A-Z0-9]{9})" # 抓 10 碼
    found_jobs_map = {} # { "工令號": [頁碼list] }

    for idx, item in enumerate(photo_gallery):
        txt = item.get('full_text', '').upper().replace(" ", "").replace("-", "")
        # 尋找所有疑似工令的字串
        matches = re.findall(job_pattern, txt)
        for job in matches:
            if job not in found_jobs_map: found_jobs_map[job] = []
            found_jobs_map[job].append(idx + 1)

    # 如果找到多種不同的工令 -> 報警
    if len(found_jobs_map) > 1:
        details = [f"{k} (P.{v})" for k, v in found_jobs_map.items()]
        header_issues.append({
            "page": "多頁", "item": "工令單號", "issue_type": "🚨 嚴重混單",
            "common_reason": f"偵測到多種工令：{', '.join(details)}",
            "failures": [{"id": "內容", "val": str(found_jobs_map)}],
            "source": "🐍 表頭稽核(OCR)"
        })

    # --- 2. 格式與日期檢查 (利用 AI JSON) ---
    h_info = ai_res_json.get("header_info", {})
    
    # 工令格式 (針對 AI 最終認定的那一組)
    ai_job = h_info.get("job_no", "Unknown")
    if ai_job and ai_job != "Unknown":
        clean_job = ai_job.upper().replace(" ", "").replace("-", "")
        if not re.match(r"^[WROY][A-Z0-9]{9}$", clean_job):
            header_issues.append({
                "page": "表頭", "item": "工令格式", "issue_type": "⚠️ 格式錯誤",
                "common_reason": f"AI 識別工令 {ai_job} 格式不符 (需10碼，W/R/O/Y開頭)",
                "failures": [{"id": "識別值", "val": ai_job}],
                "source": "🐍 表頭稽核(AI)"
            })

    # 日期邏輯 (實際 <= 預定)
    d_sch = h_info.get("scheduled_date", "Unknown")
    d_act = h_info.get("actual_date", "Unknown")
    
    if d_sch != "Unknown" and d_act != "Unknown":
        try:
            # 嘗試解析 YYYY/MM/DD
            dt_sch = datetime.strptime(d_sch.replace("-", "/"), "%Y/%m/%d")
            dt_act = datetime.strptime(d_act.replace("-", "/"), "%Y/%m/%d")
            
            if dt_act > dt_sch:
                 header_issues.append({
                    "page": "表頭", "item": "交貨時效", "issue_type": "⏰ 逾期交貨",
                    "common_reason": f"實際 {d_act} 晚於 預定 {d_sch}",
                    "failures": [{"id": "延遲天數", "val": f"{(dt_act - dt_sch).days} 天"}], 
                    "source": "🐍 表頭稽核(AI)"
                })
        except:
            pass # 日期格式讀不懂，跳過

    return header_issues
    
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
        # 這裡記得維持我們上次改的 xlsm 支援
        uploaded_xlsx = st.file_uploader("上傳 Excel 檔", type=['xlsx', 'xls', 'xlsm'], key="xlsx_uploader")
        
        if uploaded_xlsx:
            try:
                current_file_name = uploaded_xlsx.name
                if st.session_state.get('last_loaded_xlsx_name') != current_file_name:
                    # 1. 讀取 Excel (header=None 保持不變)
                    df_dict = pd.read_excel(uploaded_xlsx, sheet_name=None, header=None)
                    
                    st.session_state.photo_gallery = []
                    st.session_state.source_mode = 'excel'
                    st.session_state.last_loaded_xlsx_name = current_file_name
                    
                    for sheet_name, df in df_dict.items():
                        df = df.fillna("")
                        
                        # 🔥🔥🔥 [新增這段：暴力壓平換行符號] 🔥🔥🔥
                        # 這行指令會把所有格子裡的 "\n" (換行) 替換成 " " (空格)
                        # 這樣 "W3...\n本體..." 就會變成 "W3... 本體..." (同一行)
                        df = df.astype(str).replace(r'\n', ' ', regex=True).replace(r'\r', ' ', regex=True)
                        
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
        # 強制清除上一筆
        st.session_state.analysis_result_cache = None 
        st.session_state.auto_start_analysis = False
        total_start = time.time()
        
        with st.status("總稽核官正在進行全方位分析...", expanded=True) as status_box:
            progress_bar = st.progress(0)
            
            # 1. OCR
            status_box.write("👀 正在進行 OCR 文字識別...")
            ocr_start = time.time()
            
            def process_task(index, item):
                if item.get('full_text'): return index, item.get('header_text',''), item['full_text'], None
                try:
                    item['file'].seek(0)
                    _, h, f, _, _ = extract_layout_with_azure(item['file'], DOC_ENDPOINT, DOC_KEY)
                    return index, h, f, None
                except Exception as e: return index, None, None, str(e)

            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                futures = [executor.submit(process_task, i, item) for i, item in enumerate(st.session_state.photo_gallery)]
                for future in concurrent.futures.as_completed(futures):
                    idx, h_txt, f_txt, err = future.result()
                    if not err:
                        st.session_state.photo_gallery[idx].update({'header_text': h_txt, 'full_text': f_txt, 'file': None})
                    progress_bar.progress(0.4 * ((idx + 1) / len(st.session_state.photo_gallery)))

            ocr_duration = time.time() - ocr_start
            
            # 2. 組合文字
            combined_input = ""
            for i, p in enumerate(st.session_state.photo_gallery):
                combined_input += f"\n=== Page {i+1} ===\n{p.get('full_text','')}\n"

                        # ... (上面是 2. 組合文字 combined_input，不用動) ...

            # 3. AI 分析 (加入計時)
            status_box.write("🤖 AI 正在全卷分析...")
            
            ai_start_time = time.time()  # ⏱️ [計時開始] AI
            res_main = agent_unified_check(combined_input, combined_input, GEMINI_KEY, main_model_name)
            ai_duration = time.time() - ai_start_time # ⏱️ [計時結束] AI
            
            progress_bar.progress(0.8)
            
            # 4. Python 邏輯檢查 (加入計時)
            status_box.write("🐍 Python 正在進行邏輯比對...")
            
            py_start_time = time.time() # ⏱️ [計時開始] Python
            
            dim_data = res_main.get("dimension_data", [])
            for item in dim_data:
                new_cat = assign_category_by_python(item.get("item_title", ""))
                item["category"] = new_cat
                if "sl" not in item: item["sl"] = {}
                item["sl"]["lt"] = new_cat
            
            python_numeric_issues = python_numerical_audit(dim_data)
            python_accounting_issues = python_accounting_audit(dim_data, res_main)
            python_process_issues = python_process_audit(dim_data)
            python_header_issues = python_header_audit_batch(st.session_state.photo_gallery, res_main)

            ai_filtered_issues = []
            ai_raw_issues = res_main.get("issues", [])
            if isinstance(ai_raw_issues, list):
                for i in ai_raw_issues:
                    if isinstance(i, dict):
                        i['source'] = '🤖 總稽核 AI'
                        if not any(k in i.get("issue_type", "") for k in ["流程", "規格提取失敗", "未匹配"]):
                            ai_filtered_issues.append(i)

            all_issues = ai_filtered_issues + python_numeric_issues + python_accounting_issues + python_process_issues + python_header_issues
            
            py_duration = time.time() - py_start_time # ⏱️ [計時結束] Python

            # 5. 存檔 (Cache)
            usage = res_main.get("_token_usage", {"input": 0, "output": 0})
            
            # 修正工令讀取邏輯
            final_job_no = res_main.get("header_info", {}).get("job_no")
            if not final_job_no or final_job_no == "Unknown":
                 final_job_no = res_main.get("job_no", "Unknown")
            
            st.session_state.analysis_result_cache = {
                "job_no": final_job_no,
                "header_info": res_main.get("header_info", {}),
                "all_issues": all_issues,
                "total_duration": time.time() - total_start,
                "ocr_duration": ocr_duration,
                "ai_duration": ai_duration,     # AI 耗時
                "py_duration": py_duration,     # Python 耗時
                
                "cost_twd": (usage.get("input", 0)*0.3 + usage.get("output", 0)*2.5) / 1000000 * 32.5,
                "total_in": usage.get("input", 0),
                "total_out": usage.get("output", 0),
                
                "ai_extracted_data": dim_data,
                "freight_target": res_main.get("freight_target", 0),
                "summary_rows": res_main.get("summary_rows", []),
                "full_text_for_search": combined_input,
                "combined_input": combined_input
            }
            
            progress_bar.progress(1.0)
            status_box.update(label="✅ 分析完成！", state="complete", expanded=False)
            st.rerun()

       # --- 💡 顯示結果區塊 ---
    if st.session_state.analysis_result_cache:
        cache = st.session_state.analysis_result_cache
        all_issues = cache.get('all_issues', [])

        # --- 📋 表頭資訊偵測 (手機版強製橫排優化) ---
        st.divider()
        st.subheader("📋 表頭資訊偵測")
        
        h_info = cache.get("header_info", {}) 
        current_job = h_info.get("job_no", "未偵測")
        sch_date = h_info.get("scheduled_date", "未偵測")
        act_date = h_info.get("actual_date", "未偵測")

        # 1. 先處理紅色警示的 HTML 樣式字串
        act_date_html = f"<b>{act_date}</b>"
        try:
            if act_date != "未偵測" and sch_date != "未偵測" and act_date > sch_date:
                # 如果逾期，變紅色 (#ff4b4b 是 Streamlit 的標準紅)
                act_date_html = f"<b style='color: #ff4b4b;'>{act_date} (逾期)</b>"
        except: pass

        # 2. 使用 HTML Flexbox 強制橫向排列
        st.markdown(f"""
        <div style="display: flex; flex-direction: row; justify-content: space-between; width: 100%;">
            <div style="flex: 1; padding-right: 5px;">
                <div style="font-size: 12px; color: gray; margin-bottom: 2px;">工令單號</div>
                <div style="font-size: 16px; font-weight: bold;">{current_job}</div>
            </div>
            <div style="flex: 1; padding-right: 5px;">
                <div style="font-size: 12px; color: gray; margin-bottom: 2px;">預定交貨日</div>
                <div style="font-size: 16px; font-weight: bold;">{sch_date}</div>
            </div>
            <div style="flex: 1;">
                <div style="font-size: 12px; color: gray; margin-bottom: 2px;">實際交貨日</div>
                <div style="font-size: 16px;">{act_date_html}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.divider()

        # 3. 頂部狀態條 (修改版：詳細時間拆解)
        # 格式：總耗時 (OCR | AI | Python)
        total_t = cache.get('total_duration', 0)
        ocr_t = cache.get('ocr_duration', 0)
        ai_t = cache.get('ai_duration', 0)
        py_t = cache.get('py_duration', 0)
        
        st.success(
            f"總耗時: {total_t:.1f}s  "
            f"( OCR: {ocr_t:.1f}s | AI: {ai_t:.1f}s | Py: {py_t:.2f}s )"
        )
        
        st.info(f"💰 本次成本: NT$ {cache['cost_twd']:.2f} (In: {cache['total_in']:,} / Out: {cache['total_out']:,})")
        
        # 4. 規則展示 (v58: 完整欄位六宮格版)
        with st.expander("🏗️ 檢視 Excel 邏輯與規則參數", expanded=False):
            
            # 1. 修正資料源：改讀 analysis_result_cache
            target_list = []
            if st.session_state.analysis_result_cache:
                target_list = st.session_state.analysis_result_cache.get('all_issues', [])
            
            # 2. 找出隱藏包裹 (HIDDEN_DATA)
            hidden_payload = {}
            for item in target_list:
                if item.get('issue_type') == 'HIDDEN_DATA':
                    hidden_payload = item
                    break
            
            # 3. 解析資料
            rule_hits = hidden_payload.get('rule_hits', {})
            current_fuzz = globals().get('GLOBAL_FUZZ_THRESHOLD', hidden_payload.get('fuzz_threshold', 90))

            st.caption(f"ℹ️ 全域統一特規門檻: **{current_fuzz} 分**")
            
            try:
                # 嘗試讀取 Excel 檔案
                df_rules = pd.read_excel("rules.xlsx")
                df_rules.columns = [c.strip() for c in df_rules.columns]
                
                # 建立快速查詢表
                rule_info_map = {}
                rules_map_for_xray = {} 
                
                for _, row in df_rules.iterrows():
                    r_name = str(row.get('Item_Name', '')).strip()
                    clean_k = r_name.replace(" ", "").replace("\n", "").replace("\r", "").replace('"', '').replace("'", "").strip()
                    rule_info_map[clean_k] = row
                    rules_map_for_xray[clean_k] = row

                # 4. 顯示結果 (如果有命中)
                if rule_hits:
                    st.success(f"🎯 系統偵測到 {len(rule_hits)} 種特規項目！")
                    
                    for rule_key, hits in rule_hits.items():
                        info = rule_info_map.get(rule_key, {})
                        
                        st.markdown(f"#### ✅ {rule_key}")
                        
                        # 🔥🔥🔥 [版面修改] 改為 2 欄排列，顯示 6 個欄位 🔥🔥🔥
                        c_left, c_right = st.columns(2)
                        
                        with c_left:
                            st.markdown(f"**Local:** `{info.get('Unit_Rule_Local', '-')}`")
                            st.markdown(f"**Freight:** `{info.get('Unit_Rule_Freight', '-')}`")
                            st.markdown(f"**Agg:** `{info.get('Unit_Rule_Agg', '-')}`")
                            
                        with c_right:
                            # 嘗試讀取更多欄位，若 Excel 沒這欄位會顯示 '-'
                            st.markdown(f"**Category:** `{info.get('Category', '-')}`")
                            st.markdown(f"**Process:** `{info.get('Process_Rule', '-')}`")
                            st.markdown(f"**Logic:** `{info.get('Logic_Prompt', '-')}`")
                        # -----------------------------------------------------
                        
                        # 顯示明細表格
                        hit_df = pd.DataFrame(hits)
                        cols_to_show = ["明細名稱", "分數", "匹配類型", "頁碼"]
                        final_cols = [c for c in cols_to_show if c in hit_df.columns]
                        
                        if "分數" in final_cols:
                            st.dataframe(hit_df[final_cols].style.format({"分數": "{:.0f}"}), use_container_width=True, hide_index=True)
                        else:
                            st.dataframe(hit_df, use_container_width=True, hide_index=True)
                        st.divider()
                else:
                    if target_list:
                        st.info(f"本次工令未觸發任何特規項目 (門檻: {current_fuzz})。")
                    else:
                        st.warning("⚠️ 尚未執行分析或無分析結果。")

                # 底部：完整的規則總表
                st.markdown("---")
                with st.expander("📋 查看完整規則總表 (All Rules)", expanded=False):
                    st.dataframe(df_rules, use_container_width=True, hide_index=True)

                # 🔥 X光機 (保留)
                st.markdown("---")
                st.subheader("🕵️‍♂️ X光檢測：為什麼沒抓到？")
                st.caption(f"這裡列出前 10 筆項目的最高分規則，幫您決定 GLOBAL_FUZZ_THRESHOLD 該設多少 (目前: {current_fuzz})")
                
                sample_items = []
                acc_input = st.session_state.get('analysis_result_cache', {}).get('ai_extracted_data', [])
                if acc_input:
                    sample_items = [item.get('item_title', '') for item in acc_input[:10]]
                
                if sample_items:
                    debug_data = []
                    for item_title in sample_items:
                        clean_title = item_title.replace(" ", "").replace("\n", "").strip()
                        best_score = 0
                        best_rule = "無"
                        
                        # 記得這裡要跟您最後決定使用的 fuzz 方式同步 (目前建議 token_sort_ratio)
                        for k in rules_map_for_xray.keys():
                            sc = fuzz.token_sort_ratio(k, clean_title)
                            if sc > best_score:
                                best_score = sc
                                best_rule = k
                        
                        status = "🔴 落榜"
                        if best_score > current_fuzz: status = "🟢 錄取"
                        
                        debug_data.append({
                            "工令項目": clean_title,
                            "最像的規則": best_rule,
                            "計算分數": best_score,
                            "狀態": status
                        })
                    st.dataframe(pd.DataFrame(debug_data))

            except Exception as e:
                st.error(f"UI 顯示錯誤: {e}")
                
        # 5. 原始數據檢視
        with st.expander("📊 檢視 AI 抄錄原始數據", expanded=False):
            st.markdown("**1. 核心指標摘要**")
            sum_rows_len = len(cache.get("summary_rows", []))
            summary_df = pd.DataFrame([{
                "工令單號": cache.get("job_no", "N/A"),
                "總表行數": sum_rows_len,
                "總表狀態": "正常" if sum_rows_len > 0 else "空值"
            }])
            st.dataframe(summary_df, hide_index=True, use_container_width=True)
            st.divider()
 
            st.markdown("**2. 左上角統計表 (Summary Rows)**")
            sum_rows = cache.get("summary_rows", [])
            
            if sum_rows:
                df_sum = pd.DataFrame(sum_rows)
                
                # 1. 確保頁碼欄位存在
                if "page" not in df_sum.columns: df_sum["page"] = "?"
                
                # 2. 欄位更名 (兼容舊版 target 與新版 delivery_qty)
                rename_map = {
                    "page": "頁碼", 
                    "title": "項目名稱", 
                    "apply_qty": "申請數量",    # ✅ 新增：申請數量
                    "delivery_qty": "實交數量", # ✅ 新增：實交數量
                    "target": "實交數量"        # 舊版兼容 (若無 delivery_qty 則用 target)
                }
                df_sum.rename(columns=rename_map, inplace=True)
                
                # 3. 指定顯示順序 (確保欄位不會消失)
                # 先列出我們想要的順序
                desired_cols = ["頁碼", "項目名稱", "申請數量", "實交數量"]
                # 只保留 DataFrame 中真的存在的欄位
                final_cols = [c for c in desired_cols if c in df_sum.columns]
                
                st.dataframe(df_sum[final_cols], hide_index=True, use_container_width=True)
            else:
                st.caption("無數據")

            st.divider()
            st.markdown("**3. 全卷詳細抄錄數據 (JSON)**")
            st.json(cache.get("ai_extracted_data", []), expanded=True)

        # ========================================================
        # ⚡️ [修正重點]：現在 all_issues 已經定義了，這裡就不會報錯了
        # ========================================================
        
        # 1. 執行合併
        consolidated_list = consolidate_issues(all_issues)

        # 2. 過濾出「真正的錯誤」
        real_errors_consolidated = [i for i in consolidated_list if "未匹配" not in i.get('issue_type', '')]

        # 3. 顯示結論
        if not all_issues:
            st.balloons()
            st.success("✅ 全數合格！")
        elif not real_errors_consolidated:
            st.success(f"✅ 數值合格！ (但有 {len(consolidated_list)} 類項目未匹配規則)")
        else:
            st.error(f"發現 {len(real_errors_consolidated)} 類異常")

        # 4. 卡片循環顯示 (v39: 數值精修版)
        for item in consolidated_list:
            #  [就在這裡！插入這兩行] 
            if item.get('issue_type') == 'HIDDEN_DATA':
                continue
                
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                source_label = item.get('source', '')
                issue_type = item.get('issue_type', '異常')
                
                # 頁碼處理
                page_str = item.get('page', '?')
                if "," in str(page_str):
                    page_display = f"Pages: {page_str}"
                else:
                    page_display = f"P.{page_str}"

                c1.markdown(f"**{page_display} | {item.get('item')}** `{source_label}`")
                
                # 燈號邏輯
                if any(kw in issue_type for kw in ["統計", "數量", "流程", "溯源", "總表", "匯總", "🚨", "🛑"]):
                    c2.error(f"{issue_type}")
                else:
                    c2.warning(f"{issue_type}")
                
                st.caption(f"原因: {item.get('common_reason', '')}")
                
                failures = item.get('failures', [])
                if failures:
                    # 1. 轉成 DataFrame
                    df = pd.DataFrame(failures)
                    
                    # 2. 欄位中文化
                    rename_map = {
                        "id": "編號",
                        "val": "實測",
                        "target": "目標",
                        "calc": "狀態",
                        "note": "備註"
                    }
                    df.rename(columns=rename_map, inplace=True)
                    
                    # 3. 樣式調整 (置中與靠左)
                    styler = df.style.set_properties(**{
                        'text-align': 'center', 
                        'white-space': 'nowrap'
                    })
                    
                    styler.set_table_styles([
                        dict(selector='th', props=[('text-align', 'center')])
                    ])

                    # 針對文字較長的欄位靠左
                    left_align_cols = [c for c in ["項目名稱", "編號", "Item"] if c in df.columns]
                    if left_align_cols:
                        styler.set_properties(subset=left_align_cols, **{'text-align': 'left'})

                    # 🔥 [新增] 4. 智能數值格式化 (Smart Formatting)
                    # 邏輯：整數顯示整數 (10)，小數顯示兩位 (10.53)
                    def smart_fmt(x):
                        try:
                            f = float(x)
                            # 如果跟四捨五入後的自己差很小，就當作整數
                            if abs(f - round(f)) < 0.000001: 
                                return f"{int(f)}"
                            return f"{f:.2f}"
                        except:
                            return str(x)

                    # 鎖定可能出現數字的欄位
                    target_cols = [c for c in ["實測", "目標", "數量"] if c in df.columns]
                    if target_cols:
                        styler.format(smart_fmt, subset=target_cols)

                    # 5. 顯示表格
                    st.dataframe(styler, use_container_width=True, hide_index=True)

            st.divider()
        
        # 下載按鈕邏輯
        current_job_no = cache.get('job_no', 'Unknown')
        safe_job_no = str(current_job_no).replace("/", "_").replace("\\", "_").strip()
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
