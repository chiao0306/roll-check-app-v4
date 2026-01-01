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
    # 讀取 Excel 規則
    dynamic_rules = get_dynamic_rules(full_text_for_search)

    system_prompt = f"""
    你是一位極度嚴謹的中鋼機械品管【總稽核官】。你必須像「電腦程式」一樣執行以下雙模組稽核，禁止任何主觀解釋。
    
    {dynamic_rules}

    ---

    #### ⚔️ 模組 A：工程尺寸數據提取 (AI 翻譯官任務)
    
    1. **規格提取警報與完整性 (重要)**：
       - **多重目標處理**：若規格內包含多個「目標數值」（如：驅動端 157mm / 非驅動端 127mm），請將所有目標數字全部填入 `threshold_list`，並取最大值填入 `threshold`。這 **不是** 提取失敗。
       - **報警機制**：若規格文字中有數字，但你完全無法解析（提取後 threshold 為 0 或 null），你「必須」在 `issues` 清單回報 `🛑規格提取失敗`。只要有抓到任何一個目標數字，嚴禁報錯。
       
    2. **目標規格解析 (mm 定位與雜訊過濾)**：
       - **✅ 必抓 (目標尺寸)**：優先尋找與「至...」、「以上」、「±」、「~」、「直徑」直接關聯的數字。
       - **❌ 排除 (加工量雜訊)**：嚴禁提取「每次車修...」、「進刀量...」、「加工量...」後面的小數字（如 0.5~5mm）。
       - **📏 本體未再生底線**：針對「本體」未再生項目，其目標門檻（threshold）**絕對不會小於 120mm**。請自動忽略標題或規格中任何小於 120 的數字（如 #1機、項次2、車修3mm）。
       - **區間計算**：若有 `±` 或偏差，必須先算出最終範圍。如 `300±0.1` -> `[[299.9, 300.1]]`。
    
    3. **項目分類決策流程 (由上至下執行，命中即停止)**：
       - **LEVEL 1 (最高優先)：銲補判定**
         * 標題含「銲補」、「銲接」 -> 分類必為 `min_limit`。
         * (註：即便標題含軸頸或未再生，只要有銲補，以此為準)。
         
       - **LEVEL 2：未再生判定**
         * 標題含「未再生」時，進行二選一：
           a. 含「軸頸」 -> 分類必為 `max_limit`。 (💡 提示：即便有驅動/非驅動多個數字，也請全部放入 threshold_list，不准變更為 range)。
           b. 不含「軸頸」(本體) -> 分類必為 `un_regen`。
         * (⚠️ 警告：嚴禁因規格文字含「再生」而將其歸類為 range)。
         
       - **LEVEL 3：精加工與裝配判定**
         * 標題「不含未再生」，且包含「再生」、「研磨」、「精加工」、「車修加工」、「組裝」、「拆裝」、「真圓度」、「KEYWAY」-> 分類必為 `range`。
         * (💡 提示：這類項目要求兩位小數，規格多以「區間」(如 129~135) 或「± 公差」呈現，請務必精確算出 std_ranges)。
    
    4. **數據抄錄 (字串保護模式)**：
       - **禁止簡化**：實測值若顯示 `349.90`，必須輸出 `"349.90"`。禁止寫成 `349.9`。
       - **格式**：所有實測值必須包裹成雙引號字串。`["RollID", "實測值字串"]`。
    #### 🚫 數據抄錄純淨化指令 (核心禁令)：
       - **字體辨識優先**：利用視覺能力區分「原始打印字體」與「手寫筆跡」。
       - **絕對忽略手寫**：嚴禁抄錄任何手寫的數字、箭頭符號 (->)、刪除線、圓圈標記、勾選符號、簽名或日期。
       - **唯一數據來源**：僅提取儲存格內「原始打印」的數值。若儲存格因手寫標註產生混亂字串（如 "129.93 -> 129.94"），你必須無視手寫部分，僅輸出打印的字串 `"129.93"`。
       - **禁止描述雜訊**：不要在 JSON 內容中嘗試解釋或描述手寫的更正內容。
    
    5. **尺寸大小邏輯檢查**：
       - **物理位階準則**：`未再生車修 < 研磨 < 再生車修 < 銲補`。
       - **判定要求**：判定要求：針對同一 Roll ID，跨製程之尺寸大小必須符合上述位階邏輯。注意：同一編號出現在不同項目表格中代表「全流程紀錄」，屬於正常現象，嚴禁判定為衝突。 僅在位階不符（例如：研磨尺寸大於再生車修）時，才回報 🛑流程異常。
        
    #### 💰 模組 B：會計與流程數據提取 (AI 任務：抄錄傳票)
    **【重要禁令】：嚴禁在此判斷「數量是否正確」，計算工作由系統後台執行。**
    **提取左上角【統計表】**：必須抄錄統計表每一行（包含熱處理、拆裝、車修等）。
    **抄錄傳票 (核心要求)**：
       - **嚴禁遺漏**：頁面中若有「多個獨立標題」的表格（例如先拆裝 170、再拆裝 200），你必須將它們視為「不同的項目」分別抄錄到 `dimension_data`。
       - **禁止合併**：即便項目名稱相似，只要位置不同，就必須分成多個物件回傳。

    1. **提取左上角【統計表】(Summary Table)**：
       - 請將統計表（左上角）中每一行包含「實交數量」的項目提取出來。
       - **格式**：`summary_rows: [ {{ "title": "項目名稱", "target": 數字 }}, ... ]`
       - **提取運費**：單獨提取左上角運費項次的數字到 `freight_target`。

    2. **內文項目屬性抄錄**：
       - **item_pc_target**: 提取項目括號內的數字（如 12PC 提取 12）。
       - **accounting_rules**: 必須精確抄錄 Excel 知識庫中的 `Unit_Rule_Local` (單項)、`Unit_Rule_Agg` (聚合)、`Unit_Rule_Freight` (運費) 文字。
       - **特別要求**：若 `Unit_Rule_Agg` 包含多個資訊（如「豁免, 2SET=1PC」），必須「原封不動」全部抄錄並以逗號隔開。禁止自行刪減文字。

    3. **工件流程與尺寸位階檢查 (由 AI 判定並報於 issues)**：
       - **位階**：`未再生 < 研磨 < 再生 < 銲補`。若後段尺寸小於前段（銲補除外），報 `🛑流程異常`。
       - **溯源與重複性**：出現「研磨/再生」必須往前檢查是否有前段紀錄。
       - **特別注意**：同一編號在不同項目中多次出現是「全製程紀錄」，完全合法，**不准回報「同時存在」或「物理流程衝突」**。
    
    4. **⚖️ 流程稽核純淨化指令：
       - **無視手寫意見**：在判斷「物理位階」與「工件溯源」時，僅依據表格內的打印數據。
       - **忽略標記雜訊**：嚴禁因為數據旁邊有手寫的「OK」、「合格」或「箭頭」而影響判定。
       - **鎖定打印事實**：即使手寫更正後的數字看起來更合理，你也必須「以原始打印數值」作為判定物理邏輯的唯一依據。

    ---

    ### 📝 輸出規範 (Output Format)
    必須回傳單一 JSON。`issues` 僅存放：流程異常、規格提取失敗、表頭不一。數量與統計異常由系統自動產出，不准填入 `issues`。

    {{
      "job_no": "工令編號",
      "summary_rows": [
         {{ "title": "統計表項目名稱", "target": "實交數量數字" }} 
      ], // 💡 必須抄錄左上角統計表的「每一行」數據
      "freight_target": 0, // 💡 左上角運費項次的數字
      "issues": [ 
         {{
           "page": "頁碼", "item": "項目", "issue_type": "統計不符 / 🛑流程異常 / 🛑規格提取失敗",
           "common_reason": "原因",
           "failures": [
              {{ "id": "🔍 統計總帳基準", "val": "數", "calc": "目標" }},
              {{ "id": "項目 (P.頁碼)", "val": "數", "calc": "計入" }},
              {{ "id": "🧮 內文實際加總", "val": "數", "calc": "計算" }}
           ]
         }}
      ],
      "dimension_data": [
         {{
           "page": "數字",
           "item_title": "名稱",
           "category": "分類",
           "item_pc_target": 0, // 項目括號內的 PC 數
           "accounting_rules": {{ "local": "", "agg": "", "freight": "" }}, // 💡 從Excel精確抄錄
           "standard_logic": {{
              "logic_type": "必須從 [range, un_regen, min_limit, max_limit] 選一填入", 
              "threshold_list": [], // 規格中出現的所有數字
              "ranges_list": [],    // AI 預算好的 [[min, max]]
              "threshold": 0        // 主要的門檻數字，嚴禁填 0 (若標題有數字)
           }},
           "std_spec": "含 mm 的原始規格文字",
           "data": [ ["RollID", "實測值字串"] ] // 💡 務必保留末尾的 0，如 "349.90"
         }}
      ]
    }}

    #### 💡 AI 翻譯官範例 (禁止抄襲數字，須抓取當前標題真實數字)：
    1. range: 如 `XXX±YYY` -> {{ "logic_type": "range", "min": XXX-YYY, "max": XXX+YYY }}
    2. un_regen: 如 `至 XXXmm 再生` -> {{ "logic_type": "un_regen", "threshold": XXX }}
    3. min_limit: 如 `XXXmm 以上` -> {{ "logic_type": "min_limit", "min": XXX }}
    4. max_limit: 如 `XXXmm 以下` -> {{ "logic_type": "max_limit", "max": XXX }}
    """
    
    generation_config = {"response_mime_type": "application/json", "temperature": 0.0, "top_k": 1, "top_p": 0.95}
    
    try:
        if "gemini" in model_name.lower():
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(model_name)
            response = model.generate_content([system_prompt, combined_input], generation_config=generation_config)
            raw_content = response.text
            usage_meta = response.usage_metadata
            usage_in = usage_meta.prompt_token_count if usage_meta else 0
            usage_out = usage_meta.candidates_token_count if usage_meta else 0
        else:
            client = OpenAI(api_key=OPENAI_KEY)
            response = client.chat.completions.create(
                model=model_name,
                messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": combined_input}],
                temperature=0.0
            )
            raw_content = response.choices[0].message.content
            usage_in = response.usage.prompt_tokens
            usage_out = response.usage.completion_tokens

        # JSON 清洗
        if "```json" in raw_content:
            raw_content = raw_content.replace("```json", "").replace("```", "")
        elif "```" in raw_content:
            raw_content = raw_content.replace("```", "")
            
        try:
            parsed_data = json.loads(raw_content)
        except:
            parsed_data = {"job_no": "JSON Error", "issues": []}

        final_response = parsed_data if isinstance(parsed_data, dict) else {"job_no": "Unknown", "issues": []}
        if "issues" not in final_response: final_response["issues"] = []
        if "job_no" not in final_response: final_response["job_no"] = "Unknown"

        valid_issues = []
        for i in final_response["issues"]:
            if isinstance(i, dict) and i.get("item"):
                reason = i.get("common_reason", "")
                i_type = i.get("issue_type", "")
                if "合格" in reason and "未匹配" not in i_type: continue
                if "合格" in reason and "未匹配" in i_type: i["issue_type"] = "⚠️未匹配規則"
                valid_issues.append(i)
        
        final_response["issues"] = valid_issues
        final_response["_token_usage"] = {"input": usage_in, "output": usage_out}
        
        return final_response

    except Exception as e:
        return {"job_no": "Error", "issues": [{"item": "System Error", "common_reason": str(e)}], "_token_usage": {"input": 0, "output": 0}}
        
# --- 重點：Python 引擎獨立於 agent 函式之外 ---
def python_numerical_audit(dimension_data):
    grouped_errors = {} # 改用字典來進行分類收集
    import re
    if not dimension_data: return [] # 修正：若無資料回傳空清單

    for item in dimension_data:
        raw_data_list = item.get("data", [])
        title = item.get("item_title", "")
        cat = str(item.get("category", "")).strip()
        page_num = item.get("page", "?")
        raw_spec = str(item.get("std_spec", ""))
        
        # --- 🛡️ 數據清洗與「模式優先」預解析 (保留您的完整邏輯) ---
        trusted_stds = [] 
        logic = item.get("standard_logic", {})
        s_ranges = logic.get("ranges_list", []) if logic.get("ranges_list") else item.get("std_ranges", [])
        
        # 1. 抓取緊貼 "mm" 的數字
        mm_nums = [float(n) for n in re.findall(r"(\d+\.?\d*)\s*mm", raw_spec)]
        trusted_stds.extend(mm_nums)

        # 2. 解析 ± 或偏差結構
        pm_match = re.findall(r"(\d+\.?\d*)\s*[±]\s*(\d+\.?\d*)", raw_spec)
        for base, offset in pm_match:
            b, o = float(base), float(offset)
            s_ranges.append([b - o, b + o])
            trusted_stds.extend([b, b-o, b+o])

        # 3. 執行雜訊過濾
        all_nums = [float(n) for n in re.findall(r"(\d+\.?\d*)", raw_spec)]
        noise = [350.0, 300.0, 200.0, 145.0, 130.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0]
        clean_std = [n for n in all_nums if (n in trusted_stds) or (n not in noise and n > 5)]

        # 獲取 AI 傳來的邏輯參數
        l_type = logic.get("logic_type")
        s_list = logic.get("threshold_list", [])
        s_threshold = logic.get("threshold")

        for entry in raw_data_list:
            if not isinstance(entry, list) or len(entry) < 2: continue
            
            # 1. 先抓取 AI 抄錄下來的原始字串（可能包含手寫雜訊如 "129.93 -> 129.94"）
            rid, val_raw = str(entry[0]).strip(), str(entry[1]).strip()
            if not val_raw or val_raw in ["N/A", "nan", "M10"]: continue

            try:
                # 💡 [核心修改點]：只抓取字串中的第一個數字，無視後面的塗改
                # 使用 re.findall 找出所有符合數字格式的內容，取索引 [0] 的那一個
                val_match = re.findall(r"\d+\.?\d*", val_raw)
                val_str = val_match[0] if val_match else val_raw 

                # 2. 接下來的判定都使用這個乾淨的 val_str
                val = float(val_str)
                # 💡 精確檢查：必須含小數點且後綴長度為 2 (仍會檢查 349.90 的結尾 0)
                is_two_dec = "." in val_str and len(val_str.split(".")[-1]) == 2
                is_pure_int = "." not in val_str
                is_passed, reason, t_used, engine_label = True, "", "N/A", "未知"
                
                # ... (下方後續邏輯完全不用動) ...
                # --- 💡 [核心修正]：重新排列判定優先序，解決關鍵字碰撞 ---

                # 1. 【銲補模式】優先權最高
                if l_type == "min_limit" or "銲補" in (cat + title):
                    engine_label = "銲補(下限)"
                    if not is_pure_int:
                        is_passed, reason = False, "銲補格式錯誤: 應為純整數"
                    elif clean_std:
                        t_used = min(clean_std, key=lambda x: abs(x - val))
                        if val < t_used: is_passed, reason = False, f"銲補不足: 實測 {val} < 基準 {t_used}"

                # 2. 【未再生模式】(包含本體與軸頸) 優先於精加工
                elif l_type in ["un_regen", "max_limit"] or "未再生" in (cat + title):
                    # 分支 A: 軸頸未再生 (max_limit)
                    if "軸頸" in (cat + title):
                        engine_label = "軸頸(上限)"
                        # 1. 收集所有可能的數字標準
                        candidates = [float(n) for n in (clean_std + s_list)]
                        if s_threshold: candidates.append(float(s_threshold))
                        
                        # 2. 🛡️ 安全鎖：如果完全沒抓到基準數字，直接跳過判定，不准用 0 判斷
                        if not candidates or max(candidates) == 0:
                            continue 

                        target = max(candidates)
                        t_used = target
                        
                        # 3. 執行判定邏輯
                        if not is_pure_int: 
                            is_passed, reason = False, "軸頸格式錯誤: 應為純整數"
                        elif val > target: 
                            is_passed, reason = False, f"超過上限 {target}"
                    
                    # 分支 B: 本體未再生 (un_regen)
                    else:
                        engine_label = "未再生(本體)"
                        candidates = [float(n) for n in (clean_std + s_list) if float(n) >= 120.0]
                        if s_threshold and float(s_threshold) >= 120.0: candidates.append(float(s_threshold))
                        
                        if candidates:
                            if not candidates: continue
                            target = max(candidates)
                            t_used = target
                            if val <= target:
                                if not is_pure_int: is_passed, reason = False, f"未再生(<=標準{target}): 應為整數"
                            else:
                                if not is_two_dec: is_passed, reason = False, f"未再生(>標準{target}): 應填兩位小數(含末尾0)"
                        else:
                            is_passed = True # 沒抓到120以上標準則不判定

                # 3. 【精加工/再生/車修/組裝模式】最後判定
                elif l_type == "range" or any(x in (cat + title) for x in ["再生", "精加工", "研磨", "車修", "組裝", "拆裝", "真圓度"]):
                    engine_label = "精加工(區間)"
                    if not is_two_dec:
                        is_passed, reason = False, "精加工格式錯誤: 應填兩位小數(如.90)"
                    elif s_ranges:
                        t_used = str(s_ranges)
                        is_passed = any(r[0] <= val <= r[1] for r in s_ranges if len(r)==2)
                        if not is_passed: reason = f"尺寸不在區間 {t_used} 內"
                    elif clean_std:
                        s_min, s_max = min(clean_std), max(clean_std)
                        t_used = f"{s_min}~{s_max}"
                        if not (s_min <= val <= s_max): is_passed, reason = False, f"不在範圍內 {t_used}"

                # 💡 [合併卡片與模式顯示]
                if not is_passed:
                    # 使用 engine_label 讓畫面顯示更清楚
                    error_key = (page_num, title, reason)
                    if error_key not in grouped_errors:
                        grouped_errors[error_key] = {
                            "page": page_num,
                            "item": title,
                            "issue_type": f"數值異常({engine_label})",
                            "rule_used": f"Excel: {raw_spec}",
                            "common_reason": reason,
                            "failures": [],
                            "source": "🐍 系統判定"
                        }
                    grouped_errors[error_key]["failures"].append({
                        "id": rid, 
                        "val": val_str, 
                        "target": f"基準:{t_used}", 
                        "calc": f"⚖️ {engine_label} 引擎"
                    })
            except: continue
            
    return list(grouped_errors.values())
    
def python_accounting_audit(dimension_data, res_main):
    
    #Python 會計官：
    #1. 全項目單項核對 (本體去重/軸頸計行)
    #2. 軸頸編號重複性監控 (限2次)
    #3. 總表對帳 (A聚合/B一般雙模式)
    #4. 運費動態精算 (支援 XPC=1 換算)
    #5. 支援 Agg Rule 混合指令 (豁免籃子, 單位換算)
    
    accounting_issues = []
    from thefuzz import fuzz
    from collections import Counter
    import re
    
    # --- 1. 取得對帳基準 (來自左上角統計表) ---
    summary_rows = res_main.get("summary_rows", [])
    # 💡 關鍵修正：建立總表追蹤器，並執行「字串轉數字」安全過濾
    global_sum_tracker = {}
    for s in summary_rows:
        s_title = s.get('title', 'Unknown')
        s_target_raw = s.get('target', 0)
        try:
            # 處理可能含逗號的字串如 "4,524"
            s_target = float(str(s_target_raw).replace(',', '').strip())
        except:
            s_target = 0
        global_sum_tracker[s_title] = {"target": s_target, "actual": 0, "details": []}

    # 💡 取得運費基準數字
    freight_target_raw = res_main.get("freight_target", 0)
    try:
        freight_target = float(str(freight_target_raw).replace(',', '').strip())
    except:
        freight_target = 0

    # --- 2. 開始逐項遍歷內文數據 ---
    for item in dimension_data:
        title = item.get("item_title", "")
        page = item.get("page", "?")
        rules = item.get("accounting_rules", {})
        data_list = item.get("data", []) # 格式: [["ID", "Val"], ...]
        
        # 取得所有 ID 的清單 (清洗)
        ids = [str(e[0]).strip() for e in data_list if e and len(e) > 0]
        id_counts = Counter(ids)

        # 💡 [2.1 單項 PC 數核對] 
        try:
            target_pc = float(str(item.get("item_pc_target", 0)))
        except:
            target_pc = 0
            
        u_local = str(rules.get("local", "")) if rules.get("local") else ""
        is_body = "本體" in title
        is_journal = any(k in title for k in ["軸頸", "內孔", "Journal"])
        
        # 計算實際數量：1SET=4PCS, 1SET=2PCS, 本體去重, 其餘計行數
        if "1SET=4PCS" in u_local: 
            actual_item_qty = len(data_list) / 4
        elif "1SET=2PCS" in u_local: 
            actual_item_qty = len(data_list) / 2
        elif is_body or "PC=PC" in u_local: 
            actual_item_qty = len(set(ids)) # 去重
        else: 
            actual_item_qty = len(data_list) # 計行

        if actual_item_qty != target_pc and target_pc > 0:
            accounting_issues.append({
                "page": page, "item": title, "issue_type": "統計不符(單項)",
                "common_reason": f"標題要求 {target_pc}PC，內文核算為 {actual_item_qty}",
                "failures": [
                    {"id": f"項目標題目標", "val": target_pc, "calc": "目標"},
                    {"id": "內文實際計數", "val": actual_item_qty, "calc": "實際"}
                ],
                "source": "🐍 會計引擎"
            })

        # 💡 [2.2 軸頸三支禁令]
        if is_journal:
            for rid, count in id_counts.items():
                if count >= 3:
                    accounting_issues.append({
                        "page": page, "item": title, "issue_type": "🛑編號重複異常",
                        "common_reason": f"編號 {rid} 出現 {count} 次，違反軸頸限2次規定",
                        "failures": [{"id": rid, "val": f"{count} 次", "calc": "禁止超過2次"}],
                        "source": "🐍 會計引擎"
                    })

        # 💡 [2.3 總表與運費對帳]
        # 解析 Agg 規則 (支援 豁免, 2SET=1PC 混合格式)
        u_agg_raw = str(rules.get("agg", "")).strip()
        agg_parts = [p.strip() for p in u_agg_raw.split(",")]
        is_exempt_from_baskets = "豁免" in agg_parts
        
        agg_multiplier = 1.0
        for p in agg_parts:
            conv_match = re.search(r"(\d+)SET=1PC", p)
            if conv_match: agg_multiplier = 1.0 / float(conv_match.group(1))

        for s_title, data in global_sum_tracker.items():
            u_freight = str(rules.get("freight", "")) if rules.get("freight") else ""
            is_freight_row = "運費" in s_title
            
            match = False
            current_add_val = actual_item_qty # 預設

            if is_freight_row:
                # 🚚 運費模式
                if "豁免" in u_freight: continue
                elif "計入" in u_freight: match = True
                elif is_body and "未再生" in title: match = True
                
                if match:
                    # 動態換算：支援 2PC=1, 3PC=1...
                    conv = re.search(r"(\d+)PC=1", u_freight)
                    if conv: current_add_val = actual_item_qty / int(conv.group(1))
            else:
                # 📦 總表核對 (A/B雙模式)
                is_repair = any(k in s_title for k in ["ROLL車修", "再生"])
                is_weld   = "銲補" in s_title
                is_assem  = any(k in s_title for k in ["拆裝", "組裝", "裝配"])
                is_basket_row = is_repair or is_weld or is_assem

                if is_basket_row:
                    # A模式 (聚合籃子)：受「豁免」標籤影響
                    if is_exempt_from_baskets:
                        match = False
                    else:
                        if is_repair and any(k in title for k in ["未再生", "再生", "研磨", "車修"]): match = True
                        elif is_weld and "銲補" in title: match = True
                        elif is_assem and any(k in title for k in ["拆裝", "組裝", "真圓度"]): match = True
                
                # B模式 (一般核對)：名字對上就點貨，不受「豁免」影響
                if not match and fuzz.partial_ratio(s_title, title) > 85:
                    match = True

                if match:
                    current_add_val = actual_item_qty * agg_multiplier

            if match:
                data["actual"] += current_add_val
                label = "計入運費" if is_freight_row else "計入總帳"
                data["details"].append({"id": f"{title} (P.{page})", "val": current_add_val, "calc": label})

    # --- 3. 結算異常報告 ---
    for s_title, data in global_sum_tracker.items():
        if abs(data["actual"] - data["target"]) > 0.01 and data["target"] > 0:
            icon = "🚚" if "運費" in s_title else "🔍"
            accounting_issues.append({
                "page": "總表", "item": s_title, "issue_type": "統計不符",
                "common_reason": f"標註 {data['target']} != 內文加總 {data['actual']}",
                "failures": [{"id": f"{icon} 統計基準", "val": data["target"], "calc": "目標"}] + data["details"] + [{"id": "🧮 實際總計", "val": data["actual"], "calc": "計算"}],
                "source": "🐍 會計引擎"
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
        
        # 1. 執行 AI 分析
        t0 = time.time()
        # 💡 [修正]：不再重複傳送 full_text_for_search
        # 既然 full_text_for_search 只是用來找規則，那就不要把它當成參數傳給 agent
        res_main = agent_unified_check(combined_input, combined_input, GEMINI_KEY, main_model_name)
        time_main = time.time() - t0
        
        progress_bar.progress(100)
        status.empty()
        
        total_end = time.time()
        
        # --- 1. 成本計算 (完全依照您的版本，原封不動) ---
        usage_main = res_main.get("_token_usage", {"input": 0, "output": 0})
        
        def get_model_rate(model_name):
            name = model_name.lower()
            if "gpt" in name:
                if "mini" in name: return 0.15, 0.60
                elif "3.5" in name: return 0.50, 1.50
                else: return 2.50, 10.00
            else:
                if "flash" in name: return 0.5, 3.00
                else: return 1.25, 10.00 # Pro

        rate_in, rate_out = get_model_rate(main_model_name)
        
        cost_usd = (usage_main["input"] / 1_000_000 * rate_in) + (usage_main["output"] / 1_000_000 * rate_out)
        cost_twd = cost_usd * 32.5
        
        # --- 2. 啟動 Python 硬核數值稽核 (改在這裡執行一次即可) ---
        dim_data = res_main.get("dimension_data", [])
        python_numeric_issues = python_numerical_audit(dim_data)
        
        # --- 💡 [新增插入] 啟動 Python 會計引擎 (解決 NameError) ---
        # 這裡會執行您最看重的聚合模式、本體去重與運費核對
        python_accounting_issues = python_accounting_audit(dim_data, res_main)
        
        # --- 3. Python 表頭檢查 ---
        python_header_issues, python_debug_data = python_header_check(st.session_state.photo_gallery)
        
        # --- 4. 合併結果 (正式移交權限) ---
        ai_raw_issues = res_main.get("issues", [])
        ai_filtered_issues = []

        for i in ai_raw_issues:
            i['source'] = '🤖 總稽核 AI'
            i_type = i.get("issue_type", "")
            
            # 只有流程異常、規格提取失敗、表頭、未匹配聽 AI 的
            # 統計與數量不符現在交給 Python 引擎了，所以排除 AI 原本報的
            ai_only_tasks = ["流程", "規格提取失敗", "表頭", "未匹配"]
            
            if any(k in i_type for k in ai_only_tasks):
                ai_filtered_issues.append(i)
        
        # 最終合併所有稽核籃子
        all_issues = ai_filtered_issues + python_numeric_issues + python_accounting_issues + python_header_issues
        
        st.session_state.analysis_result_cache = {
            "job_no": res_main.get("job_no", "Unknown"),
            "all_issues": all_issues,
            "total_duration": total_end - total_start,
            "cost_twd": cost_twd,
            "total_in": usage_main["input"],
            "total_out": usage_main["output"],
            "ocr_duration": ocr_duration,
            "time_eng": time_main, # 這裡借用變數名，實為總時間
            "time_acc": 0,         # 單一代理無第二時間
            "full_text_for_search": full_text_for_search,
            "combined_input": combined_input,
            "python_debug_data": python_debug_data,
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
