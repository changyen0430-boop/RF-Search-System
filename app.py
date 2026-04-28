import streamlit as st
import requests
import io
import pandas as pd
import os
import re
import fitz  # PyMuPDF
from io import BytesIO
from PIL import Image

# --- 頁面設定 ---
st.set_page_config(page_title="RF Spec Search Pro", layout="wide", page_icon="📡")

# --- [Recall Memory 核心引擎] ---
class RecallMemoryEngine:
    def __init__(self):
        self.vendor_knowledge = {
            "AMPAK": {
                "chapter_anchor": r"(?i)Wi-Fi\s+RF\s+Specification", 
                "content_keywords": ["WIFI", "RF", "SPECIFICATION"],
                "ignore_keywords": ["REVISION", "CONTENTS"],
                "default_fallback": 8
            },
            "DEFAULT": {
                "chapter_anchor": r"(?i)(RF|Wireless)\s+Specification",
                "content_keywords": ["RF", "SPECIFICATION"],
                "ignore_keywords": ["REVISION"],
                "default_fallback": 0
            }
        }

    def get_strategy(self, vendor_name):
        v_upper = str(vendor_name).upper()
        for key in self.vendor_knowledge:
            if key in v_upper: return self.vendor_knowledge[key]
        return self.vendor_knowledge["DEFAULT"]

    def smart_navigate(self, doc, vendor_name):
        strategy = self.get_strategy(vendor_name)
        if not doc: return strategy["default_fallback"]
        for i in range(1, min(8, len(doc))):
            page = doc[i]
            links = page.get_links()
            for link in links:
                if link["kind"] == fitz.LINK_GOTO:
                    link_rect = link["from"]
                    link_text = page.get_text("text", clip=link_rect).strip()
                    if re.search(strategy["chapter_anchor"], link_text):
                        target_idx = link["page"]
                        target_text = doc[target_idx].get_text().upper()
                        if any(k in target_text for k in strategy["content_keywords"]):
                            return target_idx
        return strategy["default_fallback"]

memory_engine = RecallMemoryEngine()

# --- [PDF 處理函數] ---
def get_pdf_doc(url):
    try:
        if not url or url == "NA": return None
        if "drive.google.com" in url:
            file_id = ""
            if "/view" in url: match = re.search(r'/d/([^/]+)', url)
            else: match = re.search(r'id=([^&]+)', url)
            if match:
                file_id = match.group(1)
                download_url = f'https://drive.google.com/uc?export=download&id={file_id}'
                response = requests.get(download_url, timeout=15)
                return fitz.open(stream=response.content, filetype="pdf")
        elif os.path.exists(url):
            return fitz.open(url)
    except Exception as e:
        st.sidebar.error(f"解析 PDF 失敗: {e}")
    return None

def get_pdf_image_from_url(url, page_num):
    doc = get_pdf_doc(url)
    if doc:
        page_idx = max(0, min(page_num, len(doc)-1))
        pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(2, 2))
        return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return None

# --- 1. 資料加載邏輯 [中午的正確解法：強制從 fx 公式摳出網址] ---
@st.cache_data(ttl=3600) # 既然解析成功了，快取開起來搜尋才會秒開
def load_data():
    file_path = os.path.join(os.getcwd(), 'Final_Summary_Result_Pro.xlsx')
    if os.path.exists(file_path):
        try:
            # 1. 先用正常的 pandas 讀取基礎資料 (不讀取公式，這很快)
            df = pd.read_excel(file_path)
            
            # 2. 針對「Link to Datasheet」這一欄，我們用 openpyxl 重新「強硬」掃描一次原始公式
            from openpyxl import load_workbook
            wb = load_workbook(file_path, data_only=False) # 確保讀取公式而非結果
            ws = wb.active # 假設資料在第一個分頁
            
            # 找到 Link 欄位的索引 (假設在最後一欄)
            headers = [cell.value for cell in ws[1]]
            if "Link to Datasheet" in headers:
                col_idx = headers.index("Link to Datasheet") + 1
                formulas = []
                # 從第 2 列開始讀取每一格的「原始內容」
                for row in range(2, ws.max_row + 1):
                    cell_val = ws.cell(row=row, column=col_idx).value
                    # 轉為字串並摳出網址
                    v = str(cell_val) if cell_val else ""
                    if "HYPERLINK" in v.upper():
                        match = re.search(r'HYPERLINK\("([^"]+)"', v, re.I)
                        formulas.append(match.group(1) if match else "NA")
                    else:
                        # 如果不是公式，看它本身是不是網址
                        formulas.append(v if v.startswith("http") else "NA")
                
                # 將強行讀取的公式覆蓋回 DataFrame
                df["Link to Datasheet"] = formulas[:len(df)]
            
            return df.fillna("NA")
        except Exception as e:
            st.error(f"Excel 強制讀取失敗: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

# 改完後建議重整網頁並 Clear Cache
df = load_data()



# --- 2. 側邊欄搜尋與過濾 (恢復功能) ---
st.sidebar.header("🔍 智能過濾面板")
keyword = st.sidebar.text_input("全域搜尋 (空格隔開)", placeholder="例如: AMPAK 6E",
    key="search_input") # 給它一個固定的 Key，幫助 Streamlit 追蹤狀態)

# 補回 Wi-Fi 與 BT 選單
all_wifi = sorted([x for x in df["Feature Support_Wi-Fi"].unique() if x != "NA"]) if not df.empty else []
selected_wifi = st.sidebar.multiselect("Wi-Fi 規格過濾", all_wifi)
all_bt = sorted([x for x in df["Feature Support_BT"].unique() if x != "NA"]) if not df.empty else []
selected_bt = st.sidebar.multiselect("BT 規格過濾", all_bt)

# 執行過濾邏輯
filtered_df = df.copy()
if keyword:
    for term in keyword.split():
        filtered_df = filtered_df[filtered_df.apply(lambda r: r.astype(str).str.contains(term, case=False).any(), axis=1)]
if selected_wifi:
    filtered_df = filtered_df[filtered_df["Feature Support_Wi-Fi"].isin(selected_wifi)]
if selected_bt:
    filtered_df = filtered_df[filtered_df["Feature Support_BT"].isin(selected_bt)]
    
    
# --- [在這之間插入：3. 下載過濾後的 Excel 功能] ---
# --- 3. 下載過濾後的 Excel 功能 (美化版) ---
st.sidebar.divider()
if not filtered_df.empty:
    buffer = io.BytesIO()
    
    # 使用 xlsxwriter 進行美化寫入
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Search_Results')
        
        workbook  = writer.book
        worksheet = writer.sheets['Search_Results']

        # --- 定義美化格式 ---
        # 1. 標題格式：深藍底、白字、粗體、框線、置中
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#4472C4',  # 經典專業藍
            'font_color': 'white',
            'border': 1
        })

        # 2. 內容格式：框線、水平/垂直置中
        cell_format = workbook.add_format({
            'valign': 'vcenter',
            'align': 'center',
            'border': 1
        })

        # --- 應用格式 ---
        # 寫入標題並套用格式
        for col_num, value in enumerate(filtered_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            # 自動調整欄寬 (抓標題長度或內容長度，給個基本寬度)
            worksheet.set_column(col_num, col_num, 18, cell_format)

    st.sidebar.download_button(
        label="📥 下載美化版 Excel",
        data=buffer.getvalue(),
        file_name=f"RF_Spec_Export_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.sidebar.info("💡 尚無符合條件的資料可供下載")
# ----------------------------------------------    

# --- 4. 主要顯示區域 ---
st.title("📡 RF Spec Search System by RF Tommy")
st.metric("符合結果", len(filtered_df))

event = st.dataframe(
    filtered_df, use_container_width=True, on_select="rerun", selection_mode="single-row",
    column_config={
        "Link to Datasheet": st.column_config.LinkColumn("Datasheet 連結", display_text="Open PDF"),
        "Dimension(L W H)": st.column_config.TextColumn("Dimension (L W H)", width="medium"), # <-- 同步修改這裡
        "Vendor PN": st.column_config.TextColumn("Vendor PN", width="medium")
    }
)

# --- 5. 即時顯示與 AI 日誌 ---
st.divider()
if event and "selection" in event and len(event["selection"]["rows"]) > 0:
    selected_row_idx = event["selection"]["rows"][0]
    row_data = filtered_df.iloc[selected_row_idx]
    pdf_link = str(row_data["Link to Datasheet"])
    pn_name = str(row_data.get("Vendor PN", ""))
    vendor_name = str(row_data.get("Vendor", ""))

    st.subheader(f"📖 RF 特性即時預覽: {pn_name} ({vendor_name})")

    # 顯示預覽圖 (移除 http 限制，支援本地路徑)
    if pdf_link != "NA":
        if 'last_pn' not in st.session_state or st.session_state.last_pn != pn_name:
            doc = get_pdf_doc(pdf_link)
            st.session_state.current_page = memory_engine.smart_navigate(doc, vendor_name)
            st.session_state.last_pn = pn_name

        btn_col1, btn_col2, btn_col3 = st.columns([1, 2, 1])
        with btn_col1:
            if st.button("⬅️ PREV"): st.session_state.current_page = max(0, st.session_state.current_page - 1); st.rerun()
        with btn_col2:
            st.markdown(f"<h3 style='text-align: center;'>第 {st.session_state.current_page + 1} 頁</h3>", unsafe_allow_html=True)
        with btn_col3:
            if st.button("NEXT ➡️"): st.session_state.current_page += 1; st.rerun()

        img = get_pdf_image_from_url(pdf_link, st.session_state.current_page)
        if img: st.image(img, use_container_width=True)
        else: st.warning("⚠️ 無法載入 PDF 預覽圖片，請檢查檔案權限或路徑。")
    else:
        st.info("⚠️ 該型號目前無 PDF 連結。")

    # [關鍵補回] AI 決策邏輯分析區 (獨立顯示，不受 PDF 載入影響)
    st.markdown("---")
    with st.expander("🧠 AI 決策邏輯分析 (Logic Analysis)", expanded=True):
        log_file = 'ai_debug_report.txt'
        if os.path.exists(log_file):
            try:
                with open(log_file, 'r', encoding='utf-8') as f: all_logs = f.read()
                # 使用 pn_name 匹配 log 內容
                search_pattern = rf"=== Logic Analysis for:.*{pn_name}.*===(.*?)(?==== Logic Analysis for:|$)"
                match = re.search(search_pattern, all_logs, re.DOTALL | re.IGNORECASE)
                if match: st.code(match.group(0).strip(), language="text")
                else: st.info(f"💡 找不到 {pn_name} 的詳細分析日誌。")
            except: st.error("讀取日誌失敗。")
        else: st.warning("⚠️ 找不到 ai_debug_report.txt。")

else:
    st.info("💡 請點擊上方表格列來啟動即時預覽。")

st.caption("RecallMemory v4.6 | 搜尋與過濾完整版")