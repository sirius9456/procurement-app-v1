import streamlit as st
import pandas as pd
# 【修正點 1】新增 date 導入，解決 NameError: name 'date' is not defined
from datetime import datetime, timedelta, date 
from io import BytesIO
import os 
import json
import gspread
import logging
import time
import base64 # 新增 base64 導入，用於 PDF 預覽
# 【GCS 導入】新增 Google Cloud Storage 函式庫
from google.cloud import storage

# ******************************
# *--- 1. 全域設定與常數 ---*
# ******************************

# 配置 Streamlit 日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 版本號
APP_VERSION = "V2.2.10 (Attachment Deletion & Clickable)" 

# 時間格式
DATE_FORMAT = "%Y-%m-%d"
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# --- Google Sheets URL 設定 ---
# 已更新為您提供的網址
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
else:
    try:
        SHEET_URL = st.secrets["spreadsheet"]["url"]
    except:
        SHEET_URL = "https://docs.google.com/spreadsheets/d/16vSMLx-GYcIpV2cuyGIeZctvA2sI8zcqh9NKKyrs-uY/edit?usp=sharing"

# 工作表名稱 (測試版專用)
DATA_SHEET_NAME = '採購總表_測試'
METADATA_SHEET_NAME = '專案設定_測試'

# --- 憑證路徑設定 (智慧偵測) ---
# 優先順序：1. 環境變數 -> 2. secrets 資料夾 -> 3. 根目錄 -> 4. 預設
if "GSHEETS_CREDENTIALS_PATH" in os.environ:
    GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"]
elif os.path.exists("secrets/google_sheets_credentials.json"):
    GSHEETS_CREDENTIALS = "secrets/google_sheets_credentials.json"
elif os.path.exists("google_sheets_credentials.json"):
    GSHEETS_CREDENTIALS = "google_sheets_credentials.json"
else:
    GSHEETS_CREDENTIALS = "secrets/google_sheets_credentials.json" # 預設值

st.set_page_config(
    page_title=f"專案採購小幫手 {APP_VERSION}", 
    page_icon="🧪", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS 樣式
CUSTOM_CSS = """
<style>
    /* 強制指定中文字型 */
    html, body, [class*="css"] {
        font-family: "Microsoft JhengHei", "Noto Sans TC", "PingFang TC", sans-serif;
    }

    /* 儀表板樣式 */
    .metric-box {
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        color: white;
        margin-bottom: 10px;
    }
    .metric-title {
        font-size: 14px;
        opacity: 0.8;
    }
    .metric-value {
        font-size: 24px;
        font-weight: bold;
    }
    
    /* 專案標題樣式 */
    .project-header {
        font-size: 18px;
        font-weight: bold;
        color: #FF9800;
    }
    .item-header {
        font-size: 16px;
        font-weight: 600;
        color: #2196F3;
        margin-left: 10px;
    }
    .meta-info {
        font-size: 13px;
        color: #888;
    }
    
    /* 輸入欄位顏色統一 (適配深色模式) */
    div[data-baseweb="select"] > div, div[data-baseweb="base-input"] > input, div[data-baseweb="input"] > div { 
        background-color: #262730 !important; 
        color: white !important; 
        -webkit-text-fill-color: white !important; 
    }
    
    /* --- 日曆圖示修正 (強制白色) --- */
    /* 1. 針對 Streamlit 表格內的日期選擇器 */
    [data-testid="stDataFrame"] input[type="date"]::-webkit-calendar-picker-indicator {
        filter: invert(1) grayscale(100%) brightness(200%) !important;
        cursor: pointer;
    }
    
    /* 2. 針對一般的 date input (如側邊欄) */
    input[type="date"]::-webkit-calendar-picker-indicator {
        filter: invert(1) grayscale(100%) brightness(200%) !important;
        cursor: pointer;
    }
    
    /* 讓表格內連結看起來像連結 */
    .st-ag-row a {
        color: #2196F3 !important; /* 藍色連結 */
        text-decoration: underline !important;
        cursor: pointer !important;
    }
</style>
"""

STATUS_OPTIONS = ["詢價中", "已報價", "待採購", "已採購", "運送中", "已到貨", "已驗收", "取消"]

# *--- 1. 全域設定與常數 - 結束 ---*


# ******************************
# *--- 1. 登入與安全函式 ---*
# ******************************

def logout():
    """登出函式：清除驗證狀態並重新運行。"""
    st.session_state["authenticated"] = False
    st.rerun()

def login_form():
    """渲染登入表單並處理密碼驗證。"""
    
    DEFAULT_USERNAME = os.environ.get("AUTH_USERNAME", "dev_user")
    DEFAULT_PASSWORD = os.environ.get("AUTH_PASSWORD", "dev_pwd")
    
    credentials = {"username": DEFAULT_USERNAME, "password": DEFAULT_PASSWORD}

    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
        
    if st.session_state["authenticated"]:
        return 

    st.markdown("<br><br>", unsafe_allow_html=True)
    col_empty, col_center, col_empty2 = st.columns([1, 2, 1])
    
    with col_center:
        with st.container(border=True):
            st.title("🧪 測試版登入 (Test Env)")
            st.markdown("---")
            
            username = st.text_input("用戶名", key="login_username", value=credentials["username"], disabled=True)
            password = st.text_input("密碼", type="password", key="login_password")
            
            if st.button("登入", type="primary"):
                if username.strip() == credentials["username"].strip() and password == credentials["password"]:
                    st.session_state["authenticated"] = True
                    st.toast("✅ 測試版登入成功！")
                    st.rerun()
                else:
                    st.error("用戶名或密碼錯誤。")
            
    st.stop() 


# ******************************
# *--- 2. 數據讀取與寫入函式 (測試版) ---*
# ******************************

# 【設定】測試版專用的工作表名稱
DATA_SHEET_NAME = '採購總表_測試' 
METADATA_SHEET_NAME = '專案設定_測試'

# @st.cache_data(ttl=600, show_spinner="連線 Google Sheets...")
def load_data_from_sheets():
    """直接使用 gspread 讀取 Google Sheets 中的數據 (測試版)。"""
    
    # 【修改點 1】新增 '附件' 欄位
    expected_cols = ['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '最後修改時間', '附件', '標記刪除']
    
    if not SHEET_URL:
        st.info("❌ Google Sheets URL 尚未配置。使用空的數據結構。")
        empty_data = pd.DataFrame(columns=expected_cols)
        return empty_data, {}

    try:
        # --- 1. 授權與認證 ---
        # 檢查憑證是否存在 (使用全域變數 GSHEETS_CREDENTIALS，它已經經過智慧偵測)
        if not GSHEETS_CREDENTIALS or not os.path.exists(GSHEETS_CREDENTIALS):
             st.error(f"❌ 憑證錯誤：找不到憑證檔案。路徑: {GSHEETS_CREDENTIALS}")
             st.info("💡 提示：請確認 'google_sheets_credentials.json' 是否在根目錄、secrets 資料夾，或已設定環境變數。")
             raise FileNotFoundError("憑證檔案不存在或路徑錯誤")
            
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 2. 讀取採購總表 (Data) ---
        try:
            data_ws = sh.worksheet(DATA_SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"❌ 找不到工作表：**{DATA_SHEET_NAME}**")
            st.warning(f"請確認 Google Sheets 中是否存在名為「**{DATA_SHEET_NAME}**」的分頁。")
            return pd.DataFrame(columns=expected_cols), {}
            
        data_records = data_ws.get_all_records()
        data_df = pd.DataFrame(data_records)

        # 強制補齊欄位
        if data_df.empty:
            data_df = pd.DataFrame(columns=expected_cols)
        else:
            for col in expected_cols:
                if col not in data_df.columns:
                    if col in ['ID', '數量']:
                        data_df[col] = 0
                    elif col in ['單價', '總價']:
                        data_df[col] = 0.0
                    elif col in ['選取', '標記刪除']:
                         data_df[col] = False
                    else:
                        data_df[col] = '' # '附件' 預設為空字串

        # 【關鍵修正：布林值清洗】
        # 避免將空字串或異類格式誤判為 True，明確轉換
        def clean_bool(x):
            if isinstance(x, bool): return x
            # 只有字串明確為 "TRUE" (不分大小寫) 才算 True，其餘皆 False
            return str(x).strip().upper() == 'TRUE'

        for col in ['選取', '標記刪除']:
            if col in data_df.columns:
                data_df[col] = data_df[col].apply(clean_bool)

        # 數據類型轉換 (其他欄位)
        dtype_map = {
            'ID': 'Int64', '單價': 'float', '數量': 'Int64', '總價': 'float'
        }
        valid_dtype_map = {col: dtype for col, dtype in dtype_map.items() if col in data_df.columns}
        if valid_dtype_map:
            data_df = data_df.astype(valid_dtype_map, errors='ignore')
            
        # 確保附件欄位是字串
        if '附件' in data_df.columns:
            data_df['附件'] = data_df['附件'].astype(str)

        # 日期欄位處理
        if '預計交貨日' in data_df.columns:
            data_df['預計交貨日'] = pd.to_datetime(data_df['預計交貨日'], errors='coerce', format=DATE_FORMAT) 
        if '採購最慢到貨日' in data_df.columns:
            data_df['採購最慢到貨日'] = pd.to_datetime(data_df['採購最慢到貨日'], errors='coerce', format=DATE_FORMAT) 
        
        # --- 3. 讀取專案設定 (Metadata) ---
        try:
            metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"❌ 找不到工作表：**{METADATA_SHEET_NAME}**")
            st.warning(f"請確認 Google Sheets 中是否存在名為「**{METADATA_SHEET_NAME}**」的分頁。")
            return data_df, {}

        metadata_records = metadata_ws.get_all_records()
        
        project_metadata = {}
        if metadata_records:
            for row in metadata_records:
                try:
                    # 使用 from datetime import date 的 date
                    due_date = pd.to_datetime(str(row['專案交貨日'])).date()
                except (ValueError, TypeError):
                    due_date = datetime.now().date()
                    
                project_metadata[row['專案名稱']] = {
                    'due_date': due_date,
                    'buffer_days': int(row.get('緩衝天數', 7)),
                    'last_modified': str(row.get('最後修改', ''))
                }

        st.success(f"🧪 測試版數據已從 `{DATA_SHEET_NAME}` 及 `{METADATA_SHEET_NAME}` 載入！") 
        return data_df, project_metadata

    except Exception as e:
        logging.exception("Google Sheets 數據載入時發生致命錯誤！") 
        st.error(f"❌ 數據載入失敗！")
        st.code(f"錯誤訊息: {e}")
        empty_data = pd.DataFrame(columns=expected_cols)
        st.session_state.data_load_failed = True
        return empty_data, {}


def write_data_to_sheets(df_to_write, metadata_to_write):
    """直接使用 gspread 寫回 Google Sheets (測試版)。"""
    if st.session_state.get('data_load_failed', False) or not SHEET_URL:
        st.warning("數據載入失敗或 URL 未配置，已禁用寫入 Sheets。")
        return False
        
    try:
        # --- 1. 授權與認證 ---
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 2. 寫入採購總表 (Data) ---
        cols_to_drop = ['交期判定', '交期顯示']
        df_export = df_to_write.drop(columns=[c for c in cols_to_drop if c in df_to_write.columns], errors='ignore')

        # 日期轉字串
        for col in ['預計交貨日', '採購最慢到貨日']:
            if col in df_export.columns:
                df_export[col] = pd.to_datetime(df_export[col], errors='coerce').dt.strftime(DATE_FORMAT).fillna("")
                
        # 填充空值
        df_export = df_export.fillna("")
        
        # 【關鍵修正：布林值序列化】
        for col in ['選取', '標記刪除']:
            if col in df_export.columns:
                df_export[col] = df_export[col].apply(lambda x: bool(x))
        
        # 【修改點 2】確保附件欄位存在且為字串
        if '附件' not in df_export.columns:
            df_export['附件'] = ""
        else:
            df_export['附件'] = df_export['附件'].astype(str)

        # 轉為 object 以便相容
        df_export = df_export.astype(object) 
                
        try:
            data_ws = sh.worksheet(DATA_SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"❌ 找不到工作表：{DATA_SHEET_NAME}，無法寫入。")
            return False

        data_ws.clear()
        # 將 DataFrame 轉為列表列表 (List of Lists)
        data_to_update = [df_export.columns.values.tolist()] + df_export.values.tolist()
        data_ws.update(data_to_update)
        
        # --- 3. 寫入專案設定 (Metadata) ---
        metadata_list = [
            # 使用 from datetime import date 的 date
            {'專案名稱': name, 
             '專案交貨日': data['due_date'].strftime(DATE_FORMAT) if isinstance(data['due_date'], (datetime, date)) else str(data['due_date']),
             '緩衝天數': int(data['buffer_days']), 
             '最後修改': str(data['last_modified'])}
            for name, data in metadata_to_write.items()
        ]
        metadata_df = pd.DataFrame(metadata_list)
        
        try:
            metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"❌ 找不到工作表：{METADATA_SHEET_NAME}，無法寫入設定。")
            return False

        metadata_ws.clear()
        if not metadata_df.empty:
            metadata_ws.update([metadata_df.columns.values.tolist()] + metadata_df.values.tolist())
            
        st.cache_data.clear()
        return True
        
    except Exception as e:
        logging.exception("Google Sheets 數據寫入時發生致命錯誤！")
        st.error(f"❌ 數據寫回 Google Sheets 失敗！")
        st.code(f"錯誤訊息: {e}")
        return False
# *--- 2. 數據讀取與寫入函式 - 結束 ---*



# ******************************
# *--- 3. 輔助函式區 ---*
# ******************************
# ... (add_business_days, convert_df_to_excel, calculate_project_budget, calculate_dashboard_metrics, calculate_latest_arrival_dates 保持不變) ...

# 【GCS 輔助函式】

@st.cache_resource
def get_gcs_client():
    """初始化 GCS 客戶端 (使用 Streamlit 資源快取)。"""
    # 假設運行環境已配置 GCP 認證 (e.g., Service Account JSON, or environment variables)
    return storage.Client()

def upload_file_to_gcs(uploaded_file, quote_id):
    """將檔案上傳到 GCS Bucket，並返回物件名稱 (包含路徑)。"""
    if uploaded_file is None:
        return None
        
    client = get_gcs_client()
    bucket = client.bucket(GCS_BUCKET_NAME)
    
    # 構造 GCS 物件名稱：attachments/ID_原始檔名
    destination_blob_name = f"{GCS_FOLDER_PATH}/{quote_id}_{uploaded_file.name}"
    blob = bucket.blob(destination_blob_name)
    
    # 上傳檔案內容
    try:
        blob.upload_from_string(uploaded_file.getvalue(), content_type=uploaded_file.type)
        return destination_blob_name
    except Exception as e:
        logging.error(f"GCS 檔案上傳失敗: {e}")
        st.error(f"❌ 附件上傳到 GCS 失敗：{e}")
        return None

def delete_file_from_gcs(gcs_object_name):
    """從 GCS Bucket 中刪除檔案。"""
    if not gcs_object_name:
        return True # 如果檔案名是空的，視為成功
        
    client = get_gcs_client()
    bucket = client.bucket(GCS_BUCKET_NAME)
    blob = bucket.blob(gcs_object_name)
    
    try:
        # 檢查檔案是否存在再刪除
        if blob.exists():
            blob.delete()
            return True
        else:
            logging.warning(f"GCS 刪除警告：檔案 {gcs_object_name} 不存在。")
            return True
    except Exception as e:
        logging.error(f"GCS 檔案刪除失敗: {e}")
        return False

# *--- 3. 輔助函式區 - 結束 ---*



# ******************************
# *--- 9. 附件管理模組 (新功能) ---*
# ******************************
# 【修正點】將此區塊移到區塊 4 之前，確保主程式呼叫時函式已定義
import base64

def save_uploaded_file(uploaded_file, quote_id):
    """【GCS 實作】將上傳的檔案存到 Google Cloud Storage，並回傳 GCS 物件名稱。"""
    if uploaded_file is None:
        return None
        
    # 舊的本地檔案儲存邏輯已移除，直接呼叫 GCS 輔助函式
    gcs_object_name = upload_file_to_gcs(uploaded_file, quote_id)
    
    # 返回 GCS 物件名稱 (e.g., attachments/123_quote.pdf)
    return gcs_object_name 

def render_attachment_module(df):
    """
    渲染獨立的附件管理區塊。
    功能：選擇報價 -> 上傳/檢視附件 (支援圖片與 PDF 預覽)
    """
    st.markdown("---")
    st.subheader("📎 報價附件管理中心")
    
    # 1. 處理來自表格點擊的預覽請求
    auto_preview_id = st.session_state.get('preview_from_table_id', None)
    initial_proj = "請選擇..."
    initial_item_key = "請選擇..."
    
    if auto_preview_id is not None:
        try:
            row = df[df['ID'] == auto_preview_id].iloc[0]
            initial_proj = row['專案名稱']
            initial_item_key = f"{row['ID']} - {row['專案項目']} ({row['供應商']})"
            # 清除狀態，確保下次重新運行時不會自動選擇，除非再次點擊表格
            st.session_state.preview_from_table_id = None 
        except:
            pass
            
    # 2. 選擇器
    col_sel1, col_sel2 = st.columns([1, 2])
    
    selected_quote_id = None
    selected_quote_row = None
    
    # 篩選專案並預設選擇
    all_projects = df['專案名稱'].unique().tolist()
    initial_proj_list = ["請選擇..."] + all_projects
    initial_proj_index = initial_proj_list.index(initial_proj) if initial_proj in initial_proj_list else 0
    
    with col_sel1:
        selected_proj = st.selectbox("📂 選擇專案", initial_proj_list, index=initial_proj_index, key="att_proj_select")
        
    with col_sel2:
        if selected_proj != "請選擇...":
            # 篩選該專案下的報價項目
            proj_df = df[df['專案名稱'] == selected_proj]
            # 建立選單標籤: ID - 項目 - 供應商
            quote_options = {f"{row['ID']} - {row['專案項目']} ({row['供應商']})": row['ID'] for _, row in proj_df.iterrows()}
            
            # 篩選報價項目並預設選擇
            initial_item_list = ["請選擇..."] + list(quote_options.keys())
            initial_item_index = initial_item_list.index(initial_item_key) if initial_item_key in initial_item_list else 0
            
            selected_option = st.selectbox("📄 選擇報價項目", initial_item_list, index=initial_item_index, key="att_item_select")
            
            if selected_option != "請選擇...":
                selected_quote_id = quote_options[selected_option]
                # 取得該列資料
                selected_quote_row = df[df['ID'] == selected_quote_id].iloc[0]

    # 3. 附件操作區
    if selected_quote_id is not None and selected_quote_row is not None:
        
        col_upload, col_preview = st.columns([1, 1.5], gap="large")
        
        # 獲取 GCS 物件名稱
        gcs_object_name = str(selected_quote_row.get('附件', '')).strip()
        
        with col_upload:
            st.info(f"正在編輯 ID: **{selected_quote_id}** 的附件")
            
            # 顯示目前附件狀態
            if gcs_object_name:
                # 只顯示檔名部分
                display_filename = os.path.basename(gcs_object_name)
                st.success(f"✅ 目前 GCS 附件：`{display_filename}`")
                st.caption(f"GCS 路徑: {gcs_object_name}")
            else:
                st.warning("目前無附件")
                
            # 上傳元件
            uploaded_file = st.file_uploader("上傳新附件 (支援 JPG, PNG, PDF)", type=['png', 'jpg', 'jpeg', 'pdf'], key=f"uploader_{selected_quote_id}")
            
            if uploaded_file:
                if st.button("💾 確認上傳並儲存", type="primary"):
                    # 1. 執行上傳到 GCS
                    new_gcs_object_name = save_uploaded_file(uploaded_file, selected_quote_id)
                    
                    if new_gcs_object_name:
                        # 2. 更新 DataFrame (儲存 GCS 物件名稱)
                        idx = st.session_state.data[st.session_state.data['ID'] == selected_quote_id].index[0]
                        st.session_state.data.loc[idx, '附件'] = new_gcs_object_name
                        st.session_state.data.loc[idx, '最後修改時間'] = datetime.now().strftime(DATETIME_FORMAT)
                        
                        # 3. 寫入 Google Sheets
                        if 'write_data_to_sheets' in globals() and write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
                            st.toast(f"附件 {os.path.basename(new_gcs_object_name)} 上傳成功！")
                            time.sleep(1) 
                            st.rerun()
                        else:
                            st.error("❌ 寫入 Google Sheets 失敗，請檢查權限與連線。")
                    else:
                        st.error("❌ 檔案上傳 GCS 失敗。")


        with col_preview:
            st.markdown("#### 👁️ 附件預覽")
            if gcs_object_name:
                # 【GCS 預覽】使用 GCS 的公開存取 URL
                # 注意：這要求您的 Bucket 必須設置為公開讀取權限
                public_url = f"{GCS_BASE_URL}/{gcs_object_name}"
                display_filename = os.path.basename(gcs_object_name)
                
                # 判斷副檔名
                ext = os.path.splitext(display_filename)[1].lower()
                
                if ext in ['.png', '.jpg', '.jpeg']:
                    st.image(public_url, caption=display_filename, use_container_width=True)
                    
                elif ext == '.pdf':
                    # PDF 預覽，直接嵌入公開 URL
                    pdf_display = f'<iframe src="{public_url}" width="100%" height="600" type="application/pdf"></iframe>'
                    st.markdown(pdf_display, unsafe_allow_html=True)
                else:
                    st.info(f"此檔案格式 ({ext}) 不支援頁面內預覽 (僅支援圖片/PDF)。")
                    st.markdown(f"[點擊下載檔案: {display_filename}]({public_url})", unsafe_allow_html=True)
            else:
                st.caption("請選擇項目並上傳附件以進行預覽。")



# *--- 9. 附件管理模組 - 結束 ---*


# ******************************
# *--- 4. 邏輯處理函式 ---*
# ******************************


def handle_master_save():
    """批次處理所有 data_editor 的修改，並重新計算總價、更新個別報價時間戳記。"""
    
    if not st.session_state.edited_dataframes:
        st.info("沒有偵測到表格修改。")
        return

    main_df = st.session_state.data.copy()
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    
    changes_detected = False
    
    # 確保 DataFrame 有 '最後修改時間' 欄位，如果沒有則建立並用空字串填充
    if '最後修改時間' not in main_df.columns:
        main_df['最後修改時間'] = ''

    for _, edited_df in st.session_state.edited_dataframes.items():
        if edited_df.empty: continue
        
        for index, new_row in edited_df.iterrows():
            original_id = new_row['ID']
            idx_in_main = main_df[main_df['ID'] == original_id].index
            if idx_in_main.empty: continue
            
            main_idx = idx_in_main[0]
            
            row_changed = False

            # --- 數據比較與更新 ---
            
            # 處理 DateColumn 返回的 datetime 物件
            new_delivery_date = new_row['預計交貨日']
            if pd.notna(new_delivery_date):
                 new_delivery_date = pd.to_datetime(new_delivery_date).normalize() 
                 
                 # 比較 datetime 物件
                 if main_df.loc[main_idx, '預計交貨日'] != new_delivery_date:
                    main_df.loc[main_idx, '預計交貨日'] = new_delivery_date
                    row_changed = True

            # 檢查其他可更新欄位
            updatable_cols = ['選取', '供應商', '單價', '數量', '狀態', '標記刪除'] 
            for col in updatable_cols:
                 if str(main_df.loc[main_idx, col]) != str(new_row[col]):
                    main_df.loc[main_idx, col] = new_row[col]
                    row_changed = True
            
            # 重新計算總價 (總是執行以確保數據一致)
            current_price = float(main_df.loc[main_idx, '單價'])
            current_qty = float(main_df.loc[main_idx, '數量'])
            new_total = current_price * current_qty
            
            if main_df.loc[main_idx, '總價'] != new_total:
                main_df.loc[main_idx, '總價'] = new_total
                row_changed = True
            
            if row_changed:
                changes_detected = True
                # 【新功能】更新單個報價的最後修改時間
                main_df.loc[main_idx, '最後修改時間'] = current_time_str
                
    if changes_detected:
        st.session_state.data = main_df
        
        updated_metadata = st.session_state.project_metadata.copy()
        
        if write_data_to_sheets(st.session_state.data, updated_metadata):
            st.session_state.project_metadata = updated_metadata
            st.session_state.edited_dataframes = {}
            st.success("✅ 資料已儲存！總價、總預算及 Google Sheets 已更新。")
        
        st.rerun()
    else:
        st.info("沒有偵測到表格修改。")


def get_current_delete_ids():
    """輔助函式：從所有編輯器中匯總當前被標記刪除的 ID 列表。"""
    delete_map = {}
    
    # 遍歷所有編輯器暫存檔
    for edited_df in st.session_state.edited_dataframes.values():
        if edited_df is not None and not edited_df.empty:
            for _, row in edited_df.iterrows():
                delete_map[row['ID']] = row['標記刪除']
    
    ids_to_delete = []
    
    # 比對原始數據與編輯狀態
    for _, row in st.session_state.data.iterrows():
        item_id = row['ID']
        is_marked = delete_map.get(item_id, row['標記刪除'])
        
        if is_marked is True or str(is_marked).lower() == 'true':
            ids_to_delete.append(item_id)
            
    return ids_to_delete


def trigger_delete_confirmation():
    """
    第一步：鎖定目標。
    點擊 '刪除已標記項目' 按鈕時，立刻計算並鎖定要刪除的 ID，存入 Session State。
    """
    
    # 1. 獲取當前勾選的 ID
    ids_to_delete = get_current_delete_ids()
    
    if not ids_to_delete:
        st.warning("沒有項目被標記為刪除。請先在表格中勾選 '刪除?' 欄位。")
        st.session_state.show_delete_confirm = False
        if 'pending_delete_ids' in st.session_state:
            del st.session_state.pending_delete_ids
        return

    # 2. 將 ID 列表「鎖定」存入 session_state，供下一步使用
    st.session_state.pending_delete_ids = ids_to_delete
    st.session_state.delete_count = len(ids_to_delete)
    st.session_state.show_delete_confirm = True
    st.rerun()


def handle_batch_delete_quotes():
    """
    第二步：執行刪除並同步刪除附件檔案 (GCS)。
    """
    
    # 1. 從 Session State 讀取「鎖定」的 ID 列表
    ids_to_delete = st.session_state.get('pending_delete_ids', [])
    
    if not ids_to_delete:
        st.session_state.show_delete_confirm = False
        st.warning("刪除操作過期或未找到目標，請重新勾選並執行。")
        st.rerun()
        return

    # 2. 識別要刪除的項目及其附件
    main_df = st.session_state.data.copy() 
    deleted_quotes_df = main_df[main_df['ID'].isin(ids_to_delete)]
    
    # 3. 刪除附件檔案 (GCS)
    deleted_file_count = 0
    success = True
    for _, row in deleted_quotes_df.iterrows():
        gcs_object_name = str(row.get('附件', '')).strip()
        if gcs_object_name:
            if delete_file_from_gcs(gcs_object_name):
                deleted_file_count += 1
            else:
                success = False # 即使刪除失敗，也應繼續刪除資料庫記錄
                logging.error(f"附件刪除 GCS 失敗: {gcs_object_name}")
                
    # 4. 執行數據刪除：保留 ID 不在刪除列表中的項目
    df_after_delete = main_df[~main_df['ID'].isin(ids_to_delete)].reset_index(drop=True)
    
    # 5. 更新 Session State
    st.session_state.data = df_after_delete
    
    # 6. 寫入 Google Sheets
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.session_state.show_delete_confirm = False
        if success:
            st.success(f"✅ 已成功刪除 {len(ids_to_delete)} 筆報價。({deleted_file_count} 個附件檔案已清除) Sheets 已更新。")
        else:
            st.warning(f"已刪除 {len(ids_to_delete)} 筆報價，但有部分附件檔案從 GCS 刪除失敗。Sheets 已更新。")
        
        # 清除編輯暫存與鎖定的 ID
        st.session_state.edited_dataframes = {} 
        if 'pending_delete_ids' in st.session_state:
            del st.session_state.pending_delete_ids
    
    st.rerun()


def handle_project_modification():
    """處理修改專案設定的邏輯"""
    target_proj = st.session_state.edit_target_project
    new_name = st.session_state.edit_new_name
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    
    if not new_name:
        st.error("專案名稱不能為空")
        return
        
    if target_proj != new_name and new_name in st.session_state.project_metadata:
        st.error(f"新的專案名稱 '{new_name}' 已存在，請使用不同名稱。")
        return

    meta = st.session_state.project_metadata.pop(target_proj)
    st.session_state.project_metadata[new_name] = meta
    
    st.session_state.data.loc[st.session_state.data['專案名稱'] == target_proj, '專案名稱'] = new_name
    
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"✅ 專案已更新：{new_name}。Sheets 已更新。")
    
    st.rerun()


def handle_delete_project(project_to_delete):
    """刪除選定的專案及其所有相關報價 (GCS)。"""
    
    if not project_to_delete:
        st.error("請選擇要刪除的專案。")
        return

    # 刪除相關附件 (GCS)
    quotes_to_delete = st.session_state.data[st.session_state.data['專案名稱'] == project_to_delete]
    deleted_file_count = 0
    success = True
    for _, row in quotes_to_delete.iterrows():
        gcs_object_name = str(row.get('附件', '')).strip()
        if gcs_object_name:
            if delete_file_from_gcs(gcs_object_name):
                deleted_file_count += 1
            else:
                success = False
                logging.error(f"專案附件刪除 GCS 失敗: {gcs_object_name}")
    
    if project_to_delete in st.session_state.project_metadata:
        del st.session_state.project_metadata[project_to_delete]

    initial_count = len(st.session_state.data)
    st.session_state.data = st.session_state.data[
        st.session_state.data['專案名稱'] != project_to_delete
    ].reset_index(drop=True)
    
    deleted_count = initial_count - len(st.session_state.data)

    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        if success:
            st.success(f"✅ 專案 **{project_to_delete}** 及其相關的 {deleted_count} 筆報價已成功刪除。({deleted_file_count} 個附件檔案已清除) Sheets 已更新。")
        else:
            st.warning(f"已刪除專案 **{project_to_delete}**，但有部分附件檔案從 GCS 刪除失敗。Sheets 已更新。")
    
    st.rerun()


def handle_add_new_project():
    """處理新增專案設定的邏輯"""
    project_name = st.session_state.new_proj_name
    project_due_date = st.session_state.new_proj_due_date
    buffer_days = st.session_state.new_proj_buffer_days
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)

    if not project_name:
        st.error("專案名稱不能為空。")
        return
        
    # 如果專案已存在，則更新其時程
    if project_name in st.session_state.project_metadata:
        st.warning(f"專案 '{project_name}' 已存在，將更新其時程設定。")
    
    st.session_state.project_metadata[project_name] = {
        'due_date': project_due_date, 
        'buffer_days': buffer_days,
        'last_modified': current_time_str # 僅在新增/設定時更新此元數據
    }
    
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"✅ 已新增/更新專案設定：{project_name}。Sheets 已更新。")
    
    st.rerun()


def handle_add_new_quote(latest_arrival_date):
    """處理新增報價的邏輯"""
    project_name = st.session_state.quote_project_select
    item_name_to_use = st.session_state.item_name_to_use_final
    supplier = st.session_state.quote_supplier
    price = st.session_state.quote_price
    qty = st.session_state.quote_qty
    status = st.session_state.quote_status

    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    
    if st.session_state.quote_date_type == "1. 指定日期":
        final_delivery_date = st.session_state.quote_delivery_date
    else:
        final_delivery_date = st.session_state.calculated_delivery_date 

    if not project_name or not item_name_to_use:
        st.error("請確認已輸入專案名稱並選擇或輸入採購項目名稱。")
        return
    if project_name not in st.session_state.project_metadata:
        st.error(f"專案 '{project_name}' 的時程設定不存在。請先在 '➕ 新增專案' 區塊設定該專案的交期。")
        return
        
    total_price = price * qty
    
    new_row = {
        'ID': st.session_state.next_id, '選取': False, '專案名稱': project_name, 
        '專案項目': item_name_to_use, '供應商': supplier, '單價': price, '數量': qty, 
        '總價': total_price, 
        # DateColumn 需要 datetime 物件
        '預計交貨日': pd.to_datetime(final_delivery_date).normalize(), 
        '狀態': status, 
        # DateColumn 需要 datetime 物件
        '採購最慢到貨日': pd.to_datetime(latest_arrival_date).normalize(), 
        '標記刪除': False,
        # 【新功能】新增報價的最後修改時間
        '最後修改時間': current_time_str, 
        '附件': "" # 新增的附件欄位，預設為空字串
    }
    st.session_state.next_id += 1
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"✅ 已新增報價至 {project_name}！Sheets 已更新。")
    
    st.rerun()

# *--- 4. 邏輯處理函式 - 結束 ---*



# ******************************
# *--- 5. Session State 初始化函式 ---*
# ******************************
def initialize_session_state():
    """初始化所有 Streamlit Session State 變數。"""
    today = datetime.now().date()
    
    if 'data' not in st.session_state or 'project_metadata' not in st.session_state:
        data_df, metadata_dict = load_data_from_sheets()
        
        st.session_state.data = data_df
        st.session_state.project_metadata = metadata_dict
        
    if '標記刪除' not in st.session_state.data.columns:
        st.session_state.data['標記刪除'] = False
            
    if 'next_id' not in st.session_state:
        st.session_state.next_id = st.session_state.data['ID'].max() + 1 if not st.session_state.data.empty and pd.notna(st.session_state.data['ID'].max()) else 1
    
    if 'edited_dataframes' not in st.session_state: st.session_state.edited_dataframes = {}
    if 'calculated_delivery_date' not in st.session_state: st.session_state.calculated_delivery_date = today
    if 'show_delete_confirm' not in st.session_state: st.session_state.show_delete_confirm = False
    if 'delete_count' not in st.session_state: st.session_state.delete_count = 0
    # 新增 Session State 變數用於表格點擊預覽
    if 'preview_from_table_id' not in st.session_state: st.session_state.preview_from_table_id = None
# *--- 5. Session State 初始化函式 - 結束 ---*


# ******************************
# *--- 6. 模組化渲染函數 ---*
# ******************************

def render_sidebar_ui(df, project_metadata, today):
    """渲染整個側邊欄 UI：修改/刪除專案、新增專案、新增報價。"""
    
    with st.sidebar:
        
        # --- 區塊 1: 修改/刪除專案 ---
        with st.expander("✏️ 修改/刪除專案資訊", expanded=False): 
            all_projects = sorted(list(project_metadata.keys()))
            
            if all_projects:
                target_proj = st.selectbox("選擇目標專案", all_projects, key="edit_target_project")
                
                operation = st.selectbox(
                    "選擇操作項目", 
                    ("修改專案資訊", "刪除專案"), 
                    key="project_operation_select",
                    help="選擇 '刪除專案' 將永久移除專案及其所有報價。"
                )
                
                st.markdown("---")
                
                if operation == "修改專案資訊":
                    st.markdown("##### ✏️ 專案資訊修改")
                    st.text_input("新專案名稱", value=target_proj, key="edit_new_name")
                    
                    if st.button("確認修改專案名稱", type="primary", use_container_width=True): 
                        handle_project_modification()
                
                elif operation == "刪除專案":
                    st.markdown("##### 🗑️ 專案刪除 (⚠️ 警告)")
                    st.warning(f"您即將永久刪除專案 **{target_proj}** 及其所有相關報價資料。")
                    
                    if st.button(f"確認永久刪除 {target_proj}", type="secondary", help="此操作不可逆，將同時移除所有相關報價", use_container_width=True):
                        handle_delete_project(target_proj)
                        
            else: 
                st.info("無專案可修改/刪除。請在下方新增專案。")
        
        
        # --- 區塊 2: 新增/設定專案時程 ---
        with st.expander("➕ 新增/設定專案時程", expanded=False): 
            st.info("💡 若輸入現有專案名稱，將更新該專案的交貨日與緩衝天數。")
            
            st.text_input("專案名稱 (Project Name)", key="new_proj_name")
            
            project_due_date = st.date_input("專案交貨日 (Project Due Date)", value=today + timedelta(days=30), key="new_proj_due_date")
            buffer_days = st.number_input("採購緩衝天數 (天)", min_value=0, value=7, key="new_proj_buffer_days")
            
            latest_arrival_date_proj = project_due_date - timedelta(days=int(buffer_days))
            st.caption(f"計算得出最慢到貨日：{latest_arrival_date_proj.strftime('%Y年%m月%d日')}")

            if st.button("💾 儲存專案設定", key="btn_save_proj", use_container_width=True):
                handle_add_new_project()
        
        
        # --- 區塊 3: 新增報價 ---
        with st.expander("➕ 新增報價", expanded=False): 
            all_projects_for_quote = sorted(list(project_metadata.keys()))
            latest_arrival_date = today 
            
            if not all_projects_for_quote:
                st.warning("請先在上方新增/設定專案時程。")
                project_name = None
            else:
                project_name = st.selectbox("選擇目標專案", all_projects_for_quote, key="quote_project_select")
                
                current_meta = project_metadata.get(project_name, {'due_date': today, 'buffer_days': 7})
                buffer_days = current_meta['buffer_days']
                latest_arrival_date = current_meta['due_date'] - timedelta(days=int(buffer_days))

                st.caption(f"專案最慢到貨日: {latest_arrival_date.strftime(DATE_FORMAT)}")

            st.markdown("##### 採購項目選擇")
            
            unique_items = sorted(df['專案項目'].unique().tolist())
            item_options = ['🆕 新增項目...'] + unique_items 

            selected_item = st.selectbox("選擇現有項目", item_options, key="quote_item_select")

            item_name_to_use = None
            if selected_item == '🆕 新增項目...':
                item_name_to_use = st.text_input("輸入新的採購項目名稱", key="quote_item_new_input")
            else:
                item_name_to_use = selected_item
            
            st.session_state.item_name_to_use_final = item_name_to_use
            
            st.text_input("供應商名稱", key="quote_supplier")
            
            # 修正: 單價改為整數輸入 (min_value=0, step=1)
            st.number_input("單價 (TWD)", min_value=0, step=1, key="quote_price") 
            
            st.number_input("數量", min_value=1, value=1, key="quote_qty")
            
            st.markdown("##### 預計交貨日輸入")
            date_input_type = st.radio("選擇輸入方式", ("1. 指定日期", "2. 自然日數", "3. 工作日數"), key="quote_date_type", horizontal=True)

            if date_input_type == "1. 指定日期": 
                final_delivery_date = st.date_input("選擇確切交貨日期", today, key="quote_delivery_date") 
            
            elif date_input_type == "2. 自然日數": 
                num_days = st.number_input("自然日數", min_value=1, value=7, key="quote_num_days_input")
                final_delivery_date = today + timedelta(days=int(num_days))
                st.session_state.calculated_delivery_date = final_delivery_date 
                
            elif date_input_type == "3. 工作日數": 
                num_b_days = st.number_input("工作日數", min_value=1, value=5, key="quote_num_days_input")
                final_delivery_date = add_business_days(today, int(num_b_days))
                st.session_state.calculated_delivery_date = final_delivery_date
            
            if date_input_type != "1. 指定日期":
                final_delivery_date = st.session_state.calculated_delivery_date
                st.caption(f"計算得出的交期：{final_delivery_date.strftime(DATE_FORMAT)}")

            st.selectbox("目前狀態", STATUS_OPTIONS, key="quote_status")
            
            if st.button("📥 新增資料", key="btn_add_quote", type="primary", use_container_width=True):
                handle_add_new_quote(latest_arrival_date)


        # 恢復 V2.1.6 原始登出按鈕位置
        st.button("🚪 登出系統", on_click=logout, type="secondary", key="sidebar_logout_btn")


def render_dashboard(df, project_metadata):
    """渲染頂部儀表板區塊。"""
    
    # *--- render_dashboard - 儀表板區塊 ---*
    total_projects, total_budget, risk_items, pending_quotes = calculate_dashboard_metrics(df, project_metadata)

    st.subheader("📊 總覽儀表板")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class='metric-box' style='background-color:#33343c;'>
            <div class='metric-title'>專案總數</div>
            <div class='metric-value'>{total_projects}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='metric-box' style='background-color:#1b4d3e;'>
            <div class='metric-title'>預估/已選總預算</div>
            <div class='metric-value'>${total_budget:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div class='metric-box' style='background-color:#5a2a2a;'>
            <div class='metric-title'>交期風險項目</div>
            <div class='metric-value'>{risk_items}</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div class='metric-box' style='background-color:#2a3b5a;'>
            <div class='metric-title'>待處理報價數量</div>
            <div class='metric-value'>{pending_quotes}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    # *--- render_dashboard - 儀表板區塊 - 結束 ---*


def render_batch_operations():
    """渲染儲存/刪除按鈕及確認對話框。"""
    
    col_save, col_delete = st.columns([0.8, 0.2])
    
    is_locked = st.session_state.show_delete_confirm
    
    with col_save:
        if st.button("💾 儲存表格修改並計算總價/預算", type="primary", disabled=is_locked):
            handle_master_save()
            
    with col_delete:
        if st.button("🔴 刪除已標記項目", type="secondary", disabled=is_locked, key="btn_trigger_delete"):
            trigger_delete_confirmation()

    # 模擬確認對話框
    if st.session_state.show_delete_confirm:
        st.error(f"⚠️ 確認永久刪除 **{st.session_state.delete_count}** 筆已標記的報價嗎？此操作不可逆！")
        
        col_confirm_yes, col_confirm_no, _ = st.columns([0.2, 0.2, 0.6])
        
        with col_confirm_yes:
            if st.button("✅ 確認刪除", key="confirm_delete_yes", type="primary"):
                handle_batch_delete_quotes()
        
        with col_confirm_no:
            if st.button("❌ 取消", key="confirm_delete_no"):
                st.session_state.show_delete_confirm = False
                st.rerun()

    st.markdown("---")
    
    
def render_project_tables(df, project_metadata):
    """渲染主介面中所有專案的 Data Editor 表格。"""
    
    if df.empty:
        st.info("目前沒有採購報價資料。")
        return
        
    project_groups = df.groupby('專案名稱')
    project_names = list(project_groups.groups.keys())
    
    is_locked = st.session_state.show_delete_confirm

    # 【新增功能：處理點擊事件】
    # 檢查是否有來自表格的點擊，如果有，更新 Session State
    query_params = st.experimental_get_query_params()
    if 'preview_id' in query_params:
        try:
            clicked_id = int(query_params['preview_id'][0])
            st.session_state.preview_from_table_id = clicked_id
        except:
            pass
        # 清除 URL 參數，避免重整時重複觸發
        st.experimental_set_query_params(preview_id=None)


    for i, proj_name in enumerate(project_names):
        proj_data = project_groups.get_group(proj_name)
        meta = project_metadata.get(proj_name, {})
        proj_budget = calculate_project_budget(df, proj_name)
        
        # --- 計算最慢到貨日 (專案交期 - 緩衝天數) ---
        due_date_val = meta.get('due_date')
        if isinstance(due_date_val, str):
            try:
                due_date_val = datetime.strptime(due_date_val, "%Y-%m-%d").date()
            except:
                due_date_val = datetime.now().date()
        
        buffer_days_val = int(meta.get('buffer_days', 7))
        latest_arrival_proj = due_date_val - timedelta(days=buffer_days_val)
        latest_arrival_str = latest_arrival_proj.strftime(DATE_FORMAT)

        # 標題列
        header_html = f"""
        <span class='project-header'>💼 專案: {proj_name}</span> &nbsp;|&nbsp; 
        <span class='project-header'>總預算: ${proj_budget:,.0f}</span> &nbsp;|&nbsp; 
        <span class='meta-info'>交期: {meta.get('due_date')}</span> &nbsp;|&nbsp;
        <span class='meta-info' style='color:#a8a8a8;'>⚠️ 最慢到貨: {latest_arrival_str}</span>
        """
        
        # 建立 Expander key
        expander_key = f"expander_{proj_name}"

        # 監聽 Expander 點擊事件
        with st.expander(label=f"專案：{proj_name} (點擊展開)", expanded=False): 
            st.markdown(header_html, unsafe_allow_html=True)
            
            for item_name, item_data in proj_data.groupby('專案項目'):
                
                has_selection = item_data['選取'].any()
                sub_total = item_data[item_data['選取']]['總價'].sum() if has_selection else item_data['總價'].min()
                calc_method = "(已選)" if has_selection else "(預估)"
                
                st.markdown(f"""
                <span class='item-header'>📦 {item_name}</span> 
                <span class='meta-info'> | 計入: ${sub_total:,.0f} {calc_method}</span>
                """, unsafe_allow_html=True)

                editable_df = item_data.copy()
                
                # 【關鍵修正】逐行清洗資料，確保只有 Python date 物件或 None
                if '預計交貨日' in editable_df.columns:
                    temp_series = pd.to_datetime(editable_df['預計交貨日'], errors='coerce')
                    editable_df['預計交貨日'] = temp_series.apply(lambda x: x.date() if pd.notnull(x) else None)
                
                if '採購最慢到貨日' in editable_df.columns:
                    temp_limit = pd.to_datetime(editable_df['採購最慢到貨日'], errors='coerce')
                    editable_df['採購最慢到貨日'] = temp_limit.apply(lambda x: x.date() if pd.notnull(x) else None)
                
                if '最後修改時間' not in editable_df.columns:
                    editable_df['最後修改時間'] = ''

                editor_key = f"editor_{proj_name}_{item_name}"
                
                # 【新增功能：附件連結】在 DataFrame 中創建顯示用的連結欄位
                def create_link_markdown(row):
                    file_name = row.get('附件', '').strip()
                    quote_id = row['ID']
                    if file_name:
                        # 創建一個連結到當前頁面，但帶有 query parameter 的連結
                        # 點擊後會觸發 run_app 頂部的邏輯，設置 session state 進行預覽
                        return f"[📎 {file_name}](?preview_id={quote_id})" 
                    return ""
                
                editable_df['附件_display'] = editable_df.apply(create_link_markdown, axis=1)
                
                # 【修正點 3】表格欄位顯示順序：將 '附件_display' 放在 '最後修改時間' 之後
                cols_to_display = ['選取', '供應商', '單價', '數量', '總價', '預計交貨日', '交期判定', '狀態', '最後修改時間', '附件_display', '標記刪除'] 

                # 使用 column_order 來控制顯示
                edited_df_value = st.data_editor(
                    editable_df,
                    column_order=cols_to_display,
                    column_config={
                        "選取": st.column_config.CheckboxColumn("選取", width="tiny"), 
                        "供應商": st.column_config.Column("供應商", disabled=False), 
                        "單價": st.column_config.NumberColumn("單價", format="$%d"),
                        "數量": st.column_config.NumberColumn("數量"),
                        "總價": st.column_config.NumberColumn("總價", format="$%d", disabled=True),
                        
                        "預計交貨日": st.column_config.DateColumn(
                            "預計交貨日", 
                            min_value=datetime(2020, 1, 1).date(),
                            max_value=datetime(2030, 12, 31).date(),
                            format="YYYY-MM-DD", 
                            step=1,
                            help="點擊兩下以開啟月曆選單"
                        ),
                        
                        "交期判定": st.column_config.Column("判定", width="tiny", help="❌: 延誤 / ✅: 準時", disabled=True),
                        "狀態": st.column_config.SelectboxColumn("狀態", options=STATUS_OPTIONS),
                        
                        "最後修改時間": st.column_config.TextColumn(
                            "最後修改時間",
                            disabled=True,
                            width="medium",
                            help="報價項目最後儲存的時間"
                        ),
                        
                        # 【修正點 4】附件欄位配置為唯讀連結顯示
                        "附件_display": st.column_config.TextColumn("附件", disabled=True, width="medium", help="點擊檔名可跳轉至下方預覽"),
                        
                        "標記刪除": st.column_config.CheckboxColumn("刪除?", width="tiny"), 
                    },
                    key=editor_key,
                    hide_index=True, 
                    use_container_width=True,
                    height=150 + (len(item_data) * 35) if len(item_data) > 3 else 150,
                    disabled=is_locked
                )
                
                st.session_state.edited_dataframes[item_name] = edited_df_value 
                st.markdown("---")
    
    # *** 資料匯出區塊 ***
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("💾 資料匯出")
    st.download_button("📥 下載 Excel 報表", 
                      convert_df_to_excel(df), 
                      f'procurement_report_{datetime.now().strftime("%Y%m%d")}.xlsx', 
                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# *--- 6. 模組化渲染函數 - 結束 ---*


# ******************************
# *--- 7. 主應用程式核心邏輯 ---*
# ******************************

def run_app():
    """運行應用程式的核心邏輯，在成功登入後調用。"""
    
    if 'expander_states' not in st.session_state:
        st.session_state.expander_states = {}

    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}") 
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    
    # 【CSS 修正】強制將日曆圖示反轉為亮色 (invert 100%)，解決深色模式下看不清的問題
    st.markdown("""
        <style>
        [data-testid="stDataFrame"] input[type="date"]::-webkit-calendar-picker-indicator {
            filter: invert(1);
            cursor: pointer;
        }
        </style>
    """, unsafe_allow_html=True)

    initialize_session_state()

    # 數據自動計算
    st.session_state.data = calculate_latest_arrival_dates(
        st.session_state.data, 
        st.session_state.project_metadata
    )
    
    if st.session_state.get('data_load_failed', False):
        st.warning("應用程式無法從 Google Sheets 載入數據，請檢查上方錯誤訊息。")
        
    # --- UI 核心邏輯開始 ---
    
    # 【判定邏輯更新】
    def get_date_judgment_icon(row):
        try:
            d_val = pd.to_datetime(row['預計交貨日'])
            l_val = pd.to_datetime(row['採購最慢到貨日'])
            
            if pd.isna(d_val) or pd.isna(l_val):
                return ""
                
            # 若 預計交貨日 > 最慢到貨日 -> 延遲 (❌)
            if d_val.date() > l_val.date():
                return "❌" 
            else:
                return "✅" 
        except:
            return ""

    if not st.session_state.data.empty:
        # 建立 '交期判定'
        st.session_state.data['交期判定'] = st.session_state.data.apply(get_date_judgment_icon, axis=1)
        
        # 確保 '最後修改時間' 欄位存在
        if '最後修改時間' not in st.session_state.data.columns:
            st.session_state.data['最後修改時間'] = ''

    df = st.session_state.data
    project_metadata = st.session_state.project_metadata
    today = datetime.now().date()
    
    # 渲染所有區塊
    render_sidebar_ui(df, project_metadata, today)
    render_dashboard(df, project_metadata)
    render_batch_operations()
    render_project_tables(df, project_metadata) 
    
    # 【新增】呼叫附件管理模組 
    render_attachment_module(df)

# ******************************
# *--- 8. 程式入口點 ---*
# ******************************

def main():
    
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True) 
        
    login_form()
    
    if st.session_state.authenticated:
        run_app() 
        
if __name__ == "__main__":
    main()


