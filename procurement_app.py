import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os 
import json
import gspread
import logging
import time

# 引入 Google Cloud Storage 庫
from google.cloud import storage

# 確保 openpyxl 庫已安裝 (pip install openpyxl)

# --- 應用程式設定與常數 ---
APP_VERSION = "v2.2.5 (Production + Hyperlink Fix)"
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"]
DATE_FORMAT = "%Y-%m-%d"
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# --- Google Cloud Storage 配置 ---
# 請替換為您的儲存桶名稱
GCS_BUCKET_NAME = "procurement-attachments-bucket"
GCS_ATTACHMENT_FOLDER = "attachments"

# --- 日誌配置 ---
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# --- 數據源配置 (環境變數優先) ---
# 將憑證路徑設為全域變數，供 Gspread 和 GCS 共用
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
    try:
        GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"] 
    except KeyError:
        logger.error("GSHEETS_CREDENTIALS_PATH is missing in environment variables.")
        st.error("❌ 嚴重錯誤：找不到 GSHEETS_CREDENTIALS_PATH 環境變數。請檢查服務配置。")
        GSHEETS_CREDENTIALS = None 
else:
    # 本地開發或 Streamlit Cloud 備用
    try:
        SHEET_URL = st.secrets["app_config"]["sheet_url"]
        GSHEETS_CREDENTIALS = None # 本地通常依賴預設憑證或 secrets 中的 json 內容
    except KeyError:
        SHEET_URL = None
        GSHEETS_CREDENTIALS = None
        
DATA_SHEET_NAME = "採購總表"
METADATA_SHEET_NAME = "專案設定"


# --- Streamlit 頁面設定 ---
st.set_page_config(
    page_title=f"專案採購小幫手 {APP_VERSION}", 
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS 樣式優化 ---
CUSTOM_CSS = """
<style>
    /* 全域字體與間距調整 */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    /* Expander 樣式優化 */
    .streamlit-expanderContent { 
        padding-left: 1rem !important; 
        padding-right: 1rem !important; 
        padding-bottom: 1rem !important; 
        border-top: 1px solid #444;
    }
    
    /* 自定義標題樣式 */
    .project-header { 
        font-size: 20px !important; 
        font-weight: bold !important; 
        color: #FAFAFA; 
        font-family: 'Source Sans Pro', sans-serif;
    }
    .item-header { 
        font-size: 16px !important; 
        font-weight: 600 !important; 
        color: #E0E0E0; 
        margin-top: 10px;
        margin-bottom: 5px;
        display: block;
    }
    .meta-info { 
        font-size: 14px !important; 
        color: #9E9E9E; 
        font-weight: normal; 
    }
    
    /* 表單元件樣式覆蓋 (Dark Mode 適配) */
    div[data-baseweb="select"] > div, 
    div[data-baseweb="base-input"] > input, 
    div[data-baseweb="input"] > div { 
        background-color: #262730 !important; 
        color: white !important; 
        -webkit-text-fill-color: white !important; 
    }
    div[data-baseweb="popover"], 
    div[data-baseweb="menu"] { 
        background-color: #262730 !important; 
    }
    div[data-baseweb="option"] { 
        color: white !important; 
    }
    
    /* 儀表板指標卡片 */
    .metric-box { 
        padding: 15px 20px; 
        border-radius: 8px; 
        margin-bottom: 15px; 
        background-color: #262730; 
        text-align: center; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        border: 1px solid #444;
    }
    .metric-title { 
        font-size: 14px; 
        color: #B0B0B0; 
        margin-bottom: 8px; 
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .metric-value { 
        font-size: 28px; 
        font-weight: bold; 
        color: #FFFFFF;
    }
    
    /* 按鈕樣式微調 */
    button[kind="secondary"] {
        border: 1px solid #555 !important;
    }
</style>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
"""

# --- 身份驗證與安全函式 ---

def logout():
    """登出函式：清除驗證狀態並重新運行應用程式。"""
    st.session_state["authenticated"] = False
    # 清除相關 Session State
    keys_to_clear = ['data', 'project_metadata', 'edited_dataframes']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

def login_form():
    """渲染登入表單並處理密碼驗證。"""
    # 從 systemd 環境變數中讀取帳密，若無則使用預設值
    DEFAULT_USERNAME = os.environ.get("AUTH_USERNAME", "dev_user")
    DEFAULT_PASSWORD = os.environ.get("AUTH_PASSWORD", "dev_pwd")
    
    credentials = {"username": DEFAULT_USERNAME, "password": DEFAULT_PASSWORD}

    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
        
    if st.session_state["authenticated"]:
        return # 已驗證，跳過表單

    st.markdown("<br><br><br>", unsafe_allow_html=True)
    col_empty_l, col_center, col_empty_r = st.columns([1, 2, 1])
    
    with col_center:
        with st.container(border=True):
            st.title("🔐 系統登入")
            st.markdown("請輸入您的憑證以存取專案採購管理系統。")
            st.markdown("---")
            
            username = st.text_input("用戶名", key="login_username", value=credentials["username"], disabled=True)
            password = st.text_input("密碼", type="password", key="login_password")
            
            col_login_btn, _ = st.columns([1, 2])
            with col_login_btn:
                if st.button("登入", type="primary", use_container_width=True):
                    if username.strip() == credentials["username"].strip() and password == credentials["password"]:
                        st.session_state["authenticated"] = True
                        st.toast("✅ 登入成功！正在載入數據...", icon="🚀")
                        time.sleep(0.5)
                        st.rerun()
                    else:
                        st.error("❌ 用戶名或密碼錯誤，請重試。")
    st.stop() 


# --- GCS 檔案服務函式 (V2.2.5 完整版) ---

def get_storage_client():
    """
    獲取 GCS 客戶端。
    優先使用 JSON 金鑰檔案以支援 generate_signed_url 功能。
    """
    if GSHEETS_CREDENTIALS and os.path.exists(GSHEETS_CREDENTIALS):
        try:
            # 明確使用 Service Account JSON，確保有 Private Key
            return storage.Client.from_service_account_json(GSHEETS_CREDENTIALS)
        except Exception as e:
            logger.error(f"無法從 JSON 建立 GCS Client: {e}")
            return storage.Client() # 降級嘗試
    else:
        return storage.Client() # 使用環境預設憑證

def upload_attachment_to_gcs(file_obj, next_id):
    """
    將上傳的檔案儲存至 Google Cloud Storage。
    
    Args:
        file_obj: Streamlit UploadedFile 物件
        next_id: 下一個報價的 ID，用於檔案命名
        
    Returns:
        str: GCS URI (gs://...) 或 None (如果失敗)
    """
    if file_obj is None:
        return None
        
    try:
        storage_client = get_storage_client()
        bucket = storage_client.bucket(GCS_BUCKET_NAME)
        
        # 建立檔案名稱：attachments/{ID}_{Timestamp}{Extension}
        file_extension = os.path.splitext(file_obj.name)[1]
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        blob_name = f"{GCS_ATTACHMENT_FOLDER}/{next_id}_{timestamp}{file_extension}"
        
        blob = bucket.blob(blob_name)
        
        # 重置檔案指標並上傳
        file_obj.seek(0)
        blob.upload_from_file(
            file_obj, 
            content_type=file_obj.type
        )
        
        gcs_uri = f"gs://{GCS_BUCKET_NAME}/{blob_name}"
        logger.info(f"檔案上傳成功: {gcs_uri}")
        return gcs_uri

    except Exception as e:
        logger.exception("GCS 上傳過程中發生錯誤")
        st.error(f"❌ 附件上傳失敗。請檢查 GCS 權限配置或 Bucket 名稱。錯誤: {e}")
        return None

@st.cache_data(ttl=3600, show_spinner=False)
def generate_signed_url_cached(gcs_uri):
    """
    為私有 GCS 物件生成帶有簽章的臨時 URL (Signed URL)。
    使用 cache 避免重複請求 API，簽章有效期設為 60 分鐘。
    """
    if not gcs_uri or not isinstance(gcs_uri, str):
        return None
        
    # 如果已經是 HTTP 連結，直接返回
    if gcs_uri.startswith("http://") or gcs_uri.startswith("https://"):
        return gcs_uri
        
    # 如果不是 gs:// 格式，視為無效
    if not gcs_uri.startswith("gs://"):
        return None

    try:
        # 解析 gs://bucket_name/blob_name
        # gs:// 部分長度為 5
        path_part = gcs_uri[5:]
        parts = path_part.split('/', 1)
        
        if len(parts) != 2:
            return None
            
        bucket_name = parts[0]
        blob_name = parts[1]
        
        storage_client = get_storage_client()
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        
        # 生成 V4 簽章
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(minutes=60), # 1小時有效
            method="GET"
        )
        return url
        
    except Exception as e:
        logger.error(f"生成 Signed URL 失敗 ({gcs_uri}): {e}")
        # 不在前端顯示錯誤，避免干擾使用者體驗，回傳 None 讓 UI 顯示空白
        return None


# --- 數據讀取與寫入函式 (Gspread 完整實作) ---

@st.cache_data(ttl=600, show_spinner="正在同步 Google Sheets 數據...")
def load_data_from_sheets():
    """
    從 Google Sheets 讀取採購數據與專案設定。
    包含完整的欄位檢查與資料類型轉換。
    """
    if not SHEET_URL:
        st.warning("⚠️ Google Sheets URL 尚未配置，將使用空白數據模式。")
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除', '附件URL'])
        return empty_data, {}

    try:
        # 連線
        if GSHEETS_CREDENTIALS:
            gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        else:
            # 嘗試使用預設憑證 (在 Streamlit Cloud 可能需要 secrets)
            gc = gspread.service_account()
            
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 1. 讀取採購總表 (Data) ---
        try:
            data_ws = sh.worksheet(DATA_SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"找不到工作表: {DATA_SHEET_NAME}")
            raise

        data_records = data_ws.get_all_records()
        data_df = pd.DataFrame(data_records)

        # 欄位完整性檢查與補全
        required_columns = ['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日']
        for col in required_columns:
            if col not in data_df.columns:
                data_df[col] = "" # 補全缺失欄位
        
        if '標記刪除' not in data_df.columns:
            data_df['標記刪除'] = False
            
        if '附件URL' not in data_df.columns:
            data_df['附件URL'] = ""
            
        # 嚴格的資料類型轉換
        data_df['ID'] = pd.to_numeric(data_df['ID'], errors='coerce').astype('Int64')
        data_df['單價'] = pd.to_numeric(data_df['單價'], errors='coerce').fillna(0).astype('float')
        data_df['數量'] = pd.to_numeric(data_df['數量'], errors='coerce').fillna(1).astype('Int64')
        data_df['總價'] = pd.to_numeric(data_df['總價'], errors='coerce').fillna(0).astype('float')
        
        # 布林值處理 (Sheets 有時會回傳 TRUE/FALSE 字串)
        data_df['選取'] = data_df['選取'].apply(lambda x: True if str(x).upper() == 'TRUE' else False)
        data_df['標記刪除'] = data_df['標記刪除'].apply(lambda x: True if str(x).upper() == 'TRUE' else False)

        # --- 2. 讀取專案設定 (Metadata) ---
        try:
            metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
            metadata_records = metadata_ws.get_all_records()
        except gspread.exceptions.WorksheetNotFound:
            st.warning(f"找不到工作表: {METADATA_SHEET_NAME}，將使用預設設定。")
            metadata_records = []
        
        project_metadata = {}
        if metadata_records:
            for row in metadata_records:
                proj_name = row.get('專案名稱')
                if not proj_name: continue
                
                try: 
                    due_date = pd.to_datetime(str(row.get('專案交貨日'))).date()
                except (ValueError, TypeError): 
                    due_date = datetime.now().date()
                
                try:
                    buffer_days = int(row.get('緩衝天數', 7))
                except (ValueError, TypeError):
                    buffer_days = 7
                    
                project_metadata[proj_name] = {
                    'due_date': due_date,
                    'buffer_days': buffer_days,
                    'last_modified': str(row.get('最後修改', ''))
                }

        logger.info(f"成功載入 {len(data_df)} 筆資料，{len(project_metadata)} 個專案設定。")
        st.toast("✅ 數據已從 Google Sheets 更新", icon="☁️")
        return data_df, project_metadata

    except Exception as e:
        logger.exception("Google Sheets 數據載入失敗") 
        st.error(f"❌ 數據載入失敗！請檢查權限或網路連線。錯誤訊息: {e}")
        st.session_state.data_load_failed = True
        return pd.DataFrame(), {}


def write_data_to_sheets(df_to_write, metadata_to_write):
    """
    將 DataFrame 和 Metadata 寫回 Google Sheets。
    執行嚴格的欄位過濾，避免寫入前端輔助欄位 (如: LinkColumn)。
    """
    if st.session_state.get('data_load_failed', False) or not SHEET_URL:
        st.warning("由於載入失敗或未配置 URL，寫入操作已暫停以保護數據。")
        return False
        
    try:
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 1. 寫入採購總表 ---
        # 移除前端顯示用的輔助欄位
        columns_to_exclude = ['標記刪除', '交期顯示', '附件連結']
        df_export = df_to_write.drop(columns=columns_to_exclude, errors='ignore')
        
        # 處理日期物件轉字串 (避免 JSON 序列化錯誤)
        for col in df_export.columns:
            if pd.api.types.is_datetime64_any_dtype(df_export[col]):
                df_export[col] = df_export[col].dt.strftime(DATE_FORMAT)
        
        # 填充 NaN 為空字串
        df_export = df_export.fillna("")
        
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_ws.clear()
        # 寫入 Header 和 Data
        data_ws.update([df_export.columns.values.tolist()] + df_export.values.tolist())
        
        # --- 2. 寫入專案設定 ---
        metadata_list = []
        for name, data in metadata_to_write.items():
            metadata_list.append({
                '專案名稱': name, 
                '專案交貨日': data['due_date'].strftime(DATE_FORMAT),
                '緩衝天數': data['buffer_days'], 
                '最後修改': data['last_modified']
            })
            
        metadata_df = pd.DataFrame(metadata_list)
        metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        metadata_ws.clear()
        if not metadata_df.empty:
            metadata_ws.update([metadata_df.columns.values.tolist()] + metadata_df.values.tolist())
        
        # 清除快取，確保下次讀取最新數據
        st.cache_data.clear() 
        logger.info("數據寫入成功。")
        return True
        
    except Exception as e:
        logger.exception("Google Sheets 寫入失敗")
        st.error(f"❌ 寫入 Google Sheets 失敗！請稍後重試。錯誤: {e}")
        return False


# --- 商業邏輯與計算函式 ---

def add_business_days(start_date, num_days):
    """計算工作日 (跳過週末)。"""
    current_date = start_date
    days_added = 0
    while days_added < num_days:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5: # 0-4 是週一到週五
            days_added += 1
    return current_date

@st.cache_data
def convert_df_to_excel(df):
    """轉換 DataFrame 為 Excel 格式供下載。"""
    # 移除內部欄位
    df_export = df.drop(columns=['標記刪除', '交期顯示', '附件連結'], errors='ignore')
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='採購報價總表')
    return output.getvalue()

@st.cache_data(show_spinner=False)
def calculate_dashboard_metrics(df_state, project_metadata_state):
    """計算儀表板顯示的 KPI 指標。"""
    total_projects = len(project_metadata_state)
    total_budget = 0
    risk_items = 0
    df = df_state.copy()
    
    if df.empty:
        return 0, 0, 0, 0

    # 計算預算邏輯：若有勾選則加總勾選項目，若無則取最小值(預估)
    for _, proj_data in df.groupby('專案名稱'):
        for _, item_df in proj_data.groupby('專案項目'):
            selected_rows = item_df[item_df['選取'] == True]
            if not selected_rows.empty:
                total_budget += selected_rows['總價'].sum()
            elif not item_df.empty:
                # 若尚未選定廠商，暫時加總該項目中最低價的報價
                total_budget += item_df['總價'].min()
    
    # 計算交期風險
    temp_df = df.copy() 
    temp_df['預計交貨日_dt'] = pd.to_datetime(temp_df['預計交貨日'], errors='coerce')
    temp_df['採購最慢到貨日_dt'] = pd.to_datetime(temp_df['採購最慢到貨日'], errors='coerce')
    
    # 風險定義：預計交貨日 > 最慢到貨日
    risk_items = (temp_df['預計交貨日_dt'] > temp_df['採購最慢到貨日_dt']).sum()

    # 待處理報價
    pending_quotes = df[~df['狀態'].isin(['已收貨', '取消'])].shape[0]

    return total_projects, total_budget, risk_items, pending_quotes

def calculate_project_budget(df, project_name):
    """計算單一專案的預算。"""
    proj_df = df[df['專案名稱'] == project_name]
    total_budget = 0
    for _, item_df in proj_df.groupby('專案項目'):
        selected_rows = item_df[item_df['選取'] == True]
        if not selected_rows.empty:
            total_budget += selected_rows['總價'].sum()
        else:
            if not item_df.empty:
                total_budget += item_df['總價'].min()
    return total_budget

@st.cache_data(show_spinner=False)
def calculate_latest_arrival_dates(df, metadata):
    """
    自動計算最慢到貨日。
    公式：專案交貨日 - 緩衝天數
    """
    if df.empty or not metadata:
        return df

    # 將 metadata 轉為 DataFrame 以便 merge
    metadata_df = pd.DataFrame.from_dict(metadata, orient='index')
    metadata_df = metadata_df.reset_index().rename(columns={'index': '專案名稱'})
    metadata_df['due_date'] = metadata_df['due_date'].apply(lambda x: pd.to_datetime(x).date())
    metadata_df['buffer_days'] = metadata_df['buffer_days'].astype(int)

    # 合併資訊
    df = pd.merge(df, metadata_df[['專案名稱', 'due_date', 'buffer_days']], on='專案名稱', how='left')
    
    # 計算
    df['採購最慢到貨日_TEMP'] = (
        pd.to_datetime(df['due_date']) - 
        df['buffer_days'].apply(lambda x: timedelta(days=x))
    )
    df['採購最慢到貨日'] = df['採購最慢到貨日_TEMP'].dt.strftime(DATE_FORMAT)
    
    # 清理暫存欄位
    df = df.drop(columns=['due_date', 'buffer_days', '採購最慢到貨日_TEMP'], errors='ignore')
    return df


# --- UI 事件處理函式 (完整邏輯) ---

def save_and_rerun(df_to_save, metadata_to_save, success_message=""):
    """儲存資料並重新整理頁面。"""
    if write_data_to_sheets(df_to_save, metadata_to_save):
        st.session_state.edited_dataframes = {} # 清空編輯狀態
        if success_message:
            st.success(success_message)
            time.sleep(1) # 給使用者一點時間看訊息
        st.rerun()

def handle_master_save():
    """
    處理主表格的編輯儲存。
    合併所有 `st.data_editor` 的變更，更新主資料表。
    """
    if not st.session_state.edited_dataframes:
        st.info("ℹ️ 沒有偵測到任何表格修改。")
        return

    main_df = st.session_state.data.copy()
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    affected_projects = set()
    changes_detected = False
    
    # 遍歷每個專案項目的編輯結果
    for item_name, edited_df in st.session_state.edited_dataframes.items():
        if edited_df.empty: continue
        
        # 尋找變更的列
        for index, new_row in edited_df.iterrows():
            original_id = new_row['ID']
            # 在主表中找到對應的列
            idx_in_main = main_df[main_df['ID'] == original_id].index
            
            if idx_in_main.empty: continue
            main_idx = idx_in_main[0]
            
            # 定義可被編輯的欄位 (注意：不包含 '附件連結' 等前端欄位)
            updatable_cols = ['選取', '供應商', '單價', '數量', '狀態', '標記刪除', '附件URL'] 
            
            # 比對並更新
            for col in updatable_cols:
                # 確保欄位存在且值有變動
                if col in new_row and main_df.loc[main_idx, col] != new_row[col]:
                    main_df.loc[main_idx, col] = new_row[col]
                    changes_detected = True
            
            # 特殊處理：解析日期顯示字串 (去除後面的 emoji)
            try:
                date_val_full = str(new_row['交期顯示']).strip()
                date_part = date_val_full.split(' ')[0] # 取第一部分
                if main_df.loc[main_idx, '預計交貨日'] != date_part:
                    # 驗證日期格式
                    datetime.strptime(date_part, DATE_FORMAT)
                    main_df.loc[main_idx, '預計交貨日'] = date_part
                    changes_detected = True
            except (ValueError, IndexError): 
                pass # 格式錯誤則忽略
            
            # 自動計算總價 (單價 * 數量)
            try:
                new_total = float(new_row['單價']) * float(new_row['數量'])
                if main_df.loc[main_idx, '總價'] != new_total:
                    main_df.loc[main_idx, '總價'] = new_total
                    changes_detected = True
            except (ValueError, TypeError):
                pass
            
            affected_projects.add(main_df.loc[main_idx, '專案名稱'])

    if changes_detected:
        # 更新 Metadata 的最後修改時間
        for proj in affected_projects:
            if proj in st.session_state.project_metadata:
                st.session_state.project_metadata[proj]['last_modified'] = current_time_str
        
        # 更新 Session State
        st.session_state.data = main_df
        save_and_rerun(st.session_state.data, st.session_state.project_metadata, "✅ 所有變更已儲存！Google Sheets 同步完成。")
    else:
        st.info("ℹ️ 資料未發生實質變更。")

def trigger_delete_confirmation():
    """
    觸發刪除確認流程。
    先將 UI 上的勾選狀態同步到暫存資料，再計算要刪除的筆數。
    """
    # 1. 先建立一份暫存的 data，包含使用者剛剛在 data_editor 勾選的內容
    temp_df = st.session_state.data.copy()
    
    # 收集所有 '標記刪除' 的變更
    deletion_updates = []
    for _, edited_df in st.session_state.edited_dataframes.items():
        if not edited_df.empty:
            # 只取 ID 和 標記刪除 欄位
            subset = edited_df[['ID', '標記刪除']]
            deletion_updates.append(subset)
            
    if deletion_updates:
        # 合併所有更新
        combined_updates = pd.concat(deletion_updates)
        # 設定 index 以便 update
        temp_df.set_index('ID', inplace=True)
        combined_updates.set_index('ID', inplace=True)
        # 更新 temp_df
        temp_df.update(combined_updates)
        temp_df.reset_index(inplace=True)

    # 2. 統計要刪除的 ID
    # 轉換為 boolean 避免型別問題
    temp_df['標記刪除'] = temp_df['標記刪除'].apply(lambda x: True if x == True or str(x).lower() == 'true' else False)
    ids_to_delete = temp_df[temp_df['標記刪除'] == True]['ID'].tolist()
    
    if not ids_to_delete:
        st.warning("⚠️ 沒有項目被標記為刪除。請先在表格右側勾選 '刪除?' 欄位。")
        st.session_state.show_delete_confirm = False
        return

    # 3. 進入確認狀態
    st.session_state.delete_count = len(ids_to_delete)
    st.session_state.ids_pending_delete = ids_to_delete # 暫存要刪除的 ID
    st.session_state.show_delete_confirm = True
    st.rerun()

def handle_batch_delete_quotes():
    """執行實際的刪除操作。"""
    ids_to_delete = st.session_state.get('ids_pending_delete', [])
    
    if not ids_to_delete:
        st.error("找不到待刪除的 ID。")
        st.session_state.show_delete_confirm = False
        st.rerun()
        return
    
    # 執行過濾
    current_data = st.session_state.data
    new_data = current_data[~current_data['ID'].isin(ids_to_delete)].reset_index(drop=True)
    
    # 更新 Session State
    st.session_state.data = new_data
    
    # 重置狀態
    st.session_state.show_delete_confirm = False
    st.session_state.delete_count = 0
    st.session_state.ids_pending_delete = []
    
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已永久刪除 {len(ids_to_delete)} 筆資料。")

def cancel_delete_confirmation():
    """取消刪除操作。"""
    st.session_state.show_delete_confirm = False
    st.session_state.ids_pending_delete = []
    st.rerun()

def handle_project_modification():
    """修改專案名稱與時程。"""
    target_proj = st.session_state.edit_target_project
    new_name = st.session_state.edit_new_name
    new_date = st.session_state.edit_new_date
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    
    if not new_name:
        st.error("❌ 專案名稱不能為空")
        return
        
    # 檢查名稱衝突
    if target_proj != new_name and new_name in st.session_state.project_metadata:
        st.error(f"❌ 新的專案名稱 '{new_name}' 已存在，請使用不同名稱。")
        return

    # 更新 Metadata
    meta = st.session_state.project_metadata.pop(target_proj)
    meta['due_date'] = new_date
    meta['last_modified'] = current_time_str
    st.session_state.project_metadata[new_name] = meta
    
    # 更新 Data 中的專案名稱
    st.session_state.data.loc[st.session_state.data['專案名稱'] == target_proj, '專案名稱'] = new_name
    
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 專案資訊已更新：{new_name}。")

def handle_delete_project(project_to_delete):
    """刪除整個專案。"""
    if not project_to_delete: return
    
    # 移除 Metadata
    if project_to_delete in st.session_state.project_metadata:
        del st.session_state.project_metadata[project_to_delete]
    
    # 移除 Data 中相關列
    original_count = len(st.session_state.data)
    st.session_state.data = st.session_state.data[st.session_state.data['專案名稱'] != project_to_delete].reset_index(drop=True)
    deleted_count = original_count - len(st.session_state.data)
    
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 專案 {project_to_delete} 及其 {deleted_count} 筆報價已刪除。")

def handle_add_new_project():
    """新增專案設定。"""
    project_name = st.session_state.new_proj_name
    project_due_date = st.session_state.new_proj_due_date
    buffer_days = st.session_state.new_proj_buffer_days
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)

    if not project_name:
        st.error("❌ 專案名稱不能為空。")
        return
        
    if project_name in st.session_state.project_metadata:
        st.info(f"ℹ️ 專案 '{project_name}' 已存在，將更新其設定。")

    st.session_state.project_metadata[project_name] = {
        'due_date': project_due_date, 
        'buffer_days': buffer_days,
        'last_modified': current_time_str
    }
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已儲存專案設定：{project_name}。")

def handle_add_new_quote(latest_arrival_date, uploaded_file):
    """新增單筆報價。"""
    project_name = st.session_state.quote_project_select
    item_name_to_use = st.session_state.item_name_to_use_final
    supplier = st.session_state.quote_supplier
    price = st.session_state.quote_price
    qty = st.session_state.quote_qty
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    
    # 決定交貨日期
    if st.session_state.quote_date_type == "1. 指定日期":
        final_delivery_date = st.session_state.quote_delivery_date
    else:
        final_delivery_date = st.session_state.calculated_delivery_date 

    # 驗證輸入
    if not project_name:
        st.error("❌ 請選擇專案。")
        return
    if not item_name_to_use:
        st.error("❌ 請輸入或選擇採購項目。")
        return

    total_price = price * qty
    
    # GCS 上傳流程
    attachment_uri = ""
    next_id = st.session_state.next_id
    
    if uploaded_file is not None:
        with st.spinner(f"正在上傳附件 {uploaded_file.name}..."):
            attachment_uri = upload_attachment_to_gcs(uploaded_file, next_id)
            if attachment_uri is None: 
                # 上傳失敗，中斷流程
                return 

    # 更新專案最後修改時間
    if project_name in st.session_state.project_metadata:
        st.session_state.project_metadata[project_name]['last_modified'] = current_time_str

    # 建立新資料列
    new_row = {
        'ID': st.session_state.next_id, 
        '選取': False, 
        '專案名稱': project_name, 
        '專案項目': item_name_to_use, 
        '供應商': supplier, 
        '單價': price, 
        '數量': qty, 
        '總價': total_price, 
        '預計交貨日': final_delivery_date.strftime(DATE_FORMAT), 
        '狀態': st.session_state.quote_status, 
        '採購最慢到貨日': latest_arrival_date.strftime(DATE_FORMAT), 
        '標記刪除': False,
        '附件URL': attachment_uri 
    }
    
    st.session_state.next_id += 1
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已成功新增報價至 {project_name}！")


# --- 初始化 Session State ---
def initialize_session_state():
    """初始化所有必要的 Session State 變數。"""
    today = datetime.now().date()
    
    # 首次載入數據
    if 'data' not in st.session_state:
        data_df, metadata_dict = load_data_from_sheets()
        st.session_state.data = data_df
        st.session_state.project_metadata = metadata_dict
        
    # 計算下一個 ID
    if not st.session_state.data.empty:
        try:
            current_max = st.session_state.data['ID'].max()
            next_id_val = int(current_max) + 1 if pd.notna(current_max) else 1
        except:
            next_id_val = 1
    else:
        next_id_val = 1
    
    # 初始化其餘變數
    initial_values = {
        'next_id': next_id_val,
        'edited_dataframes': {}, # 儲存每個表格的編輯狀態
        'calculated_delivery_date': today,
        'show_delete_confirm': False,
        'delete_count': 0,
        'ids_pending_delete': []
    }
    
    for key, value in initial_values.items():
        if key not in st.session_state:
            st.session_state[key] = value
            
    # 確保必要欄位存在
    if '標記刪除' not in st.session_state.data.columns: 
        st.session_state.data['標記刪除'] = False
    if '附件URL' not in st.session_state.data.columns: 
        st.session_state.data['附件URL'] = ""


# --- 主應用程式邏輯 ---
def run_app():
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}") 
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    
    initialize_session_state()
    
    # 每次重跑時重新計算最慢到貨日 (確保時程變更時即時反映)
    st.session_state.data = calculate_latest_arrival_dates(st.session_state.data, st.session_state.project_metadata)
    
    if st.session_state.get('data_load_failed', False):
        st.warning("⚠️ 應用程式目前無法從 Google Sheets 載入數據，請檢查連線或配置。")
        
    today = datetime.now().date() 

    # --- 日期格式化 (供前端顯示用，加上紅綠燈) ---
    def format_date_with_icon(row):
        date_str = str(row['預計交貨日'])
        try:
            v_date = pd.to_datetime(row['預計交貨日']).date()
            l_date = pd.to_datetime(row['採購最慢到貨日']).date()
            # 若預計交貨日 > 最慢到貨日，顯示紅燈
            return f"{date_str} 🔴" if v_date > l_date else f"{date_str} ✅"
        except: 
            return date_str

    # 建立顯示用的欄位，不影響原始數據
    if not st.session_state.data.empty:
        st.session_state.data['交期顯示'] = st.session_state.data.apply(format_date_with_icon, axis=1)

    df = st.session_state.data
    # 依專案分組
    if not df.empty:
        project_groups = df.groupby('專案名稱')
    else:
        project_groups = []
    
    # ==========================
    #      側邊欄 (Sidebar)
    # ==========================
    with st.sidebar:
        st.button("🚪 登出系統", on_click=logout, type="secondary", use_container_width=True)
        st.markdown("---")

        # 1. 修改/刪除專案
        with st.expander("✏️ 修改/刪除專案資訊", expanded=False):
            all_projects = sorted(list(st.session_state.project_metadata.keys()))
            if all_projects:
                target_proj = st.selectbox("選擇目標專案", all_projects, key="edit_target_project")
                operation = st.selectbox("選擇操作", ("修改專案資訊", "刪除專案"), key="project_operation_select")
                st.markdown("---")
                
                current_meta = st.session_state.project_metadata.get(target_proj, {'due_date': today})
                
                if operation == "修改專案資訊":
                    st.text_input("新專案名稱", value=target_proj, key="edit_new_name")
                    st.date_input("新專案交貨日", value=current_meta['due_date'], key="edit_new_date")
                    if st.button("確認修改", type="primary", use_container_width=True): 
                        handle_project_modification()
                elif operation == "刪除專案":
                    st.warning(f"⚠️ 確認永久刪除專案 [{target_proj}]？\n此操作不可逆，將同時刪除所有關聯報價。")
                    if st.button("🔥 確認永久刪除", type="secondary", use_container_width=True): 
                        handle_delete_project(target_proj)
            else: 
                st.info("目前無專案資料。")
        
        st.markdown("---")
        
        # 2. 新增專案
        with st.expander("➕ 新增/設定專案時程", expanded=False):
            st.text_input("專案名稱", key="new_proj_name", placeholder="例如: 辦公室升級")
            project_due_date = st.date_input("專案交貨日", value=today + timedelta(days=30), key="new_proj_due_date")
            buffer_days = st.number_input("採購緩衝天數", min_value=0, value=7, key="new_proj_buffer_days")
            
            calc_date = project_due_date - timedelta(days=int(buffer_days))
            st.info(f"📅 計算之最慢到貨日：{calc_date.strftime('%Y-%m-%d')}")
            
            if st.button("💾 儲存專案設定", key="btn_save_proj", use_container_width=True): 
                handle_add_new_project()
        
        st.markdown("---")
        
        # 3. 新增報價 (GCS版)
        with st.expander("➕ 新增報價", expanded=True):
            all_projects_for_quote = sorted(list(st.session_state.project_metadata.keys()))
            latest_arrival_date = today 
            
            if not all_projects_for_quote:
                st.warning("請先在上方新增專案。")
                project_name = None
            else:
                project_name = st.selectbox("歸屬專案", all_projects_for_quote, key="quote_project_select")
                # 取得該專案設定
                current_meta = st.session_state.project_metadata.get(project_name, {'due_date': today, 'buffer_days': 7})
                latest_arrival_date = current_meta['due_date'] - timedelta(days=int(current_meta['buffer_days']))
                st.caption(f"此專案最慢到貨期限: {latest_arrival_date.strftime('%Y-%m-%d')}")

            # 項目選擇 (支援新增)
            unique_items = sorted(st.session_state.data['專案項目'].unique().tolist())
            unique_items = [i for i in unique_items if i] # 過濾空值
            
            selected_item = st.selectbox("採購項目", ['🆕 新增項目...'] + unique_items, key="quote_item_select")
            
            if selected_item == '🆕 新增項目...':
                item_name_to_use = st.text_input("輸入新項目名稱", key="quote_item_new_input")
            else:
                item_name_to_use = selected_item
                
            st.session_state.item_name_to_use_final = item_name_to_use
            
            col_sup, col_pr = st.columns(2)
            with col_sup:
                st.text_input("供應商", key="quote_supplier")
            with col_pr:
                st.number_input("單價", min_value=0, key="quote_price")
                
            st.number_input("數量", min_value=1, value=1, key="quote_qty")
            
            st.markdown("---")
            st.markdown("📆 **預計交貨日設定**")
            date_input_type = st.radio("輸入方式", ("1. 指定日期", "2. 自然日數", "3. 工作日數"), key="quote_date_type", horizontal=True, label_visibility="collapsed")
            
            if date_input_type == "1. 指定日期": 
                st.date_input("選擇日期", today, key="quote_delivery_date") 
            elif date_input_type == "2. 自然日數": 
                num_days = st.number_input("幾天後交貨?", 1, value=7, key="quote_num_days_input")
                st.session_state.calculated_delivery_date = today + timedelta(days=int(num_days))
            elif date_input_type == "3. 工作日數": 
                num_b_days = st.number_input("幾個工作天?", 1, value=5, key="quote_num_b_days_input")
                st.session_state.calculated_delivery_date = add_business_days(today, int(num_b_days))
            
            if date_input_type != "1. 指定日期":
                st.info(f"計算結果：{st.session_state.calculated_delivery_date.strftime('%Y-%m-%d')}")

            st.selectbox("初始狀態", STATUS_OPTIONS, key="quote_status")
            
            st.markdown("📎 **附件上傳**")
            uploaded_file = st.file_uploader("支援 PDF/圖片", type=['pdf', 'jpg', 'jpeg', 'png'], key="new_quote_file_uploader")

            if st.button("📥 新增資料", key="btn_add_quote", type="primary", use_container_width=True):
                handle_add_new_quote(latest_arrival_date, uploaded_file)


    # ==========================
    #      主畫面 (Main)
    # ==========================
    
    # --- 儀表板 Metrics ---
    total_projects, total_budget, risk_items, pending_quotes = calculate_dashboard_metrics(df, st.session_state.project_metadata)

    st.subheader("📊 總覽儀表板")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='metric-box'><div class='metric-title'>專案總數</div><div class='metric-value'>{total_projects}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='metric-box' style='background:#1E3A2F; border-color:#2E5A48'><div class='metric-title'>總預算 (預估/已選)</div><div class='metric-value'>${total_budget:,.0f}</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='metric-box' style='background:#3E2020; border-color:#5A2E2E'><div class='metric-title'>交期風險項</div><div class='metric-value'>{risk_items}</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='metric-box' style='background:#202E3E; border-color:#2E405A'><div class='metric-title'>待處理報價</div><div class='metric-value'>{pending_quotes}</div></div>", unsafe_allow_html=True)
    
    st.markdown("---")
    
    # --- 批次操作工具列 ---
    col_save, col_delete = st.columns([0.8, 0.2])
    is_locked = st.session_state.show_delete_confirm
    
    with col_save:
        if st.button("💾 儲存所有表格修改 (並重新計算總價)", type="primary", disabled=is_locked):
            handle_master_save()
            
    with col_delete:
        if st.button("🗑️ 刪除已標記項目", type="secondary", disabled=is_locked, key="btn_trigger_delete"):
            trigger_delete_confirmation()

    # --- 刪除確認對話框 ---
    if st.session_state.show_delete_confirm:
        st.markdown(
            f"""
            <div style="padding: 1rem; border: 1px solid #ff4b4b; border-radius: 0.5rem; background-color: rgba(255, 75, 75, 0.1); margin-bottom: 1rem;">
                <h4 style="color: #ff4b4b; margin:0;">⚠️ 危險操作確認</h4>
                <p style="margin: 0.5rem 0;">您即將永久刪除 <strong>{st.session_state.delete_count}</strong> 筆資料。此操作無法復原！</p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        col_yes, col_no, _ = st.columns([0.15, 0.15, 0.7])
        with col_yes: 
            if st.button("✅ 確認刪除", key="confirm_delete_yes", type="primary"): 
                handle_batch_delete_quotes()
        with col_no: 
            if st.button("❌ 取消", key="confirm_delete_no"): 
                cancel_delete_confirmation()

    st.markdown("---")

    # --- 專案列表 (核心表格區域) ---
    if not project_groups:
        st.info("👋 歡迎使用！目前沒有資料，請從左側側邊欄新增專案與報價。")
    
    for proj_name, proj_data in project_groups:
        meta = st.session_state.project_metadata.get(proj_name, {})
        proj_budget = calculate_project_budget(df, proj_name)
        last_mod = meta.get('last_modified', 'N/A')
        
        # 專案標題 HTML
        header_html = f"""
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <span class='project-header'>💼 {proj_name}</span> &nbsp;
                <span class='meta-info'>(交期: {meta.get('due_date')})</span>
            </div>
            <div>
                <span class='project-header'>${proj_budget:,.0f}</span>
            </div>
        </div>
        <div style="font-size: 0.8em; color: #666; text-align: right; margin-top: -5px;">最後修改: {last_mod}</div>
        """
        
        with st.expander(label=f"專案：{proj_name}", expanded=False):
            st.markdown(header_html, unsafe_allow_html=True)
            
            # 依項目分組
            for item_name, item_data in proj_data.groupby('專案項目'):
                st.markdown(f"<span class='item-header'>📦 {item_name}</span>", unsafe_allow_html=True)
                
                # --- 準備顯示用的 DataFrame ---
                display_df = item_data.copy()
                display_df['附件連結'] = None
                
                # 預先生成 Signed URL (針對此區塊的資料)
                # 這裡使用 cached function 以提升效能
                for idx, row in display_df.iterrows():
                    uri = row.get('附件URL', '')
                    if uri and isinstance(uri, str) and uri.strip():
                        signed_url = generate_signed_url_cached(uri)
                        if signed_url:
                            display_df.at[idx, '附件連結'] = signed_url

                editor_key = f"ed_{proj_name}_{item_name}"
                
                # --- 定義欄位順序與配置 ---
                # 隱藏原始 GS 路徑，將 '附件連結' 放在顯眼位置
                column_order = [
                    'ID', '選取', '供應商', '單價', '數量', '總價', 
                    '交期顯示', '狀態', '附件連結', '標記刪除', 
                    '附件URL' # 放到最後並隱藏/唯讀
                ]

                edited_df_value = st.data_editor(
                    display_df[column_order],
                    column_config={
                        "ID": st.column_config.NumberColumn(
                            "ID", disabled=True, width="small"
                        ),
                        "選取": st.column_config.CheckboxColumn(
                            "選", width="small", help="勾選以計入預算"
                        ),
                        "供應商": st.column_config.TextColumn(
                            "供應商", width="medium"
                        ),
                        "單價": st.column_config.NumberColumn(
                            "單價", format="$%d", min_value=0
                        ),
                        "數量": st.column_config.NumberColumn(
                            "數", min_value=1, width="small"
                        ),
                        "總價": st.column_config.NumberColumn(
                            "總價", format="$%d", disabled=True
                        ),
                        "交期顯示": st.column_config.TextColumn(
                            "預計交貨日", disabled=False, width="medium", help="格式: YYYY-MM-DD"
                        ),
                        "狀態": st.column_config.SelectboxColumn(
                            "狀態", options=STATUS_OPTIONS, width="small"
                        ),
                        # 核心修復：使用 LinkColumn 顯示簽署後的 URL
                        "附件連結": st.column_config.LinkColumn(
                            "附件檔案", 
                            display_text="📄 開啟附件", 
                            help="點擊在新分頁預覽附件 (有效期1小時)",
                            width="medium"
                        ),
                        "標記刪除": st.column_config.CheckboxColumn(
                            "刪除?", width="small", help="勾選後點擊上方紅色按鈕執行刪除"
                        ),
                        # 隱藏或縮小原始路徑
                        "附件URL": st.column_config.TextColumn(
                            "系統路徑", 
                            disabled=True, 
                            width="small",
                            help="原始 GCS 路徑 (gs://)"
                        ),
                    },
                    hide_index=True,
                    key=editor_key,
                    disabled=is_locked
                )
                
                # 儲存編輯狀態到 Session State
                st.session_state.edited_dataframes[item_name] = edited_df_value 
                
                st.markdown("---")

    # --- 資料匯出區塊 ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("💾 資料備份與匯出")
    
    col_dl, _ = st.columns([0.2, 0.8])
    with col_dl:
        file_name = f'procurement_report_{datetime.now().strftime("%Y%m%d")}.xlsx'
        st.download_button(
            label="📥 下載 Excel 報表", 
            data=convert_df_to_excel(df), 
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )


def main():
    # 進入點：先驗證登入
    login_form()
    
    # 只有驗證通過才執行主程式
    if st.session_state.authenticated:
        run_app() 
        
if __name__ == "__main__":
    main()
