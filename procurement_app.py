import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os 
import json
import gspread
import logging
import time
from google.cloud import storage # 引入 Google Cloud Storage (GCS) 庫

# 確保 openpyxl 庫已安裝 (pip install openpyxl)

# 配置 Streamlit 日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__) # 定義 logger

# --- 應用程式設定 ---
APP_VERSION = "v2.1.6 + Attachments/UI Fix (v2)" # 更新版本號
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"]
DATE_FORMAT = "%Y-%m-%d" # 日期格式
DATETIME_FORMAT = "%Y-%m-%d %H:%M" # 恢復 V2.1.6 時間戳格式

# --- Google Cloud Storage (GCS) 配置 ---
GCS_BUCKET_NAME = "procurement-attachments-bucket" # 請替換為您的 GCS 儲存桶名稱
GCS_ATTACHMENT_FOLDER = "attachments"

# --- 數據源配置 (安全與 Gspread 連線) ---
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
    try:
        GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"] 
    except KeyError:
        logging.error("GCE_SHEET_URL is set, but GSHEETS_CREDENTIALS_PATH is missing.")
        st.error("❌ 錯誤：在 GCE 環境中未找到 GSHEETS_CREDENTIALS_PATH 環境變數。")
        GSHEETS_CREDENTIALS = None 
else:
    # 備用邏輯，本地或 Streamlit Cloud 使用
    try:
        SHEET_URL = st.secrets["app_config"]["sheet_url"]
        GSHEETS_CREDENTIALS = None
    except KeyError:
        SHEET_URL = None
        GSHEETS_CREDENTIALS = None
        
DATA_SHEET_NAME = "採購總表"
METADATA_SHEET_NAME = "專案設定"


st.set_page_config(
    page_title=f"專案採購小幫手 {APP_VERSION}", 
    page_icon="🛠️", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS 樣式修正 (修復中文字型、調整 Expander 標題、新增底部登出區) ---
CUSTOM_CSS = """
<style>
    /* 強制指定中文字型，解決部分環境標題亂碼問題 */
    html, body, [class*="css"] {
        font-family: "Microsoft JhengHei", "Noto Sans TC", "PingFang TC", sans-serif;
    }

    /* 確保 Streamlit 內建標題顯示正確 */
    .st-emotion-cache-18ni7ap.e1nzilvr1 { 
        font-family: "Microsoft JhengHei", "Noto Sans TC", "PingFang TC", sans-serif !important;
    }
    
    .streamlit-expanderContent { padding-left: 1rem !important; padding-right: 1rem !important; padding-bottom: 1rem !important; }
    
    /* 專案標題樣式 (保持 V2.1.6 基礎) */
    .project-header { font-size: 20px !important; font-weight: bold !important; color: #FAFAFA; }
    .item-header { font-size: 16px !important; font-weight: 600 !important; color: #E0E0E0; }
    .meta-info { font-size: 14px !important; color: #9E9E9E; font-weight: normal; }
    
    /* 輸入欄位顏色統一 */
    div[data-baseweb="select"] > div, div[data-baseweb="base-input"] > input, div[data-baseweb="input"] > div { background-color: #262730 !important; color: white !important; -webkit-text-fill-color: white !important; }
    div[data-baseweb="popover"], div[data-baseweb="menu"] { background-color: #262730 !important; }
    div[data-baseweb="option"] { color: white !important; }
    li[aria-selected="true"] { background-color: #FF4B4B !important; color: white !important; }
    
    /* 儀表板卡片樣式 */
    .metric-box { padding: 10px 15px; border-radius: 8px; margin-bottom: 10px; text-align: center; }
    .metric-title { font-size: 14px; color: #9E9E9E; margin-bottom: 5px; }
    .metric-value { font-size: 24px; font-weight: bold; }

    /* 修正 2: 登出按鈕移到底部並縮小 */
    .sidebar-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        width: var(--sidebar-width);
        padding: 10px 15px;
        background-color: var(--background-color);
        z-index: 100; 
        border-top: 1px solid var(--primary-border-color);
    }
    .sidebar-footer button {
        width: 100%;
        padding: 5px; /* 縮小按鈕 */
        font-size: 14px;
    }
    /* 修正 1: 附件預覽 Modal 的樣式 */
    .attachment-preview {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.9);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 1000;
    }
    .attachment-content {
        max-width: 90%;
        max-height: 90%;
        position: relative;
    }
    .attachment-content img, .attachment-content iframe {
        max-width: 100%;
        max-height: 100%;
        display: block;
        border-radius: 8px;
    }
    .close-button {
        position: absolute;
        top: 20px;
        right: 20px;
        color: white;
        font-size: 30px;
        font-weight: bold;
        cursor: pointer;
        background: rgba(0,0,0,0.5);
        border-radius: 50%;
        padding: 5px 10px;
    }
</style>
"""

# ==============================================================================
# GCS 服務相關函式 (新增 - 配合 GCS 功能)
# ==============================================================================

def get_storage_client():
    """獲取 GCS 客戶端。"""
    # 這裡的邏輯與 Gspread 類似，用於認證 GCS
    if GSHEETS_CREDENTIALS and os.path.exists(GSHEETS_CREDENTIALS):
        try:
            return storage.Client.from_service_account_json(GSHEETS_CREDENTIALS)
        except Exception as e:
            logger.error(f"GCS Client initialization failed with JSON: {e}")
            return storage.Client() 
    return storage.Client()

def upload_attachment_to_gcs(file_obj, next_id):
    """將檔案上傳到 GCS 私有儲存桶。"""
    if not file_obj: return None
    try:
        client = get_storage_client()
        bucket = client.bucket(GCS_BUCKET_NAME)
        ext = os.path.splitext(file_obj.name)[1]
        ts = datetime.now().strftime("%Y%m%d%H%M%S")
        blob_name = f"{GCS_ATTACHMENT_FOLDER}/{next_id}_{ts}{ext}" 
        blob = bucket.blob(blob_name)
        file_obj.seek(0)
        
        content_type = file_obj.type if file_obj.type else 'application/octet-stream'
        blob.upload_from_file(file_obj, content_type=content_type)
        
        logger.info(f"Attachment uploaded: {blob_name}")
        return f"gs://{GCS_BUCKET_NAME}/{blob_name}"
    except Exception as e:
        logger.exception("GCS upload failed.")
        st.error(f"❌ 附件上傳失敗。請檢查 GCS 權限配置或 Bucket 名稱。錯誤: {e}")
        return None

@st.cache_data(ttl=3600, show_spinner=False)
def generate_signed_url_cached(gcs_uri):
    """為 GCS 私有物件生成帶有簽章的臨時 URL (Signed URL)。"""
    if not gcs_uri or not gcs_uri.startswith("gs://"): return None
    try:
        path_part = gcs_uri[5:]
        parts = path_part.split('/', 1)
        if len(parts) != 2: return None
            
        client = get_storage_client()
        bucket = client.bucket(parts[0])
        blob = bucket.blob(parts[1])
        
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(hours=1),
            method="GET"
        )
        return url
        
    except Exception as e:
        logger.error(f"Failed to generate Signed URL for {gcs_uri}: {e}")
        return None


# ==============================================================================
# 登入與安全函式
# ==============================================================================

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
            st.title("🔐 請登入以繼續")
            st.markdown("---")
            
            username = st.text_input("用戶名", key="login_username", value=credentials["username"], disabled=True)
            password = st.text_input("密碼", type="password", key="login_password")
            
            if st.button("登入", type="primary"):
                if username.strip() == credentials["username"].strip() and password == credentials["password"]:
                    st.session_state["authenticated"] = True
                    st.toast("✅ 登入成功！")
                    st.rerun()
                else:
                    st.error("用戶名或密碼錯誤。")
            
    st.stop() 


# ==============================================================================
# 數據讀取與寫入函式 (Gspread)
# ==============================================================================

@st.cache_data(ttl=600, show_spinner="連線 Google Sheets...")
def load_data_from_sheets():
    """直接使用 gspread 讀取 Google Sheets 中的數據。"""
    
    if not SHEET_URL:
        st.info("❌ Google Sheets URL 尚未配置。使用空的數據結構。")
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '附件URL', '標記刪除']) # 新增 '附件URL'
        return empty_data, {}

    try:
        # --- 1. 授權與認證 ---
        if not GSHEETS_CREDENTIALS or not os.path.exists(GSHEETS_CREDENTIALS):
             logging.warning("GSHEETS_CREDENTIALS_PATH 未配置或檔案不存在，嘗試使用默認認證。")
             gc = gspread.service_account()
        else:
             gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
            
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 2. 讀取採購總表 (Data) ---
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_records = data_ws.get_all_records()
        data_df = pd.DataFrame(data_records)

        # 確保核心欄位存在 (新增 '附件URL')
        required_cols = ['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除', '附件URL']
        for col in required_cols:
            if col not in data_df.columns: 
                data_df[col] = "" 
                
        # 數據類型轉換與處理 (使用更穩健的方式處理 astype)
        dtype_map = {
            'ID': 'Int64', 
            '選取': 'bool', 
            '單價': 'float', 
            '數量': 'Int64', 
            '總價': 'float',
            '標記刪除': 'bool'
        }
        
        # 確保只對存在的欄位進行 astype
        valid_dtype_map = {col: dtype for col, dtype in dtype_map.items() if col in data_df.columns}

        if valid_dtype_map:
            data_df = data_df.astype(valid_dtype_map, errors='ignore')

        if '標記刪除' not in data_df.columns:
            data_df['標記刪除'] = False

        # V2.1.6 邏輯: 日期保持字串格式
        
        # --- 3. 讀取專案設定 (Metadata) ---
        metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        metadata_records = metadata_ws.get_all_records()
        
        project_metadata = {}
        if metadata_records:
            for row in metadata_records:
                try:
                    due_date = pd.to_datetime(str(row['專案交貨日'])).date()
                except (ValueError, TypeError):
                    due_date = datetime.now().date()
                    
                project_metadata[row['專案名稱']] = {
                    'due_date': due_date,
                    'buffer_days': int(row.get('緩衝天數', 7)),
                    'last_modified': str(row.get('最後修改', ''))
                }

        st.success("✅ 數據已從 Google Sheets 載入！")
        return data_df, project_metadata

    except Exception as e:
        logging.exception("Google Sheets 數據載入時發生致命錯誤！") 
        
        st.error(f"❌ 數據載入失敗！請檢查 Sheets 分享權限、工作表名稱或憑證檔案。")
        st.code(f"錯誤訊息: {e}")
        
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除', '附件URL'])
        st.session_state.data_load_failed = True
        return empty_data, {}


def write_data_to_sheets(df_to_write, metadata_to_write):
    """直接使用 gspread 寫回 Google Sheets。"""
    if st.session_state.get('data_load_failed', False) or not SHEET_URL:
        st.warning("數據載入失敗或 URL 未配置，已禁用寫入 Sheets。")
        return False
        
    try:
        # --- 1. 授權與認證 ---
        if not GSHEETS_CREDENTIALS or not os.path.exists(GSHEETS_CREDENTIALS):
             gc = gspread.service_account()
        else:
             gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)

        sh = gc.open_by_url(SHEET_URL)
        
        # --- 2. 寫入採購總表 (Data) ---
        # 排除 UI 輔助欄位 (新增 '附件連結', '預覽', '最後修改時間' 的排除)
        cols_to_drop = ['標記刪除', '交期顯示', '附件連結', '預覽', '最後修改時間'] 
        df_export = df_to_write.copy()
        for col in cols_to_drop:
            if col in df_export.columns:
                df_export = df_export.drop(columns=[col])

        # V2.1.6 邏輯: 欄位保持字串
        
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_ws.clear()
        data_ws.update([df_export.columns.values.tolist()] + df_export.values.tolist())
        
        # --- 3. 寫入專案設定 (Metadata) ---
        metadata_list = [
            {'專案名稱': name, 
             '專案交貨日': data['due_date'].strftime(DATE_FORMAT),
             '緩衝天數': data['buffer_days'], 
             '最後修改': data['last_modified']}
            for name, data in metadata_to_write.items()
        ]
        metadata_df = pd.DataFrame(metadata_list)
        metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        metadata_ws.clear()
        if not metadata_df.empty:
            metadata_ws.update([metadata_df.columns.values.tolist()] + metadata_df.values.tolist())
            
        st.cache_data.clear() 
        return True
        
    except Exception as e:
        logging.exception("Google Sheets 數據寫入時發生致命錯誤！")
        st.error(f"❌ 數據寫回 Google Sheets 失敗！")
        st.code(f"寫入錯誤訊息: {e}")
        return False


# ==============================================================================
# 輔助函式區
# ==============================================================================

def add_business_days(start_date, num_days):
    """計算工作日 (跳過週末)。"""
    current_date = start_date
    days_added = 0
    while days_added < num_days:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5: days_added += 1
    return current_date

@st.cache_data
def convert_df_to_excel(df):
    """將 DataFrame 轉換為 Excel 二進位檔案 (使用 BytesIO)。"""
    # 排除所有 UI 輔助欄位 (包含新的附件相關欄位)
    df_export = df.drop(columns=['標記刪除', '交期顯示', '附件URL', '附件連結', '預覽', '最後修改時間'], errors='ignore') 
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='採購報價總表')
    
    processed_data = output.getvalue()
    return processed_data


@st.cache_data(show_spinner=False)
def calculate_dashboard_metrics(df_state, project_metadata_state):
    """計算儀表板所需的總體指標。"""
    
    total_projects = len(project_metadata_state)
    total_budget = 0
    risk_items = 0
    df = df_state.copy()
    
    if df.empty:
        return 0, 0, 0, 0

    # 1. 計算總預算
    for _, proj_data in df.groupby('專案名稱'):
        if proj_data['專案名稱'].iloc[0] not in project_metadata_state: continue 
            
        for _, item_df in proj_data.groupby('專案項目'):
            selected_rows = item_df[item_df['選取'] == True]
            if not selected_rows.empty:
                total_budget += selected_rows['總價'].sum()
            elif not item_df.empty:
                total_budget += item_df['總價'].min()
    
    # 2. 計算風險項目 (使用字串轉日期進行比較)
    temp_df_risk = df.copy() 
    temp_df_risk['預計交貨日_dt'] = pd.to_datetime(temp_df_risk['預計交貨日'], errors='coerce')
    temp_df_risk['採購最慢到貨日_dt'] = pd.to_datetime(temp_df_risk['採購最慢到貨日'], errors='coerce')
    risk_items = (temp_df_risk['預計交貨日_dt'] > temp_df_risk['採購最慢到貨日_dt']).sum()
    

    # 3. 計算需要處理的報價數量
    pending_quotes = df[~df['狀態'].isin(['已收貨', '取消'])].shape[0]

    return total_projects, total_budget, risk_items, pending_quotes


def calculate_project_budget(df, project_name):
    # 此函式用於單一專案的預算顯示
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
    """根據專案設定，計算每個採購項目的採購最慢到貨日。(恢復 V2.1.6 邏輯)"""
    
    if df.empty or not metadata:
        return df

    metadata_df = pd.DataFrame.from_dict(metadata, orient='index')
    metadata_df = metadata_df.reset_index().rename(columns={'index': '專案名稱'})
    
    metadata_df['due_date'] = metadata_df['due_date'].apply(lambda x: pd.to_datetime(x).date())
    metadata_df['buffer_days'] = metadata_df['buffer_days'].astype(int)

    df = pd.merge(df, metadata_df[['專案名稱', 'due_date', 'buffer_days']], on='專案名稱', how='left')

    # 將 due_date 轉換為 Timestamp，才能減去 Timedelta
    df['due_date_ts'] = pd.to_datetime(df['due_date'])

    # 計算最慢到貨日 (Timestamp - Timedelta)，並轉回字串
    df['採購最慢到貨日_NEW'] = (
        df['due_date_ts'] - 
        df['buffer_days'].apply(lambda x: timedelta(days=x) if pd.notna(x) and x is not None else timedelta(days=0))
    ).dt.strftime('%Y-%m-%d')
    
    df['採購最慢到貨日'] = df['採購最慢到貨日_NEW']
    
    df = df.drop(columns=['due_date', 'buffer_days', '採購最慢到貨日_NEW', 'due_date_ts'], errors='ignore') 
    
    return df

# ==============================================================================
# UI 邏輯處理函式
# ==============================================================================

def handle_master_save():
    """批次處理所有 data_editor 的修改，並重新計算總價、更新專案時間戳記。(修復寫入失敗回滾)"""
    
    if not st.session_state.edited_dataframes:
        st.info("沒有偵測到表格修改。")
        return

    main_df = st.session_state.data.copy() # 複製一份用於操作
    original_data = st.session_state.data.copy() # 儲存原始數據用於回滾
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    
    affected_projects = set() 
    changes_detected = False
    
    for _, edited_df in st.session_state.edited_dataframes.items():
        if edited_df.empty: continue
        
        for index, new_row in edited_df.iterrows():
            original_id = new_row['ID'] # ID 是欄位
            idx_in_main = main_df[main_df['ID'] == original_id].index
            if idx_in_main.empty: continue
            
            main_idx = idx_in_main[0]
            
            row_changed = False

            # --- 數據比較與更新 ---
            
            # V2.1.6 邏輯: 處理可編輯的交期顯示欄位 (TextColumn)
            try:
                date_str_parts = str(new_row['交期顯示']).strip().split(' ')
                date_part = date_str_parts[0]
                if main_df.loc[main_idx, '預計交貨日'] != date_part:
                    datetime.strptime(date_part, "%Y-%m-%d") # 驗證日期格式
                    main_df.loc[main_idx, '預計交貨日'] = date_part
                    row_changed = True
            except:
                # 如果無法解析日期，則不進行更改
                pass
            
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
                proj = main_df.loc[main_idx, '專案名稱']
                affected_projects.add(proj)
                
    if changes_detected:
        st.session_state.data = main_df # 將修改寫入 session state (待定)
        
        # V2.1.6 邏輯：更新專案層級的最後修改時間
        # 這裡需要複製一份 metadata 進行修改
        updated_metadata = st.session_state.project_metadata.copy()
        for proj in affected_projects:
            if proj in updated_metadata:
                updated_metadata[proj]['last_modified'] = current_time_str
        
        # 寫回 Google Sheets
        if write_data_to_sheets(st.session_state.data, updated_metadata):
            st.session_state.project_metadata = updated_metadata # 成功才持久化 metadata
            st.session_state.edited_dataframes = {} # 清除編輯狀態
            st.success("✅ 資料已儲存！總價、總預算及 Google Sheets 已更新。")
        else:
            # 寫入失敗，回滾 session state.data
            st.session_state.data = original_data 
            st.error("❌ 儲存失敗！資料未寫回 Google Sheets。")
        
        st.rerun()
    else:
        st.info("沒有偵測到表格修改。")


# 批次刪除的觸發函式
def trigger_delete_confirmation():
    """點擊 '刪除已標記項目' 按鈕時，觸發確認流程。(恢復 V2.1.6 邏輯)"""
    
    temp_df = st.session_state.data.copy()
    
    # V2.1.6 邏輯: 合併 edited_dataframes (ID 是欄位)
    combined_edited_df = pd.concat(
        [edited_df.set_index('ID')[['標記刪除']] for edited_df in st.session_state.edited_dataframes.values() if not edited_df.empty],
        axis=0, 
        ignore_index=False
    )
    
    if not combined_edited_df.empty:
        temp_df = temp_df.set_index('ID')
        temp_df.update(combined_edited_df)
        temp_df = temp_df.reset_index()

    # 確保 '標記刪除' 是布林值，否則 tolist() 可能失敗
    temp_df['標記刪除'] = temp_df['標記刪除'].astype(bool) 
    
    ids_to_delete = temp_df[temp_df['標記刪除'] == True]['ID'].tolist()
    
    if not ids_to_delete:
        st.warning("沒有項目被標記為刪除。請先在表格中勾選 '刪除?' 欄位。")
        st.session_state.show_delete_confirm = False
        return

    st.session_state.delete_count = len(ids_to_delete)
    st.session_state.show_delete_confirm = True
    st.rerun()

def handle_batch_delete_quotes():
    """執行批次刪除操作。(修復刪除失敗與回滾)"""
    
    main_df = st.session_state.data.copy()
    original_data = st.session_state.data.copy() # 儲存原始數據用於回滾
    
    # V2.1.6 邏輯: 合併 edited_dataframes (ID 是欄位)
    combined_edited_df = pd.concat(
        [edited_df.set_index('ID')[['標記刪除']] for edited_df in st.session_state.edited_dataframes.values() if not edited_df.empty],
        axis=0, 
        ignore_index=False
    )
    
    if not combined_edited_df.empty:
        main_df = main_df.set_index('ID')
        main_df.update(combined_edited_df)
        main_df = main_df.reset_index()

    main_df['標記刪除'] = main_df['標記刪除'].astype(bool)
    
    ids_to_delete = main_df[main_df['標記刪除'] == True]['ID'].tolist()
    
    if not ids_to_delete:
        st.session_state.show_delete_confirm = False
        st.rerun()
        return

    # 1. 在 Session State 中執行刪除
    df_after_delete = main_df[main_df['標記刪除'] == False].drop(columns=['標記刪除'], errors='ignore')
    st.session_state.data = df_after_delete
    
    # 2. 寫回 Google Sheets
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.session_state.show_delete_confirm = False # 重設確認狀態
        st.success(f"✅ 已成功刪除 {len(ids_to_delete)} 筆報價。Sheets 已更新。")
        st.session_state.edited_dataframes = {} # 清除編輯狀態
    else:
        # 寫入失敗，回滾 session state.data
        st.session_state.data = original_data
        st.error("❌ 刪除失敗！資料未寫回 Google Sheets。請檢查 Sheets 權限。")
    
    st.rerun()

def handle_project_modification():
    """處理修改專案設定的邏輯"""
    target_proj = st.session_state.edit_target_project
    new_name = st.session_state.edit_new_name
    new_date = st.session_state.edit_new_date
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    
    if not new_name:
        st.error("專案名稱不能為空")
        return
        
    if target_proj != new_name and new_name in st.session_state.project_metadata:
        st.error(f"新的專案名稱 '{new_name}' 已存在，請使用不同名稱。")
        return

    # 儲存原始 metadata
    original_metadata = st.session_state.project_metadata.copy()

    # 1. 更新 Metadata
    meta = st.session_state.project_metadata.pop(target_proj)
    meta['due_date'] = new_date
    meta['last_modified'] = current_time_str
    st.session_state.project_metadata[new_name] = meta
    
    # 2. 更新 Dataframe
    st.session_state.data.loc[st.session_state.data['專案名稱'] == target_proj, '專案名稱'] = new_name
    
    # 寫回 Google Sheets
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"✅ 專案已更新：{new_name}。Sheets 已更新。")
    else:
        # 寫入失敗，回滾 metadata
        st.session_state.project_metadata = original_metadata
        st.error("❌ 專案修改失敗！資料未寫回 Google Sheets。")
    
    st.rerun()

def handle_delete_project(project_to_delete):
    """刪除選定的專案及其所有相關報價。"""
    
    if not project_to_delete:
        st.error("請選擇要刪除的專案。")
        return

    # 儲存原始數據用於回滾
    original_data = st.session_state.data.copy()
    original_metadata = st.session_state.project_metadata.copy()

    # 1. 刪除專案設定 (Metadata)
    if project_to_delete in st.session_state.project_metadata:
        del st.session_state.project_metadata[project_to_delete]

    # 2. 刪除所有相關報價 (Data)
    initial_count = len(st.session_state.data)
    st.session_state.data = st.session_state.data[
        st.session_state.data['專案名稱'] != project_to_delete
    ].reset_index(drop=True)
    
    deleted_count = initial_count - len(st.session_state.data)

    # 寫回 Google Sheets
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"✅ 專案 **{project_to_delete}** 及其相關的 {deleted_count} 筆報價已成功刪除。Sheets 已更新。")
    else:
        # 寫入失敗，回滾數據
        st.session_state.data = original_data
        st.session_state.project_metadata = original_metadata
        st.error("❌ 專案刪除失敗！資料未寫回 Google Sheets。")
    
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
        
    if project_name in st.session_state.project_metadata:
        st.warning(f"專案 '{project_name}' 已存在，將更新其時程設定。")
    
    # 儲存原始 metadata
    original_metadata = st.session_state.project_metadata.copy()

    # 更新 session state
    st.session_state.project_metadata[project_name] = {
        'due_date': project_due_date, 
        'buffer_days': buffer_days,
        'last_modified': current_time_str
    }
    
    # 寫回 Google Sheets
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"✅ 已新增/更新專案設定：{project_name}。Sheets 已更新。")
    else:
        # 寫入失敗，回滾 metadata
        st.session_state.project_metadata = original_metadata
        st.error("❌ 新增專案設定失敗！資料未寫回 Google Sheets。")
    
    st.rerun()

def handle_add_new_quote(latest_arrival_date):
    """處理新增報價的邏輯 (恢復 V2.1.6 邏輯)"""
    project_name = st.session_state.quote_project_select
    item_name_to_use = st.session_state.item_name_to_use_final
    supplier = st.session_state.quote_supplier
    price = st.session_state.quote_price
    qty = st.session_state.quote_qty
    status = st.session_state.quote_status
    uploaded_file = st.session_state.new_quote_file_uploader # 修正: 重新引入 file_uploader key

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
        
    # --- 修正 1: GCS 附件上傳 ---
    uri = ""
    if uploaded_file:
        with st.spinner(f"正在上傳附件 {uploaded_file.name}..."):
            uri = upload_attachment_to_gcs(uploaded_file, st.session_state.next_id) or ""
            
    total_price = price * qty
    
    # 儲存原始數據用於回滾
    original_data = st.session_state.data.copy()
    original_metadata = st.session_state.project_metadata.copy()
    
    # 1. 更新 Metadata (暫時更新)
    st.session_state.project_metadata[project_name]['last_modified'] = current_time_str # 更新專案層級時間

    # 2. 新增報價 (暫時更新)
    new_row = {
        'ID': st.session_state.next_id, '選取': False, '專案名稱': project_name, 
        '專案項目': item_name_to_use, '供應商': supplier, '單價': price, '數量': qty, 
        '總價': total_price, 
        '預計交貨日': final_delivery_date.strftime(DATE_FORMAT), # 儲存為字串
        '狀態': status, 
        '採購最慢到貨日': latest_arrival_date.strftime(DATE_FORMAT), # 儲存為字串
        '標記刪除': False,
        '附件URL': uri # 新增 '附件URL'
    }
    st.session_state.next_id += 1
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    
    # 3. 寫回 Google Sheets
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"✅ 已新增報價至 {project_name}！Sheets 已更新。")
    else:
        # 寫入失敗，回滾數據
        st.session_state.data = original_data
        st.session_state.project_metadata = original_metadata
        st.session_state.next_id -= 1 # 回滾 next_id
        st.error("❌ 新增報價失敗！資料未寫回 Google Sheets。")
    
    st.rerun()


# --- Session State 初始化函式 (使用 Gspread 邏輯) ---
def initialize_session_state():
    """初始化所有 Streamlit Session State 變數。從 Sheets 讀取數據。(恢復 V2.1.6 邏輯)"""
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
    if 'preview_url' not in st.session_state: st.session_state.preview_url = None # 新增 preview 狀態
    if 'preview_type' not in st.session_state: st.session_state.preview_type = None # 新增 preview 狀態


# --- 主應用程式核心邏輯 (在登入成功後調用) ---
def run_app():
    """運行應用程式的核心邏輯，在成功登入後調用。"""
    
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}") 
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    initialize_session_state()
    today = datetime.now().date() 

    # 數據自動計算：在初始化後，計算最慢到貨日
    st.session_state.data = calculate_latest_arrival_dates(
        st.session_state.data, 
        st.session_state.project_metadata
    )
    
    if st.session_state.get('data_load_failed', False):
        st.warning("應用程式無法從 Google Sheets 載入數據，請檢查上方錯誤訊息。")
        
    # 格式化日期顯示 (恢復 V2.1.6 邏輯: 處理字串日期)
    def format_date_with_icon(row):
        date_str = str(row['預計交貨日'])
        try:
            v_date = pd.to_datetime(row['預計交貨日']).date()
            l_date = pd.to_datetime(row['採購最慢到貨日']).date()
            icon = "🔴" if v_date > l_date else "✅"
            return f"{date_str} {icon}"
        except:
            return date_str

    if not st.session_state.data.empty:
        st.session_state.data['交期顯示'] = st.session_state.data.apply(format_date_with_icon, axis=1)

    df = st.session_state.data
    project_groups = df.groupby('專案名稱')
    
    # --- 修正 1: 附件預覽 Modal 邏輯 ---
    if st.session_state.preview_url:
        preview_url = st.session_state.preview_url
        preview_type = st.session_state.preview_type
        
        # 使用 markdown 創建一個全屏的 overlay
        st.markdown(
            f"""
            <div class="attachment-preview" id="preview-modal">
                <span class="close-button" onclick="document.getElementById('preview-modal').style.display='none';">×</span>
                <div class="attachment-content">
            """, unsafe_allow_html=True
        )
        
        if preview_type.startswith('image'):
            st.image(preview_url, caption="圖片預覽", use_column_width="always")
        elif preview_type == 'application/pdf':
            # Embed PDF using iframe
            st.markdown(f'<iframe src="{preview_url}" width="800" height="600"></iframe>', unsafe_allow_html=True)
        else:
            st.warning("不支援的預覽類型。")
            
        st.markdown("""
                </div>
            </div>
            """, unsafe_allow_html=True
        )

    
    # *** 側邊欄 UI 邏輯 (使用 with st.sidebar 區塊) ***
    with st.sidebar:
        
        # 區塊 1: 修改/刪除專案
        with st.expander("✏️ 修改/刪除專案資訊", expanded=False): 
            all_projects = sorted(list(st.session_state.project_metadata.keys()))
            
            if all_projects:
                target_proj = st.selectbox("選擇目標專案", all_projects, key="edit_target_project")
                
                operation = st.selectbox(
                    "選擇操作項目", 
                    ("修改專案資訊", "刪除專案"), 
                    key="project_operation_select",
                    help="選擇 '刪除專案' 將永久移除專案及其所有報價。"
                )
                
                st.markdown("---")
                
                current_meta = st.session_state.project_metadata.get(target_proj, {'due_date': today})
                
                if operation == "修改專案資訊":
                    st.markdown("##### ✏️ 專案資訊修改")
                    st.text_input("新專案名稱", value=target_proj, key="edit_new_name")
                    st.date_input("新專案交貨日", value=current_meta['due_date'], key="edit_new_date")
                    
                    if st.button("確認修改專案", type="primary", use_container_width=True): 
                        handle_project_modification()
                
                elif operation == "刪除專案":
                    st.markdown("##### 🗑️ 專案刪除 (⚠️ 警告)")
                    st.warning(f"您即將永久刪除專案 **{target_proj}** 及其所有相關報價資料。")
                    
                    if st.button(f"確認永久刪除 {target_proj}", type="secondary", help="此操作不可逆，將同時移除所有相關報價", use_container_width=True):
                        handle_delete_project(target_proj)
                        
            else: 
                st.info("無專案可修改/刪除。請在下方新增專案。")
        
        st.markdown("---")
        
        # 區塊 2: 新增/設定專案時程
        with st.expander("➕ 新增/設定專案時程", expanded=False): 
            st.text_input("專案名稱 (Project Name)", key="new_proj_name")
            
            project_due_date = st.date_input("專案交貨日 (Project Due Date)", value=today + timedelta(days=30), key="new_proj_due_date")
            buffer_days = st.number_input("採購緩衝天數 (天)", min_value=0, value=7, key="new_proj_buffer_days")
            
            latest_arrival_date_proj = project_due_date - timedelta(days=int(buffer_days))
            st.caption(f"計算得出最慢到貨日：{latest_arrival_date_proj.strftime('%Y年%m月%d日')}")

            if st.button("💾 儲存專案設定", key="btn_save_proj", use_container_width=True):
                handle_add_new_project()
        
        st.markdown("---")
        
        # 區塊 3: 新增報價
        with st.expander("➕ 新增報價", expanded=False): 
            all_projects_for_quote = sorted(list(st.session_state.project_metadata.keys()))
            latest_arrival_date = today 
            
            if not all_projects_for_quote:
                st.warning("請先在上方新增/設定專案時程。")
                project_name = None
            else:
                project_name = st.selectbox("選擇目標專案", all_projects_for_quote, key="quote_project_select")
                
                current_meta = st.session_state.project_metadata.get(project_name, {'due_date': today, 'buffer_days': 7})
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
            st.number_input("單價 (TWD)", min_value=0.0, key="quote_price") 
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
            
            # --- 修正 1: 附件上傳 UI ---
            st.markdown("---")
            st.markdown("##### 📎 附件上傳 (圖片/PDF)")
            st.file_uploader("選擇檔案", type=['png', 'jpg', 'jpeg', 'pdf'], key="new_quote_file_uploader")
            
            if st.button("📥 新增資料", key="btn_add_quote", type="primary", use_container_width=True):
                handle_add_new_quote(latest_arrival_date)


        # --- 修正 2: 登出按鈕移到底部 ---
        st.markdown('<div class="sidebar-footer">', unsafe_allow_html=True)
        st.button("🚪 登出系統", on_click=logout, type="secondary", key="sidebar_logout_btn") # 使用新的 key
        st.markdown('</div>', unsafe_allow_html=True)


    # --- 主介面 ---
    
    # *** 儀表板區塊 ***
    total_projects, total_budget, risk_items, pending_quotes = calculate_dashboard_metrics(df, st.session_state.project_metadata)

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
    
    # *** 批次操作區塊 ***
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

    # *** 專案 Expander 列表 ***
    if project_groups is not None:
        for proj_name, proj_data in project_groups:
            meta = st.session_state.project_metadata.get(proj_name, {})
            proj_budget = calculate_project_budget(df, proj_name)
            
            # V2.1.6 樣式
            last_modified_proj = meta.get('last_modified', 'N/A')
            if not last_modified_proj.strip(): last_modified_proj = 'N/A'
                 
            header_html = f"""
            <span class='project-header'>💼 專案: {proj_name}</span> &nbsp;|&nbsp; 
            <span class='project-header'>總預算: ${proj_budget:,.0f}</span> &nbsp;|&nbsp; 
            <span class='meta-info'>交期: {meta.get('due_date')}</span> 
            <span style='float:right; font-size:14px; color:#FFC107;'>🕒 最後修改: {last_modified_proj}</span>
            """
            
            expander_key = f"expander_{proj_name}"

            # 修正 1: Expander 展開/收合狀態持久化
            if expander_key not in st.session_state:
                st.session_state[expander_key] = False # 預設為收合
            
            def toggle_expander_state(key):
                st.session_state[key] = not st.session_state[key]

            # 為了讓 Expander 點擊時能單獨控制狀態，我們使用 on_change 配合一個不可見的按鈕或狀態管理。
            # Streamlit Expander 的行為很難在 data_editor 內事件觸發時維持，
            # 最直接的方式是依賴 Streamlit 的內建狀態，但如果崩潰，則需要手動處理。
            
            # 這裡我們使用一個技巧：檢查是否在刪除確認中，如果是則保持狀態不變
            is_expanded = st.session_state.get(expander_key, False)
            
            # 在 Streamlit 中，data_editor 的交互會觸發重新運行，這會重新繪製 Expander。
            # 為了防止第一次互動時 Expander 重新讀取 False 狀態而關閉，
            # 我們需要確保它的狀態在每次運行時都能被保留。
            # 由於我們沒有使用 on_change，我們依賴 Streamlit 的 key 機制。

            # 檢查 Expander 的狀態是否因為 data_editor 互動而丟失 (簡單地嘗試保留狀態)
            # 在 data_editor 內互動會導致整個應用程式重新運行。
            # 在重新運行中，Expander 嘗試根據 key 恢復其內部狀態。
            # 這裡我們使用一個簡單的 trick: 讓 Expander 在 data_editor 互動時使用一個固定的狀態，而不是默認的 False

            # 設置 Expander 的 key 和初始狀態
            with st.expander(label=f"專案：{proj_name} (點擊展開)", expanded=is_expanded, key=expander_key): 
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
                    
                    editor_key = f"editor_{proj_name}_{item_name}"
                    
                    # --- 修正 1: 附件連結與預覽按鈕邏輯 ---
                    editable_df['附件連結'] = None
                    editable_df['預覽'] = None
                    
                    # 使用 lambda 函數來確保 set_preview 傳遞正確的參數
                    def create_preview_callback(url, mime_type):
                        return lambda: set_preview(url, mime_type)
                    
                    def set_preview(url, mime_type):
                        # 設置狀態並觸發重新運行以顯示 Modal
                        st.session_state.preview_url = url
                        st.session_state.preview_type = mime_type
                    
                    # 預處理附件連結
                    for idx_orig in editable_df.index: # 使用原始索引來操作 df
                        gcs_uri = editable_df.loc[idx_orig, '附件URL']
                        if gcs_uri:
                            # 獲取簽名 URL
                            signed_url = generate_signed_url_cached(gcs_uri)
                            if signed_url:
                                # 簡化 MIME 類型判斷
                                mime_type = 'application/pdf' if gcs_uri.lower().endswith('.pdf') else 'image/jpeg' 
                                editable_df.loc[idx_orig, '附件連結'] = signed_url
                                
                                # 創建預覽按鈕 (使用 Streamlit 的按鈕)
                                item_id = editable_df.loc[idx_orig, 'ID'] 
                                st.button("📄 預覽", 
                                          key=f"preview_btn_{item_id}", 
                                          on_click=create_preview_callback(signed_url, mime_type),
                                          help="點擊預覽圖片或 PDF",
                                          disabled=is_locked)

                    # 選擇要顯示的欄位 (新增 '附件連結')
                    cols_to_display = ['ID', '選取', '供應商', '單價', '數量', '總價', '交期顯示', '狀態', '附件連結', '標記刪除']


                    edited_df_value = st.data_editor(
                        editable_df[cols_to_display],
                        column_config={
                            "ID": st.column_config.Column("ID", disabled=True, width="tiny"), 
                            "選取": st.column_config.CheckboxColumn("選取", width="tiny"), 
                            "供應商": st.column_config.Column("供應商", disabled=False), 
                            "單價": st.column_config.NumberColumn("單價", format="$%d"),
                            "數量": st.column_config.NumberColumn("數量"),
                            "總價": st.column_config.NumberColumn("總價", format="$%d", disabled=True),
                            
                            "交期顯示": st.column_config.TextColumn("預計交貨日 (YYYY-MM-DD)", width="medium", help="可編輯，圖示會自動更新"), 
                            
                            "狀態": st.column_config.SelectboxColumn("狀態", options=STATUS_OPTIONS),
                            
                            # 新增附件連結欄位
                            "附件連結": st.column_config.LinkColumn("附件連結", display_text="連結", width="small", help="點擊開啟附件連結"),
                            
                            "標記刪除": st.column_config.CheckboxColumn("刪除?", width="tiny", help="勾選後點擊上方按鈕執行刪除"), 
                        },
                        key=editor_key,
                        hide_index=True, 
                        use_container_width=True,
                        height=150 + (len(item_data) * 35) if len(item_data) > 3 else 150,
                        disabled=is_locked
                    )
                    
                    # 傳遞附件URL (雖然沒有編輯，但為了防止遺失，傳遞回 session_state)
                    if '附件URL' in editable_df.columns:
                        edited_df_value['附件URL'] = editable_df['附件URL']

                    st.session_state.edited_dataframes[item_name] = edited_df_value 
                    st.markdown("---")

    # *** 資料匯出區塊 ***
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("💾 資料匯出")
    st.download_button("📥 下載 Excel 報表", 
                      convert_df_to_excel(df), 
                      f'procurement_report_{datetime.now().strftime("%Y%m%d")}.xlsx', 
                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# --- 程式進入點 ---
def main():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True) 
        
    login_form()
    
    if st.session_state.authenticated:
        run_app() 
        
if __name__ == "__main__":
    main()
