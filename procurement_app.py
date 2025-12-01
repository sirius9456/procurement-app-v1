import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os 
import json
import gspread
import logging
import time

# ==============================================================================
# 依賴庫導入與環境初始化
# ==============================================================================

# 引入 Google Cloud Storage (GCS) 庫，用於附件上傳與簽章 URL 生成
from google.cloud import storage

# 確保 openpyxl 庫已安裝 (用於 Excel 匯出)

# --- 應用程式設定與常數定義 ---
# 依照您的要求，版本號回復並鎖定在 v2.2.5
APP_VERSION = "v2.2.5" 
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"] # 報價狀態選項
DATE_FORMAT = "%Y-%m-%d"                            # 日期標準格式
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"               # 時間戳標準格式

# --- Google Cloud Storage (GCS) 配置 ---
# ⚠️ 注意：請確保您的 GCE 服務帳戶有 Storage Object Admin/Creator 權限
GCS_BUCKET_NAME = "procurement-attachments-bucket"
GCS_ATTACHMENT_FOLDER = "attachments"

# --- 日誌配置：用於 Streamlit 後端紀錄，方便除錯 ---
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# --- 數據源配置：優先從環境變數讀取 Google Sheets 資訊 ---
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
    try:
        # GSheets 憑證路徑，用於 Gspread 連線與 GCS Signed URL 簽署
        GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"] 
    except KeyError:
        logger.error("GSHEETS_CREDENTIALS_PATH is missing in environment.")
        st.error("❌ 嚴重錯誤：找不到 GSHEETS_CREDENTIALS_PATH 環境變數。")
        GSHEETS_CREDENTIALS = None 
else:
    # 非 GCE 環境下的備用配置
    try:
        SHEET_URL = st.secrets["app_config"]["sheet_url"]
        GSHEETS_CREDENTIALS = None 
    except KeyError:
        SHEET_URL = None
        GSHEETS_CREDENTIALS = None
        
DATA_SHEET_NAME = "採購總表"     # 存放報價數據的工作表名稱
METADATA_SHEET_NAME = "專案設定" # 存放專案設定 (交期, 緩衝天數) 的工作表名稱


# --- Streamlit 頁面設定 ---
st.set_page_config(
    page_title=f"專案採購小幫手 {APP_VERSION}", 
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS 樣式優化 ---
# 恢復較為基礎的樣式，取消 V2.1.6 的特殊格式
CUSTOM_CSS = """
<style>
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    
    .status-ok { color: #4CAF50; font-weight: bold; }
    .status-risk { color: #FF4B4B; font-weight: bold; }
    
    /* 調整 DataFrame 編輯器的字體大小 */
    .stDataFrame { font-size: 14px; }
</style>
"""


# ==============================================================================
# 輔助函式區
# ==============================================================================

# --- 身份驗證與登出 ---
def logout():
    """清除 Session State 並重新啟動應用程式以登出。"""
    st.session_state["authenticated"] = False
    for key in ['data', 'project_metadata', 'edited_dataframes']:
        if key in st.session_state: del st.session_state[key]
    st.rerun()

def login_form():
    """顯示登入表單並進行密碼驗證。"""
    DEFAULT_USERNAME = os.environ.get("AUTH_USERNAME", "dev_user")
    DEFAULT_PASSWORD = os.environ.get("AUTH_PASSWORD", "dev_pwd")
    credentials = {"username": DEFAULT_USERNAME, "password": DEFAULT_PASSWORD}

    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
        
    if st.session_state["authenticated"]: return

    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col_center, _ = st.columns([1, 2, 1])
    with col_center:
        with st.container(border=True):
            st.title("🔐 系統登入")
            username = st.text_input("用戶名", value=credentials["username"], disabled=True)
            password = st.text_input("密碼", type="password")
            if st.button("登入", type="primary", use_container_width=True):
                if username.strip() == credentials["username"].strip() and password == credentials["password"]:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("密碼錯誤")
    st.stop() 


# --- GCS 服務相關函式 ---
def get_storage_client():
    """
    獲取 GCS 客戶端。優先使用 Service Account JSON 檔案進行驗證。
    """
    if GSHEETS_CREDENTIALS and os.path.exists(GSHEETS_CREDENTIALS):
        try:
            return storage.Client.from_service_account_json(GSHEETS_CREDENTIALS)
        except Exception as e:
            logger.error(f"GCS Client initialization failed with JSON: {e}")
            return storage.Client() 
    return storage.Client()

def upload_attachment_to_gcs(file_obj, next_id):
    """
    將檔案上傳到 GCS 私有儲存桶。
    """
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
    """
    為 GCS 私有物件生成帶有簽章的臨時 URL (Signed URL)。
    """
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

# --- 數據工具 ---
def add_business_days(start_date, num_days):
    """計算工作日 (跳過週末)。"""
    current_date = start_date
    days_added = 0
    while days_added < num_days:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5: 
            days_added += 1
    return current_date

@st.cache_data
def convert_df_to_excel(df):
    """轉換 DataFrame 為 Excel 格式供下載。"""
    # 這裡保留所有欄位，因為要恢復完整性
    df_export = df.copy().fillna("")
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='採購報價總表')
    return output.getvalue()


# --- Gspread 數據處理 ---
@st.cache_data(ttl=600, show_spinner="正在同步 Google Sheets 數據...")
def load_data_from_sheets():
    """從 Google Sheets 載入所有採購數據與專案設定。"""
    if not SHEET_URL:
        st.warning("⚠️ Google Sheets URL 尚未配置。")
        return pd.DataFrame(), {}
        
    try:
        # 連線認證
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS) if GSHEETS_CREDENTIALS else gspread.service_account()
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 1. 讀取採購總表 (Data) ---
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_records = data_ws.get_all_records()
        data_df = pd.DataFrame(data_records)

        # 確保核心欄位存在
        required_cols = ['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '最後修改時間', '附件URL', '標記刪除']
        for col in required_cols:
            if col not in data_df.columns: 
                data_df[col] = "" 
        
        # 資料類型轉換
        data_df['ID'] = pd.to_numeric(data_df['ID'], errors='coerce').astype('Int64')
        data_df['單價'] = pd.to_numeric(data_df['單價'], errors='coerce').fillna(0).astype('float')
        data_df['數量'] = pd.to_numeric(data_df['數量'], errors='coerce').fillna(1).astype('Int64')
        data_df['總價'] = pd.to_numeric(data_df['總價'], errors='coerce').fillna(0).astype('float')
        
        data_df['選取'] = data_df['選取'].astype(str).str.upper() == 'TRUE'
        data_df['標記刪除'] = data_df['標記刪除'].astype(str).str.upper() == 'TRUE'
        
        # 恢復 V2.2.5 邏輯：不強制轉換日期為 datetime 物件，避免編輯器格式問題
        # 保持為字串或讓 Pandas 自動推斷，但不在 data_editor 中強制指定 DateColumn
        
        logger.info(f"Loaded {len(data_df)} records.")

        # --- 2. 讀取專案設定 (Metadata) ---
        meta_records = sh.worksheet(METADATA_SHEET_NAME).get_all_records()
        project_metadata = {}
        for row in meta_records:
            name = row.get('專案名稱')
            if name:
                project_metadata[name] = {
                    'due_date': pd.to_datetime(str(row.get('專案交貨日'))).date(),
                    'buffer_days': int(row.get('緩衝天數', 7)),
                    'last_modified': str(row.get('最後修改', ''))
                }
        
        st.toast("✅ 數據已從 Google Sheets 更新", icon="☁️")
        return data_df, project_metadata
    except Exception as e:
        logger.exception("Google Sheets 數據載入失敗") 
        st.error(f"❌ 數據載入失敗！錯誤: {e}")
        st.session_state.data_load_failed = True
        return pd.DataFrame(), {}

def write_data_to_sheets(df, meta):
    """將修改後的數據與專案設定寫回 Google Sheets。"""
    if st.session_state.get('data_load_failed', False): return False
    try:
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 1. 寫入採購總表 ---
        # 移除前端純顯示的欄位，保留原始欄位 (包含 附件URL)
        cols_to_drop = ['交期狀態', '附件連結'] 
        # 注意：V2.2.5 應該會保留 '標記刪除' 狀態寫回，或者由 delete function 處理
        # 這裡為了安全，我們過濾掉 UI 產生的臨時欄位
        
        export_df = df.copy()
        for c in cols_to_drop:
            if c in export_df.columns:
                export_df = export_df.drop(columns=[c])
        
        export_df = export_df.fillna("")
        
        # 確保日期格式正確
        for col in export_df.columns:
            if pd.api.types.is_datetime64_any_dtype(export_df[col]):
                export_df[col] = export_df[col].dt.strftime(DATE_FORMAT)

        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_ws.clear()
        data_ws.update([export_df.columns.tolist()] + export_df.values.tolist())
        
        # --- 2. 寫入專案設定 ---
        meta_list = [{'專案名稱': k, '專案交貨日': v['due_date'].strftime(DATE_FORMAT), '緩衝天數': v['buffer_days'], '最後修改': v['last_modified']} for k,v in meta.items()]
        meta_df = pd.DataFrame(meta_list)
        
        metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        metadata_ws.clear()
        if not meta_df.empty:
            metadata_ws.update([meta_df.columns.tolist()] + meta_df.values.tolist())
            
        st.cache_data.clear()
        logger.info("Data successfully written to Google Sheets.")
        return True
    except Exception as e:
        logger.exception("Google Sheets write operation failed.")
        st.error(f"❌ 數據寫回失敗！請檢查權限。錯誤: {e}")
        return False


# --- 數據計算與指標 ---
def calculate_latest_arrival(df, meta):
    """計算每個採購項目的最慢到貨日 (專案交期 - 緩衝天數)。"""
    if df.empty or not meta: return df
    meta_df = pd.DataFrame.from_dict(meta, orient='index').reset_index().rename(columns={'index': '專案名稱'})
    meta_df['due_date'] = pd.to_datetime(meta_df['due_date']).dt.date
    df = pd.merge(df, meta_df[['專案名稱', 'due_date', 'buffer_days']], on='專案名稱', how='left')
    
    # 計算邏輯
    df['temp_due_date'] = pd.to_datetime(df['due_date'])
    df['採購最慢到貨日'] = (df['temp_due_date'] - pd.to_timedelta(df['buffer_days'].astype(int), unit='D')).dt.strftime(DATE_FORMAT)
    
    return df.drop(columns=['due_date', 'buffer_days', 'temp_due_date'], errors='ignore')

def calculate_project_budget(df, project_name):
    """計算單一專案的總預算。"""
    proj_df = df[df['專案名稱'] == project_name]
    budget = 0
    for _, item in proj_df.groupby('專案項目'):
        sel = item[item['選取'] == True]
        budget += sel['總價'].sum() if not sel.empty else item['總價'].min()
    return budget

def calculate_metrics(df, meta):
    """計算儀表板的總體指標。"""
    if df.empty: return 0, 0, 0, 0
    total_projects = len(meta)
    
    budget = 0
    for _, proj in df.groupby('專案名稱'):
        for _, item in proj.groupby('專案項目'):
            sel = item[item['選取'] == True]
            budget += sel['總價'].sum() if not sel.empty else item['總價'].min()
            
    risk = (pd.to_datetime(df['預計交貨日'], errors='coerce') > pd.to_datetime(df['採購最慢到貨日'], errors='coerce')).sum()
    pending = df[~df['狀態'].isin(['已收貨', '取消'])].shape[0]
    return total_projects, budget, risk, pending


# ==============================================================================
# UI 事件處理與數據流控制
# ==============================================================================

def save_and_rerun(df, meta, msg=""):
    """儲存資料到 Sheets 並重新整理 UI。"""
    if write_data_to_sheets(df, meta):
        st.session_state.edited_dataframes = {}
        if msg: 
            st.toast(msg, icon="✅")
            time.sleep(1)
        st.rerun()

def handle_master_save():
    """處理所有表格編輯，更新總價並寫回 Sheets。"""
    if not st.session_state.edited_dataframes:
        st.info("無變更")
        return

    main_df = st.session_state.data.copy()
    now_str = datetime.now().strftime(DATETIME_FORMAT)
    changes = False
    
    for _, edited_df in st.session_state.edited_dataframes.items():
        if edited_df.empty: continue
        
        for _, new_row in edited_df.iterrows():
            idx = main_df[main_df['ID'] == new_row['ID']].index
            if idx.empty: continue
            idx = idx[0] 
            
            row_changed = False
            # 檢查欄位
            check_cols = ['選取', '供應商', '單價', '數量', '狀態', '標記刪除', '預計交貨日']
            
            for col in check_cols:
                val_main = main_df.loc[idx, col]
                val_new = new_row.get(col)
                
                # 簡單字串比較，不做複雜的日期物件轉換，避免錯誤
                if str(val_main) != str(val_new):
                    main_df.loc[idx, col] = val_new
                    row_changed = True
            
            # 重新計算總價
            new_total = float(new_row['單價']) * float(new_row['數量'])
            if main_df.loc[idx, '總價'] != new_total:
                main_df.loc[idx, '總價'] = new_total
                row_changed = True
            
            if row_changed:
                main_df.loc[idx, '最後修改時間'] = now_str
                changes = True
                
                # 更新專案 metadata 時間
                proj = main_df.loc[idx, '專案名稱']
                if proj in st.session_state.project_metadata:
                    st.session_state.project_metadata[proj]['last_modified'] = now_str

    if changes:
        st.session_state.data = main_df
        save_and_rerun(st.session_state.data, st.session_state.project_metadata, "✅ 儲存成功！")
    else:
        st.info("ℹ️ 未偵測到實質變更。")

def handle_add_new_quote(latest_arrival, file):
    """處理新增報價邏輯。"""
    proj = st.session_state.quote_project_select
    item = st.session_state.item_name_to_use_final
    if not proj or not item:
        st.error("❌ 請填寫專案名稱和採購項目。")
        return

    uri = ""
    if file:
        with st.spinner(f"正在上傳附件 {file.name}..."):
            uri = upload_attachment_to_gcs(file, st.session_state.next_id) or ""

    now_str = datetime.now().strftime(DATETIME_FORMAT)
    
    # 決定預計交貨日
    if st.session_state.quote_date_type == "1. 指定日期":
        del_date = st.session_state.quote_delivery_date
    else:
        del_date = st.session_state.calculated_delivery_date

    # 建立新資料列
    new_row = {
        'ID': st.session_state.next_id, '選取': False, '專案名稱': proj, '專案項目': item,
        '供應商': st.session_state.quote_supplier, '單價': st.session_state.quote_price,
        '數量': st.session_state.quote_qty, '總價': st.session_state.quote_price * st.session_state.quote_qty,
        '預計交貨日': del_date.strftime(DATE_FORMAT), 
        '狀態': st.session_state.quote_status,
        '採購最慢到貨日': latest_arrival, # 這裡 latest_arrival 已經是字串
        '最後修改時間': now_str, 
        '標記刪除': False, '附件URL': uri
    }
    
    st.session_state.next_id += 1
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    st.session_state.project_metadata[proj]['last_modified'] = now_str
    
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已成功新增報價至 {proj}！")

def handle_project_modification():
    """處理專案名稱或交貨日的修改。"""
    target_proj = st.session_state.edit_target_project
    new_name = st.session_state.edit_new_name
    new_date = st.session_state.edit_new_date
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)
    
    if not new_name:
        st.error("❌ 專案名稱不能為空")
        return
    if target_proj != new_name and new_name in st.session_state.project_metadata:
        st.error(f"❌ 新的專案名稱 '{new_name}' 已存在。")
        return

    meta = st.session_state.project_metadata.pop(target_proj)
    meta['due_date'] = new_date
    meta['last_modified'] = current_time_str
    st.session_state.project_metadata[new_name] = meta
    
    st.session_state.data.loc[st.session_state.data['專案名稱'] == target_proj, '專案名稱'] = new_name
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 專案資訊已更新：{new_name}。")

def handle_delete_project(project_to_delete):
    """永久刪除整個專案及其所有關聯報價。"""
    if not project_to_delete: return
    
    if project_to_delete in st.session_state.project_metadata:
        del st.session_state.project_metadata[project_to_delete]
    
    original_count = len(st.session_state.data)
    st.session_state.data = st.session_state.data[st.session_state.data['專案名稱'] != project_to_delete].reset_index(drop=True)
    deleted_count = original_count - len(st.session_state.data)
    
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 專案 {project_to_delete} 及其 {deleted_count} 筆報價已刪除。")

def handle_add_new_project():
    """處理新增專案設定。"""
    project_name = st.session_state.new_proj_name
    project_due_date = st.session_state.new_proj_due_date
    buffer_days = st.session_state.new_proj_buffer_days
    current_time_str = datetime.now().strftime(DATETIME_FORMAT)

    if not project_name:
        st.error("❌ 專案名稱不能為空。")
        return
        
    st.session_state.project_metadata[project_name] = {
        'due_date': project_due_date, 
        'buffer_days': buffer_days,
        'last_modified': current_time_str
    }
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已儲存專案設定：{project_name}。")

def trigger_delete_confirmation():
    """觸發確認刪除標記項目的流程。"""
    temp_df = st.session_state.data.copy()
    
    deletion_updates = []
    for _, edited_df in st.session_state.edited_dataframes.items():
        if not edited_df.empty and '標記刪除' in edited_df.columns:
            deletion_updates.append(edited_df[['ID', '標記刪除']])
            
    if deletion_updates:
        combined_updates = pd.concat(deletion_updates)
        temp_df.set_index('ID', inplace=True)
        combined_updates.set_index('ID', inplace=True)
        temp_df.update(combined_updates)
        temp_df.reset_index(inplace=True)

    temp_df['標記刪除'] = temp_df['標記刪除'].apply(lambda x: True if x == True or str(x).lower() == 'true' else False)
    ids_to_delete = temp_df[temp_df['標記刪除'] == True]['ID'].tolist()
    
    if not ids_to_delete:
        st.warning("⚠️ 沒有項目被標記為刪除。")
        st.session_state.show_delete_confirm = False
        return

    st.session_state.delete_count = len(ids_to_delete)
    st.session_state.ids_pending_delete = ids_to_delete 
    st.session_state.show_delete_confirm = True
    st.rerun()

def handle_batch_delete_quotes():
    """執行批次刪除操作。"""
    ids_to_delete = st.session_state.get('ids_pending_delete', [])
    
    if not ids_to_delete:
        st.session_state.show_delete_confirm = False
        st.rerun()
        return
    
    current_data = st.session_state.data
    new_data = current_data[~current_data['ID'].isin(ids_to_delete)].reset_index(drop=True)
    
    st.session_state.data = new_data
    
    st.session_state.show_delete_confirm = False
    st.session_state.delete_count = 0
    st.session_state.ids_pending_delete = []
    
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已永久刪除 {len(ids_to_delete)} 筆資料。")


# ==============================================================================
# 應用程式進入點與運行邏輯
# ==============================================================================

def run_app():
    """應用程式的主運行邏輯。"""
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    
    if 'data' not in st.session_state:
        d, m = load_data_from_sheets()
        st.session_state.data = d
        st.session_state.project_metadata = m
    
    if 'next_id' not in st.session_state:
        st.session_state.next_id = st.session_state.data['ID'].max() + 1 if not st.session_state.data.empty and pd.notna(st.session_state.data['ID'].max()) else 1
    if 'edited_dataframes' not in st.session_state: st.session_state.edited_dataframes = {}

    st.session_state.data = calculate_latest_arrival(st.session_state.data, st.session_state.project_metadata)
    
    def get_status_icon(row):
        try:
            proj_date = pd.to_datetime(row['預計交貨日']).date()
            latest_date = pd.to_datetime(row['採購最慢到貨日']).date()
            if proj_date > latest_date: return "🔴 落後" 
            elif proj_date <= latest_date: return "✅ 正常"
            else: return "N/A"
        except: return "N/A"
    
    if not st.session_state.data.empty:
        st.session_state.data['交期狀態'] = st.session_state.data.apply(get_status_icon, axis=1)

    df = st.session_state.data
    
    # ==========================
    #      側邊欄 (Sidebar UI)
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
                
                current_meta = st.session_state.project_metadata.get(target_proj, {'due_date': datetime.now().date()})
                
                if operation == "修改專案資訊":
                    st.text_input("新專案名稱", value=target_proj, key="edit_new_name")
                    st.date_input("新專案交貨日", value=current_meta['due_date'], key="edit_new_date")
                    if st.button("確認修改專案", type="primary", use_container_width=True): 
                        handle_project_modification()
                elif operation == "刪除專案":
                    st.warning(f"⚠️ 確認永久刪除專案 [{target_proj}]？")
                    if st.button("🔥 確認永久刪除", type="secondary", use_container_width=True): 
                        handle_delete_project(target_proj)
            else: 
                st.info("目前無專案資料。請在下方新增。")
        
        st.markdown("---")
        
        # 2. 新增專案
        with st.expander("➕ 新增/設定專案時程", expanded=False):
            st.text_input("專案名稱", key="new_proj_name", placeholder="例如: 辦公室升級")
            project_due_date = st.date_input("專案交貨日", value=datetime.now().date() + timedelta(days=30), key="new_proj_due_date")
            buffer_days = st.number_input("採購緩衝天數", min_value=0, value=7, key="new_proj_buffer_days")
            
            calc_date = project_due_date - timedelta(days=int(buffer_days))
            st.info(f"📅 計算之最慢到貨日：{calc_date.strftime(DATE_FORMAT)}")
            
            if st.button("💾 儲存專案設定", key="btn_save_proj", use_container_width=True): 
                handle_add_new_project()

        st.markdown("---")
        
        # 3. 新增報價
        with st.expander("➕ 新增報價", expanded=True):
            all_projects_for_quote = sorted(list(st.session_state.project_metadata.keys()))
            latest_arrival_date = datetime.now().date()
            
            if not all_projects_for_quote:
                st.warning("請先在上方新增專案。")
                project_name = None
            else:
                st.session_state.quote_project_select = st.selectbox("歸屬專案", all_projects_for_quote, key="quote_project_select_sb")
                current_meta = st.session_state.project_metadata.get(st.session_state.quote_project_select, {'due_date': datetime.now().date(), 'buffer_days': 7})
                latest_arrival_date = current_meta['due_date'] - timedelta(days=int(current_meta['buffer_days']))
                st.caption(f"此專案最慢到貨期限: {latest_arrival_date.strftime(DATE_FORMAT)}")

            unique_items = sorted([x for x in df['專案項目'].unique() if x])
            selected_item = st.selectbox("採購項目", ['🆕 新項目'] + unique_items, key="quote_item_select")
            
            if selected_item == '🆕 新項目':
                item_name_to_use = st.text_input("輸入新項目名稱", key="quote_item_new_input")
            else:
                item_name_to_use = selected_item
            st.session_state.item_name_to_use_final = item_name_to_use
            
            col_sup, col_pr = st.columns(2)
            st.session_state.quote_supplier = col_sup.text_input("供應商", key="quote_supplier_input")
            st.session_state.quote_price = col_pr.number_input("單價", min_value=0.0, key="quote_price_input")
            st.session_state.quote_qty = st.number_input("數量", min_value=1, value=1, key="quote_qty_input")
            
            st.markdown("---")
            st.markdown("📆 **預計交貨日設定**", unsafe_allow_html=True)
            date_input_type = st.radio("輸入方式", ("1. 指定日期", "2. 自然日數", "3. 工作日數"), key="quote_date_type", horizontal=True)
            
            today = datetime.now().date()
            if date_input_type == "1. 指定日期": 
                st.session_state.quote_delivery_date = st.date_input("選擇日期", today, key="quote_delivery_date_input") 
            elif date_input_type == "2. 自然日數": 
                num_days = st.number_input("幾天後交貨?", 1, value=7, key="quote_num_days_input")
                st.session_state.calculated_delivery_date = today + timedelta(days=int(num_days))
                st.info(f"計算結果：{st.session_state.calculated_delivery_date.strftime(DATE_FORMAT)}")
            elif date_input_type == "3. 工作日數": 
                num_b_days = st.number_input("幾個工作天?", 1, value=5, key="quote_num_b_days_input")
                st.session_state.calculated_delivery_date = add_business_days(today, int(num_b_days))
                st.info(f"計算結果：{st.session_state.calculated_delivery_date.strftime(DATE_FORMAT)}")
            
            st.session_state.quote_status = st.selectbox("初始狀態", STATUS_OPTIONS, key="quote_status_select")
            
            st.markdown("📎 **附件上傳**", unsafe_allow_html=True)
            uploaded_file = st.file_uploader("支援 PDF/圖片", type=['pdf', 'jpg', 'jpeg', 'png'], key="new_quote_file_uploader")

            if st.button("📥 新增資料", key="btn_add_quote", type="primary", use_container_width=True):
                handle_add_new_quote(latest_arrival_date, uploaded_file)


    # ==========================
    #      主畫面 (Main UI)
    # ==========================
    
    n, b, r, p = calculate_metrics(df, st.session_state.project_metadata)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("專案數", n)
    c2.metric("總預算", f"${b:,.0f}")
    c3.metric("風險項", r)
    c4.metric("待處理", p)
    
    st.markdown("---")
    
    col_save, col_delete = st.columns([0.8, 0.2])
    with col_save:
        if st.button("💾 儲存所有表格修改 (並重新計算總價)", type="primary"): 
            handle_master_save()
    with col_delete:
        if st.button("🗑️ 刪除已標記項目", type="secondary"):
            trigger_delete_confirmation()

    if st.session_state.get('show_delete_confirm'):
        st.error(f"⚠️ 確認永久刪除 {st.session_state.delete_count} 筆資料？此操作不可復原！")
        cy, cn, _ = st.columns([0.15, 0.15, 0.7])
        with cy: 
            if st.button("✅ 確認刪除", key="confirm_delete_yes", type="primary"): handle_batch_delete_quotes()
        with cn: 
            if st.button("❌ 取消", key="confirm_delete_no"): st.session_state.show_delete_confirm = False
    
    st.markdown("---")

    if df.empty:
        st.info("👋 歡迎使用！目前沒有資料，請從左側側邊欄新增專案與報價。")
        
    for proj_name, proj_data in df.groupby('專案名稱'):
        meta = st.session_state.project_metadata.get(proj_name, {})
        
        st.subheader(f"💼 {proj_name} | 交期: {meta.get('due_date')}")
        
        with st.expander("展開明細", expanded=True):
            for item_name, item_data in proj_data.groupby('專案項目'):
                st.markdown(f"**📦 {item_name}**")
                
                display = item_data.copy()
                display['附件連結'] = None
                for idx, row in display.iterrows():
                    if row.get('附件URL'):
                        url = generate_signed_url_cached(row['附件URL'])
                        if url: display.at[idx, '附件連結'] = url
                
                # Column Config - 回復最原始設定，取消 DateColumn 以解決錯誤
                edited = st.data_editor(
                    display[['ID', '選取', '供應商', '單價', '數量', '總價', 
                             '預計交貨日', '交期狀態', '狀態', '附件連結', '最後修改時間', '標記刪除', '附件URL']],
                    column_config={
                        "ID": st.column_config.NumberColumn("ID", disabled=True, width="small"),
                        "選取": st.column_config.CheckboxColumn("選", width="small"),
                        "總價": st.column_config.NumberColumn(format="$%d", disabled=True),
                        # 取消 DateColumn 設定，避免字串/日期型別衝突
                        "預計交貨日": st.column_config.TextColumn("預計交貨日", help="格式: YYYY-MM-DD"),
                        
                        "交期狀態": st.column_config.TextColumn("期限判定", disabled=True, width="small"), 
                        "最後修改時間": st.column_config.TextColumn("最後修改", disabled=True, width="medium"), 
                        "附件連結": st.column_config.LinkColumn("附件", display_text="📄 開啟", width="small"),
                        "標記刪除": st.column_config.CheckboxColumn("刪除?", width="small"),
                        # 恢復顯示系統路徑 (V2.2.5 預設顯示)
                        "附件URL": st.column_config.TextColumn("系統路徑 (gs://)", disabled=True) 
                    },
                    hide_index=True,
                    key=f"ed_{proj_name}_{item_name}",
                    num_rows="dynamic"
                )
                st.session_state.edited_dataframes[item_name] = edited
            st.markdown("<br>", unsafe_allow_html=True)

def main():
    login_form()
    if st.session_state.authenticated: run_app()

if __name__ == "__main__":
    main()
