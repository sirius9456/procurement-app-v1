import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os 
import json
import gspread
import logging

# 引入 Google Cloud Storage 庫
from google.cloud import storage

# 確保 openpyxl 庫已安裝 (pip install openpyxl)

# 配置 Streamlit 日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# --- 應用程式設定 ---
APP_VERSION = "v2.2.4 (Full Features + Hyperlink Fix)"
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"]

# --- Google Cloud Storage 配置 ---
GCS_BUCKET_NAME = "procurement-attachments-bucket"
GCS_ATTACHMENT_FOLDER = "attachments"

# --- 數據源配置 ---
# 將憑證路徑設為全域變數，供 Gspread 和 GCS 共用
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
    try:
        GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"] 
    except KeyError:
        logging.error("GSHEETS_CREDENTIALS_PATH is missing.")
        st.error("❌ 錯誤：找不到 GSHEETS_CREDENTIALS_PATH 環境變數。")
        GSHEETS_CREDENTIALS = None 
else:
    try:
        SHEET_URL = st.secrets["app_config"]["sheet_url"]
        GSHEETS_CREDENTIALS = None
    except KeyError:
        SHEET_URL = None
        GSHEETS_CREDENTIALS = None
        
DATA_SHEET_NAME = "採購總表"
METADATA_SHEET_NAME = "專案設定"


st.set_page_config(page_title=f"專案採購小幫手 {APP_VERSION}", layout="wide")

# --- CSS 樣式 ---
CUSTOM_CSS = """
<style>
.streamlit-expanderContent { padding-left: 1rem !important; padding-right: 1rem !important; padding-bottom: 1rem !important; }
.project-header { font-size: 20px !important; font-weight: bold !important; color: #FAFAFA; }
.item-header { font-size: 16px !important; font-weight: 600 !important; color: #E0E0E0; }
.meta-info { font-size: 14px !important; color: #9E9E9E; font-weight: normal; }
div[data-baseweb="select"] > div, div[data-baseweb="base-input"] > input, div[data-baseweb="input"] > div { background-color: #262730 !important; color: white !important; -webkit-text-fill-color: white !important; }
div[data-baseweb="popover"], div[data-baseweb="menu"] { background-color: #262730 !important; }
div[data-baseweb="option"] { color: white !important; }
li[aria-selected="true"] { background-color: #FF4B4B !important; color: white !important; }
.metric-box { padding: 10px 15px; border-radius: 8px; margin-bottom: 10px; background-color: #262730; text-align: center; }
.metric-title { font-size: 14px; color: #9E9E9E; margin-bottom: 5px; }
.metric-value { font-size: 24px; font-weight: bold; }
</style>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
"""

# --- 登入與安全函式 ---

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


# --- GCS 檔案服務函式 (V2.2.4 修復版) ---

def get_storage_client():
    """獲取 GCS 客戶端，優先使用 JSON 金鑰以支援簽署功能。"""
    if GSHEETS_CREDENTIALS and os.path.exists(GSHEETS_CREDENTIALS):
        # 關鍵修復：明確使用 Service Account JSON，確保有 Private Key 進行簽署
        return storage.Client.from_service_account_json(GSHEETS_CREDENTIALS)
    else:
        return storage.Client()

def upload_attachment_to_gcs(file_obj, next_id):
    """將檔案上傳到 GCS。"""
    try:
        storage_client = get_storage_client()
        bucket = storage_client.bucket(GCS_BUCKET_NAME)
        file_extension = os.path.splitext(file_obj.name)[1]
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        blob_name = f"{GCS_ATTACHMENT_FOLDER}/{next_id}_{timestamp}{file_extension}"
        
        blob = bucket.blob(blob_name)
        file_obj.seek(0)
        blob.upload_from_file(file_obj, content_type=file_obj.type)
        return f"gs://{GCS_BUCKET_NAME}/{blob_name}"

    except Exception as e:
        logging.error(f"GCS上傳失敗: {e}")
        st.error("❌ 附件上傳失敗，請檢查 GCS 權限。")
        return None

def generate_signed_url_cached(gcs_uri):
    """生成簽章 URL (有效期 1 小時)。"""
    if not gcs_uri or not isinstance(gcs_uri, str):
        return None
    if gcs_uri.startswith("http"):
        return gcs_uri
    if not gcs_uri.startswith("gs://"):
        return None

    try:
        parts = gcs_uri[5:].split('/', 1)
        bucket_name = parts[0]
        blob_name = parts[1]
        
        storage_client = get_storage_client() # 使用帶私鑰的客戶端
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(minutes=60),
            method="GET"
        )
        return url
    except Exception as e:
        logging.error(f"生成 Signed URL 失敗: {e}")
        return None


# --- 數據讀取與寫入函式 (Gspread) ---

@st.cache_data(ttl=600, show_spinner="連線 Google Sheets...")
def load_data_from_sheets():
    if not SHEET_URL:
        st.info("❌ Google Sheets URL 尚未配置。")
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除', '附件URL'])
        return empty_data, {}

    try:
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # 讀取 Data
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_records = data_ws.get_all_records()
        data_df = pd.DataFrame(data_records)

        if '附件URL' not in data_df.columns:
            data_df['附件URL'] = ""
            
        data_df = data_df.astype({'ID': 'Int64', '選取': 'bool', '單價': 'float', '數量': 'Int64', '總價': 'float'})
        if '標記刪除' not in data_df.columns: data_df['標記刪除'] = False

        # 讀取 Metadata
        metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        metadata_records = metadata_ws.get_all_records()
        
        project_metadata = {}
        if metadata_records:
            for row in metadata_records:
                try: due_date = pd.to_datetime(str(row['專案交貨日'])).date()
                except: due_date = datetime.now().date()
                project_metadata[row['專案名稱']] = {
                    'due_date': due_date,
                    'buffer_days': int(row['緩衝天數']),
                    'last_modified': str(row['最後修改'])
                }

        st.success("✅ 數據已從 Google Sheets 載入！")
        return data_df, project_metadata

    except Exception as e:
        logging.exception("Google Sheets 數據載入失敗") 
        st.error(f"❌ 數據載入失敗！錯誤訊息: {e}")
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除', '附件URL'])
        st.session_state.data_load_failed = True
        return empty_data, {}


def write_data_to_sheets(df_to_write, metadata_to_write):
    if st.session_state.get('data_load_failed', False) or not SHEET_URL:
        st.warning("數據載入失敗，已禁用寫入 Sheets。")
        return False
        
    try:
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # ⚠️ 關鍵：移除輔助顯示欄位，只寫入原始數據
        # 移除 '附件連結' (这是生成的 Signed URL，不應寫入 Sheets)
        # 移除 '交期顯示'
        df_export = df_to_write.drop(columns=['標記刪除', '交期顯示', '附件連結'], errors='ignore')
        
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_ws.clear()
        data_ws.update([df_export.columns.values.tolist()] + df_export.values.tolist())
        
        metadata_list = [
            {'專案名稱': name, 
             '專案交貨日': data['due_date'].strftime('%Y-%m-%d'),
             '緩衝天數': data['buffer_days'], 
             '最後修改': data['last_modified']}
            for name, data in metadata_to_write.items()
        ]
        metadata_df = pd.DataFrame(metadata_list)
        metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        metadata_ws.clear()
        metadata_ws.update([metadata_df.columns.values.tolist()] + metadata_df.values.tolist())
        
        st.cache_data.clear() 
        return True
        
    except Exception as e:
        logging.exception("Google Sheets 寫入失敗")
        st.error(f"❌ 寫入失敗！錯誤訊息: {e}")
        return False


# --- 輔助函式區 ---

def add_business_days(start_date, num_days):
    current_date = start_date
    days_added = 0
    while days_added < num_days:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5: days_added += 1
    return current_date

@st.cache_data
def convert_df_to_excel(df):
    # 移除輔助欄位後匯出
    df_export = df.drop(columns=['標記刪除', '交期顯示', '附件連結'], errors='ignore')
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='採購報價總表')
    processed_data = output.getvalue()
    return processed_data

@st.cache_data(show_spinner=False)
def calculate_dashboard_metrics(df_state, project_metadata_state):
    total_projects = len(project_metadata_state)
    total_budget = 0
    risk_items = 0
    df = df_state.copy()
    
    if df.empty:
        return 0, 0, 0, 0

    for _, proj_data in df.groupby('專案名稱'):
        for _, item_df in proj_data.groupby('專案項目'):
            selected_rows = item_df[item_df['選取'] == True]
            if not selected_rows.empty:
                total_budget += selected_rows['總價'].sum()
            elif not item_df.empty:
                total_budget += item_df['總價'].min()
    
    temp_df_risk = df.copy() 
    temp_df_risk['預計交貨日_dt'] = pd.to_datetime(temp_df_risk['預計交貨日'], errors='coerce')
    temp_df_risk['採購最慢到貨日_dt'] = pd.to_datetime(temp_df_risk['採購最慢到貨日'], errors='coerce')
    risk_items = (temp_df_risk['預計交貨日_dt'] > temp_df_risk['採購最慢到貨日_dt']).sum()

    pending_quotes = df[~df['狀態'].isin(['已收貨', '取消'])].shape[0]

    return total_projects, total_budget, risk_items, pending_quotes

def calculate_project_budget(df, project_name):
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
    if df.empty or not metadata:
        return df

    metadata_df = pd.DataFrame.from_dict(metadata, orient='index')
    metadata_df = metadata_df.reset_index().rename(columns={'index': '專案名稱'})
    metadata_df['due_date'] = metadata_df['due_date'].apply(lambda x: pd.to_datetime(x).date())
    metadata_df['buffer_days'] = metadata_df['buffer_days'].astype(int)

    df = pd.merge(df, metadata_df[['專案名稱', 'due_date', 'buffer_days']], on='專案名稱', how='left')
    
    df['採購最慢到貨日_TEMP'] = (
        pd.to_datetime(df['due_date']) - 
        df['buffer_days'].apply(lambda x: timedelta(days=x))
    )
    df['採購最慢到貨日'] = df['採購最慢到貨日_TEMP'].dt.strftime('%Y-%m-%d')
    df = df.drop(columns=['due_date', 'buffer_days', '採購最慢到貨日_TEMP'], errors='ignore')
    return df


# --- UI 邏輯處理函式 ---

def save_and_rerun(df_to_save, metadata_to_save, success_message=""):
    if write_data_to_sheets(df_to_save, metadata_to_save):
        st.session_state.edited_dataframes = {}
        if success_message:
            st.success(success_message)
        st.rerun()

def handle_master_save():
    """批次處理修改並儲存。"""
    if not st.session_state.edited_dataframes:
        st.info("沒有偵測到表格修改。")
        return

    main_df = st.session_state.data
    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    affected_projects = set()
    changes_detected = False
    
    for _, edited_df in st.session_state.edited_dataframes.items():
        if edited_df.empty: continue
        
        for index, new_row in edited_df.iterrows():
            original_id = new_row['ID']
            idx_in_main = main_df[main_df['ID'] == original_id].index
            if idx_in_main.empty: continue
            main_idx = idx_in_main[0]
            
            # 更新可編輯欄位 (包含 附件URL，雖然通常不會手動改它)
            # ⚠️ 注意：不要從 edited_df 讀取 '附件連結' 回寫到 main_df 的 '附件URL'
            updatable_cols = ['選取', '供應商', '單價', '數量', '狀態', '標記刪除', '附件URL'] 
            for col in updatable_cols:
                if col in new_row and main_df.loc[main_idx, col] != new_row[col]:
                    main_df.loc[main_idx, col] = new_row[col]
                    changes_detected = True
            
            try:
                date_str_parts = str(new_row['交期顯示']).strip().split(' ')
                date_part = date_str_parts[0]
                if main_df.loc[main_idx, '預計交貨日'] != date_part:
                    main_df.loc[main_idx, '預計交貨日'] = date_part
                    changes_detected = True
            except: pass
            
            new_total = float(new_row['單價']) * float(new_row['數量'])
            if main_df.loc[main_idx, '總價'] != new_total:
                main_df.loc[main_idx, '總價'] = new_total
                changes_detected = True
            
            affected_projects.add(main_df.loc[main_idx, '專案名稱'])

    if changes_detected:
        st.session_state.data = main_df.copy()
        for proj in affected_projects:
            if proj in st.session_state.project_metadata:
                st.session_state.project_metadata[proj]['last_modified'] = current_time_str
        
        save_and_rerun(st.session_state.data, st.session_state.project_metadata, "✅ 資料已儲存！Google Sheets 已更新。")
    else:
        st.info("沒有偵測到表格修改。")

def trigger_delete_confirmation():
    """觸發刪除確認流程。"""
    temp_df = st.session_state.data.copy()
    combined_edited_df = pd.concat(
        [edited_df.set_index('ID')[['標記刪除']] for edited_df in st.session_state.edited_dataframes.values() if not edited_df.empty],
        axis=0, ignore_index=False
    )
    if not combined_edited_df.empty:
        temp_df = temp_df.set_index('ID')
        temp_df.update(combined_edited_df)
        temp_df = temp_df.reset_index()

    ids_to_delete = temp_df[temp_df['標記刪除'] == True]['ID'].tolist()
    if not ids_to_delete:
        st.warning("沒有項目被標記為刪除。")
        st.session_state.show_delete_confirm = False
        return

    st.session_state.delete_count = len(ids_to_delete)
    st.session_state.show_delete_confirm = True
    st.rerun()

def handle_batch_delete_quotes():
    """執行批次刪除。"""
    ids_to_delete = []
    # 再次確認要刪除的 ID (從 session state 或 edited data)
    # 簡單起見，直接從 data 篩選 (假設已經在 trigger 階段 update 暫存，或者直接掃描 edited)
    # 更好的做法是直接操作 st.session_state.data，因為在 trigger 前通常會先 save，或者這裡再合併一次
    
    # 為確保準確，我們先執行一次類似 save 的合併 (但不寫入 sheets) 到 local variable
    current_data = st.session_state.data.copy()
    for _, edited_df in st.session_state.edited_dataframes.items():
        if not edited_df.empty:
            for _, row in edited_df.iterrows():
                if row.get('標記刪除') == True:
                    current_data.loc[current_data['ID'] == row['ID'], '標記刪除'] = True
    
    ids_to_delete = current_data[current_data['標記刪除'] == True]['ID'].tolist()
    
    if ids_to_delete:
        st.session_state.data = current_data[~current_data['ID'].isin(ids_to_delete)].reset_index(drop=True)
        st.session_state.show_delete_confirm = False
        st.session_state.delete_count = 0
        save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已刪除 {len(ids_to_delete)} 筆資料。")
    else:
        st.session_state.show_delete_confirm = False
        st.rerun()

def cancel_delete_confirmation():
    st.session_state.show_delete_confirm = False
    st.rerun()

def handle_project_modification():
    target_proj = st.session_state.edit_target_project
    new_name = st.session_state.edit_new_name
    new_date = st.session_state.edit_new_date
    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    if not new_name:
        st.error("專案名稱不能為空")
        return
    if target_proj != new_name and new_name in st.session_state.project_metadata:
        st.error(f"新的專案名稱 '{new_name}' 已存在。")
        return

    meta = st.session_state.project_metadata.pop(target_proj)
    meta['due_date'] = new_date
    meta['last_modified'] = current_time_str
    st.session_state.project_metadata[new_name] = meta
    
    st.session_state.data.loc[st.session_state.data['專案名稱'] == target_proj, '專案名稱'] = new_name
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 專案已更新：{new_name}。")

def handle_delete_project(project_to_delete):
    if not project_to_delete: return
    if project_to_delete in st.session_state.project_metadata:
        del st.session_state.project_metadata[project_to_delete]
    
    st.session_state.data = st.session_state.data[st.session_state.data['專案名稱'] != project_to_delete].reset_index(drop=True)
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 專案 {project_to_delete} 已刪除。")

def handle_add_new_project():
    project_name = st.session_state.new_proj_name
    project_due_date = st.session_state.new_proj_due_date
    buffer_days = st.session_state.new_proj_buffer_days
    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    if not project_name:
        st.error("專案名稱不能為空。")
        return
        
    st.session_state.project_metadata[project_name] = {
        'due_date': project_due_date, 
        'buffer_days': buffer_days,
        'last_modified': current_time_str
    }
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已儲存專案設定：{project_name}。")

def handle_add_new_quote(latest_arrival_date, uploaded_file):
    project_name = st.session_state.quote_project_select
    item_name_to_use = st.session_state.item_name_to_use_final
    supplier = st.session_state.quote_supplier
    price = st.session_state.quote_price
    qty = st.session_state.quote_qty
    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    if st.session_state.quote_date_type == "1. 指定日期":
        final_delivery_date = st.session_state.quote_delivery_date
    else:
        final_delivery_date = st.session_state.calculated_delivery_date 

    if not project_name or not item_name_to_use:
        st.error("請確認已輸入專案名稱並選擇項目。")
        return

    total_price = price * qty
    
    # GCS 上傳
    attachment_uri = ""
    next_id = st.session_state.next_id
    if uploaded_file is not None:
        with st.spinner(f"正在上傳附件 {uploaded_file.name}..."):
            attachment_uri = upload_attachment_to_gcs(uploaded_file, next_id)
            if attachment_uri is None: return 

    st.session_state.project_metadata[project_name]['last_modified'] = current_time_str

    new_row = {
        'ID': st.session_state.next_id, '選取': False, '專案名稱': project_name, 
        '專案項目': item_name_to_use, '供應商': supplier, '單價': price, '數量': qty, 
        '總價': total_price, '預計交貨日': final_delivery_date.strftime('%Y-%m-%d'), 
        '狀態': st.session_state.quote_status, '採購最慢到貨日': latest_arrival_date.strftime('%Y-%m-%d'), 
        '標記刪除': False,
        '附件URL': attachment_uri 
    }
    st.session_state.next_id += 1
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, f"✅ 已新增報價至 {project_name}！")


# --- 初始化 Session State ---
def initialize_session_state():
    today = datetime.now().date()
    if 'data' not in st.session_state:
        data_df, metadata_dict = load_data_from_sheets()
        st.session_state.data = data_df
        st.session_state.project_metadata = metadata_dict
        
    next_id_val = st.session_state.data['ID'].max() + 1 if not st.session_state.data.empty else 1
    
    initial_values = {
        'next_id': next_id_val,
        'edited_dataframes': {},
        'calculated_delivery_date': today,
        'show_delete_confirm': False,
        'delete_count': 0,
    }
    for key, value in initial_values.items():
        st.session_state.setdefault(key, value)
        
    if '標記刪除' not in st.session_state.data.columns: st.session_state.data['標記刪除'] = False
    if '附件URL' not in st.session_state.data.columns: st.session_state.data['附件URL'] = ""


# --- 主程式 ---
def run_app():
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}") 
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    initialize_session_state()
    
    st.session_state.data = calculate_latest_arrival_dates(st.session_state.data, st.session_state.project_metadata)
    
    if st.session_state.get('data_load_failed', False):
        st.warning("應用程式無法從 Google Sheets 載入數據。")
        
    today = datetime.now().date() 

    # --- 日期格式化 (供顯示用) ---
    def format_date_with_icon(row):
        date_str = str(row['預計交貨日'])
        try:
            v_date = pd.to_datetime(row['預計交貨日']).date()
            l_date = pd.to_datetime(row['採購最慢到貨日']).date()
            return f"{date_str} 🔴" if v_date > l_date else f"{date_str} ✅"
        except: return date_str

    if not st.session_state.data.empty:
        st.session_state.data['交期顯示'] = st.session_state.data.apply(format_date_with_icon, axis=1)

    df = st.session_state.data
    project_groups = df.groupby('專案名稱')
    
    # --- 側邊欄 ---
    with st.sidebar:
        st.button("登出", on_click=logout, type="secondary")
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
                    if st.button("確認修改", type="primary"): handle_project_modification()
                elif operation == "刪除專案":
                    st.warning(f"確認刪除專案 {target_proj}？此操作不可逆。")
                    if st.button("確認永久刪除", type="secondary"): handle_delete_project(target_proj)
            else: 
                st.info("無專案。")
        
        st.markdown("---")
        
        # 2. 新增專案
        with st.expander("➕ 新增/設定專案時程", expanded=False):
            st.text_input("專案名稱", key="new_proj_name")
            project_due_date = st.date_input("專案交貨日", value=today + timedelta(days=30), key="new_proj_due_date")
            buffer_days = st.number_input("緩衝天數", min_value=0, value=7, key="new_proj_buffer_days")
            st.caption(f"最慢到貨日：{(project_due_date - timedelta(days=int(buffer_days))).strftime('%Y-%m-%d')}")
            if st.button("儲存設定", key="btn_save_proj"): handle_add_new_project()
        
        st.markdown("---")
        
        # 3. 新增報價 (GCS版)
        with st.expander("➕ 新增報價", expanded=True):
            all_projects_for_quote = sorted(list(st.session_state.project_metadata.keys()))
            latest_arrival_date = today 
            
            if not all_projects_for_quote:
                st.warning("請先設定專案。")
                project_name = None
            else:
                project_name = st.selectbox("選擇專案", all_projects_for_quote, key="quote_project_select")
                current_meta = st.session_state.project_metadata.get(project_name, {'due_date': today, 'buffer_days': 7})
                latest_arrival_date = current_meta['due_date'] - timedelta(days=int(current_meta['buffer_days']))
                st.caption(f"最慢到貨: {latest_arrival_date.strftime('%Y-%m-%d')}")

            unique_items = sorted(st.session_state.data['專案項目'].unique().tolist())
            selected_item = st.selectbox("項目", ['新增項目...'] + unique_items, key="quote_item_select")
            item_name_to_use = st.text_input("新項目名稱", key="quote_item_new_input") if selected_item == '新增項目...' else selected_item
            st.session_state.item_name_to_use_final = item_name_to_use
            
            st.text_input("供應商", key="quote_supplier")
            st.number_input("單價", min_value=0, key="quote_price")
            st.number_input("數量", min_value=1, value=1, key="quote_qty")
            
            date_input_type = st.radio("交期方式", ("1. 指定日期", "2. 自然日數", "3. 工作日數"), key="quote_date_type", horizontal=True)
            if date_input_type == "1. 指定日期": 
                st.date_input("交貨日期", today, key="quote_delivery_date") 
            elif date_input_type == "2. 自然日數": 
                num_days = st.number_input("自然日數", 1, value=7, key="quote_num_days_input")
                st.session_state.calculated_delivery_date = today + timedelta(days=int(num_days))
            elif date_input_type == "3. 工作日數": 
                num_b_days = st.number_input("工作日數", 1, value=5, key="quote_num_b_days_input")
                st.session_state.calculated_delivery_date = add_business_days(today, int(num_b_days))
            
            if date_input_type != "1. 指定日期":
                st.caption(f"交期：{st.session_state.calculated_delivery_date.strftime('%Y-%m-%d')}")

            st.selectbox("狀態", STATUS_OPTIONS, key="quote_status")
            uploaded_file = st.file_uploader("附件 (PDF/圖片)", type=['pdf', 'jpg', 'jpeg', 'png'], key="new_quote_file_uploader")

            if st.button("新增資料", key="btn_add_quote"):
                handle_add_new_quote(latest_arrival_date, uploaded_file)


    # --- 儀表板 ---
    total_projects, total_budget, risk_items, pending_quotes = calculate_dashboard_metrics(df, st.session_state.project_metadata)

    st.subheader("📊 總覽儀表板")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='metric-box'><div class='metric-title'>專案數</div><div class='metric-value'>{total_projects}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='metric-box' style='background:#21442c'><div class='metric-title'>總預算</div><div class='metric-value'>${total_budget:,.0f}</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='metric-box' style='background:#5c2d2d'><div class='metric-title'>風險項</div><div class='metric-value'>{risk_items}</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='metric-box' style='background:#2a3b5c'><div class='metric-title'>待處理</div><div class='metric-value'>{pending_quotes}</div></div>", unsafe_allow_html=True)
    
    st.markdown("---")
    
    # --- 批次操作 ---
    col_save, col_delete = st.columns([0.8, 0.2])
    is_locked = st.session_state.show_delete_confirm
    with col_save:
        if st.button("💾 儲存表格修改並計算總價", type="primary", disabled=is_locked):
            handle_master_save()
    with col_delete:
        if st.button("🔴 刪除已標記項目", type="secondary", disabled=is_locked, key="btn_trigger_delete"):
            trigger_delete_confirmation()

    if st.session_state.show_delete_confirm:
        st.error(f"⚠️ 確認永久刪除 {st.session_state.delete_count} 筆資料？")
        cy, cn, _ = st.columns([0.2, 0.2, 0.6])
        with cy: 
            if st.button("✅ 確認刪除", key="confirm_delete_yes", type="primary"): handle_batch_delete_quotes()
        with cn: 
            if st.button("❌ 取消", key="confirm_delete_no"): cancel_delete_confirmation()

    st.markdown("---")

    # --- 專案列表 (整合超連結) ---
    for proj_name, proj_data in project_groups:
        meta = st.session_state.project_metadata.get(proj_name, {})
        proj_budget = calculate_project_budget(df, proj_name)
        header_html = f"""
        <span class='project-header'>💼 {proj_name}</span> &nbsp;|&nbsp; 
        <span class='project-header'>總預算: ${proj_budget:,.0f}</span> &nbsp;|&nbsp; 
        <span class='meta-info'>交期: {meta.get('due_date')}</span> 
        """
        
        with st.expander(label=f"專案：{proj_name}", expanded=False):
            st.markdown(header_html, unsafe_allow_html=True)
            
            for item_name, item_data in proj_data.groupby('專案項目'):
                st.markdown(f"<span class='item-header'>📦 {item_name}</span>", unsafe_allow_html=True)
                
                # --- 建立顯示用 DataFrame (生成超連結) ---
                display_df = item_data.copy()
                display_df['附件連結'] = None
                
                # 預先生成所有 Signed URL
                for idx, row in display_df.iterrows():
                    uri = row.get('附件URL', '')
                    if uri:
                        signed_url = generate_signed_url_cached(uri)
                        if signed_url:
                            display_df.at[idx, '附件連結'] = signed_url

                editor_key = f"ed_{proj_name}_{item_name}"
                
                # 版面配置：隱藏原始 GS路徑，只顯示可點擊連結
                # 將 '附件URL' 移到最後並設為 disabled (或寬度極小)，主要顯示 '附件連結'
                column_order = [
                    'ID', '選取', '供應商', '單價', '數量', '總價', 
                    '交期顯示', '狀態', '附件連結', '標記刪除', 
                    '附件URL' # 放在最後
                ]

                edited_df_value = st.data_editor(
                    display_df[column_order],
                    column_config={
                        "ID": st.column_config.Column(disabled=True, width="small"),
                        "附件連結": st.column_config.LinkColumn(
                            "附件 (點擊開啟)", 
                            display_text="📄 開啟附件", 
                            help="點擊在新視窗開啟檔案",
                            width="medium"
                        ),
                        "附件URL": st.column_config.TextColumn(
                            "系統路徑", 
                            disabled=True, 
                            width="small",
                            help="原始 gs:// 路徑"
                        ),
                        "交期顯示": st.column_config.TextColumn("交貨日", disabled=False),
                        "總價": st.column_config.NumberColumn(disabled=True),
                        "選取": st.column_config.CheckboxColumn("選", width="small"),
                        "標記刪除": st.column_config.CheckboxColumn("刪?", width="small"),
                    },
                    hide_index=True,
                    key=editor_key,
                    disabled=is_locked
                )
                st.session_state.edited_dataframes[item_name] = edited_df_value 
                st.markdown("---")

    # --- 匯出 ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("💾 資料匯出")
    st.download_button("📥 下載 Excel", convert_df_to_excel(df), f'report_{datetime.now().strftime("%Y%m%d")}.xlsx')


def main():
    login_form()
    if st.session_state.authenticated:
        run_app() 
        
if __name__ == "__main__":
    main()
