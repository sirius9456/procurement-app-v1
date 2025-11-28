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
APP_VERSION = "v2.2.1 (Security Update - Signed URL)" # 版本更新為安全更新測試版
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"]

# --- Google Cloud Storage 配置 (請務必替換為您的儲存桶名稱) ---
# ⚠️ WARNING: 請替換為您在 GCP 上建立的儲存桶名稱！
GCS_BUCKET_NAME = "procurement-attachments-bucket" 
GCS_ATTACHMENT_FOLDER = "attachments"

# --- 數據源配置 (安全與 Gspread 連線) ---
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
    try:
        # GCE 服務帳戶自動獲得 GCS 存取權限 (若角色正確)
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


st.set_page_config(page_title=f"專案採購小幫手 {APP_VERSION}", layout="wide")

# --- CSS 樣式修正 (保持不變) ---
CUSTOM_CSS = """
<style>
/* 保持原樣 */
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

/* Modal 樣式，確保背景色與 Streamlit 主題一致 */
/* Streamlit Modal API doesn't allow custom styling, but we keep this for reference */
</style>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
"""

# --- 登入與安全函式 (保持不變) ---

def logout():
    """登出函式：清除驗證狀態並重新運行。"""
    st.session_state["authenticated"] = False
    st.rerun()

def login_form():
    """渲染登入表單並處理密碼驗證。"""
    
    # 從 systemd 環境變數中讀取密碼 (安全關鍵!)
    DEFAULT_USERNAME = os.environ.get("AUTH_USERNAME", "dev_user")
    DEFAULT_PASSWORD = os.environ.get("AUTH_PASSWORD", "dev_pwd")
    
    credentials = {"username": DEFAULT_USERNAME, "password": DEFAULT_PASSWORD}

    # 初始化驗證狀態
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
        
    if st.session_state["authenticated"]:
        return # 如果已驗證，則跳過登入表單

    # 渲染登入介面
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_empty, col_center, col_empty2 = st.columns([1, 2, 1])
    
    with col_center:
        with st.container(border=True):
            st.title("🔐 請登入以繼續")
            st.markdown("---")
            
            # 用戶名輸入框預設為環境變數的值，禁用更改
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


# --- GCS 檔案服務函式 (安全更新 V2.2.1) ---

def upload_attachment_to_gcs(file_obj, next_id):
    """將檔案上傳到 GCS，不設置公開權限 (儲存桶保持私有)。"""
    if GCS_BUCKET_NAME == "procurement-attachments-bucket":
        st.warning("GCS 儲存桶名稱未設置。請修改 GCS_BUCKET_NAME 變數。")
        return None
        
    try:
        # GCE 服務帳戶自動認證
        storage_client = storage.Client()
        bucket = storage_client.bucket(GCS_BUCKET_NAME)
        
        # 建立 GCS 上的檔案路徑: attachments/{next_id}-{timestamp}-{filename_ext}
        file_extension = os.path.splitext(file_obj.name)[1]
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        blob_name = f"{GCS_ATTACHMENT_FOLDER}/{next_id}_{timestamp}{file_extension}"
        
        blob = bucket.blob(blob_name)
        
        # 上傳檔案
        file_obj.seek(0) # 確保從檔案開頭讀取
        blob.upload_from_file(file_obj, content_type=file_obj.type)
        
        # ⚠️ CRITICAL: 移除 blob.make_public()，確保儲存桶是私有的。
        
        # 返回檔案的 GCS 存儲路徑 (gs://bucket/blob_name)，以便後續生成 Signed URL
        return f"gs://{GCS_BUCKET_NAME}/{blob_name}"

    except Exception as e:
        logging.error(f"GCS 上傳失敗，請檢查 GCE 服務帳戶是否有 Storage Object Creator 權限: {e}")
        st.error("❌ 附件上傳失敗，請檢查 GCS 權限。")
        return None

def get_signed_attachment_url(gcs_uri):
    """根據 GCS URI 生成一個有時效限制 (5 分鐘) 的簽章 URL。"""
    if not gcs_uri.startswith("gs://"):
        return gcs_uri # 如果已經是普通的 URL，直接返回
    
    try:
        storage_client = storage.Client()
        # 解析 URI 獲取 bucket 和 blob 名稱
        parts = gcs_uri[5:].split('/', 1)
        bucket_name = parts[0]
        blob_name = parts[1]
        
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        
        # 生成簽章 URL，時效 5 分鐘
        signed_url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(minutes=5),
            method="GET"
        )
        return signed_url
        
    except Exception as e:
        logging.error(f"生成 Signed URL 失敗: {e}")
        return None


# --- 數據讀取與寫入函式 (Gspread) ---

@st.cache_data(ttl=600, show_spinner="連線 Google Sheets...")
def load_data_from_sheets():
    """直接使用 gspread 讀取 Google Sheets 中的數據。"""
    
    if not SHEET_URL:
        st.info("❌ Google Sheets URL 尚未配置。使用空的數據結構。")
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除', '附件URL'])
        return empty_data, {}

    try:
        # --- 1. 授權與認證 ---
        if not GSHEETS_CREDENTIALS or not os.path.exists(GSHEETS_CREDENTIALS):
             st.error(f"❌ 憑證錯誤：找不到憑證檔案 {GSHEETS_CREDENTIALS}")
             raise FileNotFoundError("憑證檔案不存在或路徑錯誤")
             
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 2. 讀取採購總表 (Data) ---
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_records = data_ws.get_all_records()
        data_df = pd.DataFrame(data_records)

        # 確保 '附件URL' 欄位存在
        if '附件URL' not in data_df.columns:
            data_df['附件URL'] = ""
            
        # 數據類型轉換與處理
        data_df = data_df.astype({
            'ID': 'Int64', '選取': 'bool', '單價': 'float', '數量': 'Int64', '總價': 'float'
        })
        if '標記刪除' not in data_df.columns:
            data_df['標記刪除'] = False

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
                    'buffer_days': int(row['緩衝天數']),
                    'last_modified': str(row['最後修改'])
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
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 2. 寫入採購總表 (Data) ---
        # 確保 '交期顯示' 不寫入 Sheets，但 '附件URL' 要保留
        df_export = df_to_write.drop(columns=['標記刪除', '交期顯示'], errors='ignore')
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_ws.clear()
        data_ws.update([df_export.columns.values.tolist()] + df_export.values.tolist())
        
        # ... (省略 metadata 寫入，保持不變) ...
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
        
        # 效能優化：成功寫入後，清除 Streamlit 快取
        st.cache_data.clear() 
        return True
        
    except Exception as e:
        logging.exception("Google Sheets 數據寫入時發生致命錯誤！")
        st.error(f"❌ 數據寫回 Google Sheets 失敗！")
        st.code(f"寫入錯誤訊息: {e}")
        return False


# --- 輔助函式區 (保持不變) ---

def add_business_days(start_date, num_days):
    current_date = start_date
    days_added = 0
    while days_added < num_days:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5: days_added += 1
    return current_date

@st.cache_data
def convert_df_to_excel(df):
    """將 DataFrame 轉換為 Excel 二進位檔案 (使用 BytesIO)。"""
    df_export = df.drop(columns=['標記刪除', '交期顯示'], errors='ignore')
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='採購報價總表')
    
    processed_data = output.getvalue()
    return processed_data


# ... (省略 calculate_dashboard_metrics, calculate_project_budget, calculate_latest_arrival_dates) ...

@st.cache_data(show_spinner=False)
def calculate_dashboard_metrics(df_state, project_metadata_state):
    """計算儀表板所需的總體指標。此函式會被緩存。"""
    
    total_projects = len(project_metadata_state)
    total_budget = 0
    risk_items = 0
    df = df_state.copy()
    
    if df.empty:
        return 0, 0, 0, 0

    # 1. 計算總預算
    for _, proj_data in df.groupby('專案名稱'):
        for _, item_df in proj_data.groupby('專案項目'):
            selected_rows = item_df[item_df['選取'] == True]
            if not selected_rows.empty:
                total_budget += selected_rows['總價'].sum()
            elif not item_df.empty:
                total_budget += item_df['總價'].min()
    
    # 2. 計算風險項目
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


# 專案交期自動計算邏輯 (V2.1.8 優化精簡)
@st.cache_data(show_spinner=False)
def calculate_latest_arrival_dates(df, metadata):
    """根據專案設定，計算每個採購項目的採購最慢到貨日。"""
    
    if df.empty or not metadata:
        return df

    metadata_df = pd.DataFrame.from_dict(metadata, orient='index')
    metadata_df = metadata_df.reset_index().rename(columns={'index': '專案名稱'})
    
    # 保持 due_date 的類型處理
    metadata_df['due_date'] = metadata_df['due_date'].apply(lambda x: pd.to_datetime(x).date())
    metadata_df['buffer_days'] = metadata_df['buffer_days'].astype(int)

    df = pd.merge(df, metadata_df[['專案名稱', 'due_date', 'buffer_days']], on='專案名稱', how='left')
    
    # 【程式碼精簡】: 直接轉換並執行減法運算，避免中間欄位
    df['採購最慢到貨日_TEMP'] = (
        pd.to_datetime(df['due_date']) - 
        df['buffer_days'].apply(lambda x: timedelta(days=x))
    )
    
    df['採購最慢到貨日'] = df['採購最慢到貨日_TEMP'].dt.strftime('%Y-%m-%d')
    
    # 清理輔助欄位並返回
    df = df.drop(columns=['due_date', 'buffer_days', '採購最慢到貨日_TEMP'], errors='ignore')
    return df


# --- UI 邏輯處理函式 (更新：支援附件URL) ---

# 設置預覽 URL 的函式
def set_preview_url(gcs_uri):
    """將 GCS URI 轉換為簽章 URL 並設置預覽狀態。"""
    
    if gcs_uri.startswith("gs://"):
        signed_url = get_signed_attachment_url(gcs_uri)
        if signed_url:
            st.session_state.preview_url = signed_url
            st.session_state.show_preview_modal = True
            st.rerun() # 觸發重新運行以顯示 Modal
        else:
            st.error("無法生成有效的附件簽章 URL。請檢查 GCE 服務帳戶權限。")
            
    # 如果不是 gs:// 格式，可能是手動輸入的外部 URL，我們允許嘗試預覽
    elif gcs_uri.startswith("http"):
        st.session_state.preview_url = gcs_uri
        st.session_state.show_preview_modal = True
        st.rerun()
    else:
        st.warning("無效的 GCS URI 或 URL。")


# 抽離寫入與重跑邏輯 (優化精簡)
def save_and_rerun(df_to_save, metadata_to_save, success_message=""):
    """將數據寫回 Sheets，並在成功後執行 st.rerun。"""
    
    if write_data_to_sheets(df_to_save, metadata_to_save):
        st.session_state.edited_dataframes = {} # 清除編輯狀態
        if success_message:
            st.success(success_message)
        st.rerun()
        
    pass


def handle_master_save():
    """批次處理所有 data_editor 的修改，並重新計算總價與預算。"""
    
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
            
            # --- 數據比較與更新 ---
            # 簡化更新邏輯
            updatable_cols = ['選取', '供應商', '單價', '數量', '狀態', '標記刪除', '附件URL'] # 包含附件URL
            for col in updatable_cols:
                if main_df.loc[main_idx, col] != new_row.get(col): # 使用 get 處理可能缺失的欄位
                    main_df.loc[main_idx, col] = new_row[col]
                    changes_detected = True
            
            # 處理日期解析
            try:
                date_str_parts = str(new_row['交期顯示']).strip().split(' ')
                date_part = date_str_parts[0]
                if main_df.loc[main_idx, '預計交貨日'] != date_part:
                    datetime.strptime(date_part, "%Y-%m-%d")
                    main_df.loc[main_idx, '預計交貨日'] = date_part
                    changes_detected = True
            except:
                pass
            
            # 重新計算總價
            current_price = float(main_df.loc[main_idx, '單價'])
            current_qty = float(main_df.loc[main_idx, '數量'])
            new_total = current_price * current_qty
            
            if main_df.loc[main_idx, '總價'] != new_total:
                main_df.loc[main_idx, '總價'] = new_total
                changes_detected = True
            
            affected_projects.add(main_df.loc[main_idx, '專案名稱'])

    if changes_detected:
        st.session_state.data = main_df.copy() # 寫回 session state 觸發更新
        
        # 更新 metadata 的最後修改時間
        for proj in affected_projects:
            if proj in st.session_state.project_metadata:
                st.session_state.project_metadata[proj]['last_modified'] = current_time_str
                
        # 使用精簡後的 save_and_rerun 函式
        save_and_rerun(
            st.session_state.data, 
            st.session_state.project_metadata, 
            success_message="✅ 資料已儲存！總價、總預算及 Google Sheets 已更新。"
        )

    else:
        st.info("沒有偵測到表格修改。")


# ... (省略 handle_batch_delete_quotes, trigger_delete_confirmation, cancel_delete_confirmation) ...
def handle_batch_delete_quotes():
    """根據 '標記刪除' 欄位，批次刪除報價。"""
    
    main_df = st.session_state.data.copy()
    
    # 優化: 合併 edited_dataframes (僅處理 '標記刪除' 欄位)
    combined_edited_df = pd.concat(
        [edited_df.set_index('ID')[['標記刪除']] for edited_df in st.session_state.edited_dataframes.values() if not edited_df.empty],
        axis=0, 
        ignore_index=False
    )
    
    if not combined_edited_df.empty:
        main_df = main_df.set_index('ID')
        main_df.update(combined_edited_df)
        main_df = main_df.reset_index()

    ids_to_delete = main_df[main_df['標記刪除'] == True]['ID'].tolist()
    
    if not ids_to_delete:
        st.warning("沒有項目被標記為刪除。")
        st.session_state.show_delete_confirm = False
        st.rerun()
        return

    st.session_state.data = main_df[main_df['標記刪除'] == False].drop(columns=['標記刪除'], errors='ignore')
    
    # 使用精簡後的 save_and_rerun 函式
    save_and_rerun(
        st.session_state.data, 
        st.session_state.project_metadata, 
        success_message=f"✅ 已成功刪除 {len(ids_to_delete)} 筆報價。Sheets 已更新。"
    )

def trigger_delete_confirmation():
    """點擊 '刪除已標記項目' 按鈕時，觸發確認流程。"""
    
    temp_df = st.session_state.data.copy()
    
    # 優化: 合併 edited_dataframes (僅處理 '標記刪除' 欄位)
    combined_edited_df = pd.concat(
        [edited_df.set_index('ID')[['標記刪除']] for edited_df in st.session_state.edited_dataframes.values() if not edited_df.empty],
        axis=0, 
        ignore_index=False
    )
    
    if not combined_edited_df.empty:
        temp_df = temp_df.set_index('ID')
        temp_df.update(combined_edited_df)
        temp_df = temp_df.reset_index()

    ids_to_delete = temp_df[temp_df['標記刪除'] == True]['ID'].tolist()
    
    if not ids_to_delete:
        st.warning("沒有項目被標記為刪除。請先在表格中勾選 '刪除?' 欄位。")
        st.session_state.show_delete_confirm = False
        return

    st.session_state.delete_count = len(ids_to_delete)
    st.session_state.show_delete_confirm = True
    st.rerun()

def cancel_delete_confirmation():
    """取消刪除確認。"""
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
        st.error(f"新的專案名稱 '{new_name}' 已存在，請使用不同名稱。")
        return

    # 1. 更新 Metadata
    meta = st.session_state.project_metadata.pop(target_proj)
    meta['due_date'] = new_date
    meta['last_modified'] = current_time_str
    st.session_state.project_metadata[new_name] = meta
    
    # 2. 更新 Dataframe
    st.session_state.data.loc[st.session_state.data['專案名稱'] == target_proj, '專案名稱'] = new_name
    
    # 使用精簡後的 save_and_rerun 函式
    save_and_rerun(
        st.session_state.data, 
        st.session_state.project_metadata, 
        success_message=f"✅ 專案已更新：{new_name}。Sheets 已更新。"
    )

def handle_delete_project(project_to_delete):
    """刪除選定的專案及其所有相關報價。"""
    
    if not project_to_delete:
        st.error("請選擇要刪除的專案。")
        return

    # 1. 刪除專案設定 (Metadata)
    if project_to_delete in st.session_state.project_metadata:
        del st.session_state.project_metadata[project_to_delete]

    # 2. 刪除所有相關報價 (Data)
    initial_count = len(st.session_state.data)
    st.session_state.data = st.session_state.data[
        st.session_state.data['專案名稱'] != project_to_delete
    ].reset_index(drop=True)
    
    deleted_count = initial_count - len(st.session_state.data)

    # 使用精簡後的 save_and_rerun 函式
    save_and_rerun(
        st.session_state.data, 
        st.session_state.project_metadata, 
        success_message=f"✅ 專案 **{project_to_delete}** 及其相關的 {deleted_count} 筆報價已成功刪除。Sheets 已更新。"
    )


def handle_add_new_project():
    """處理新增專案設定的邏輯"""
    project_name = st.session_state.new_proj_name
    project_due_date = st.session_state.new_proj_due_date
    buffer_days = st.session_state.new_proj_buffer_days
    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    if not project_name:
        st.error("專案名稱不能為空。")
        return
        
    if project_name in st.session_state.project_metadata:
        st.warning(f"專案 '{project_name}' 已存在，將更新其時程設定。")
    
    st.session_state.project_metadata[project_name] = {
        'due_date': project_due_date, 
        'buffer_days': buffer_days,
        'last_modified': current_time_str
    }
    
    # 使用精簡後的 save_and_rerun 函式
    save_and_rerun(
        st.session_state.data, 
        st.session_state.project_metadata, 
        success_message=f"✅ 已新增/更新專案設定：{project_name}。Sheets 已更新。"
    )

def handle_add_new_quote(latest_arrival_date, uploaded_file):
    """處理新增報價的邏輯 (V2.1.9 支援附件)"""
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

    status = st.session_state.quote_status
    
    if not project_name or not item_name_to_use:
        st.error("請確認已輸入專案名稱並選擇或輸入採購項目名稱。")
        return

    if project_name not in st.session_state.project_metadata:
        st.error(f"專案 '{project_name}' 的時程設定不存在。請先在 '➕ 新增專案' 區塊設定該專案的交期。")
        return

    total_price = price * qty
    
    # --- 附件處理核心邏輯 ---
    attachment_uri = ""
    next_id = st.session_state.next_id # 預先取得 ID
    if uploaded_file is not None:
        st.info(f"正在上傳附件 {uploaded_file.name}...")
        attachment_uri = upload_attachment_to_gcs(uploaded_file, next_id)
        if attachment_uri is None:
            # GCS 上傳失敗，停止新增報價
            return 
    # -------------------------
    
    st.session_state.project_metadata[project_name]['last_modified'] = current_time_str

    new_row = {
        'ID': next_id, '選取': False, '專案名稱': project_name, 
        '專案項目': item_name_to_use, '供應商': supplier, '單價': price, '數量': qty, 
        '總價': total_price, '預計交貨日': final_delivery_date.strftime('%Y-%m-%d'), 
        '狀態': status, '採購最慢到貨日': latest_arrival_date.strftime('%Y-%m-%d'), 
        '標記刪除': False,
        '附件URL': attachment_uri # 儲存 GCS URI (gs://...)
    }
    st.session_state.next_id += 1
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    
    # 使用精簡後的 save_and_rerun 函式
    save_and_rerun(
        st.session_state.data, 
        st.session_state.project_metadata, 
        success_message=f"✅ 已新增報價至 {project_name}！附件已儲存至 GCS。Sheets 已更新。"
    )


# --- Session State 初始化函式 (V2.2.0 新增預覽狀態) ---
def initialize_session_state():
    """初始化所有 Streamlit Session State 變數。從 Sheets 讀取數據。"""
    today = datetime.now().date()
    
    # 1. 數據與元數據載入 (只在 session 首次啟動時執行)
    if 'data' not in st.session_state:
        data_df, metadata_dict = load_data_from_sheets()
        
        st.session_state.data = data_df
        st.session_state.project_metadata = metadata_dict
        
    # 2. 使用 setdefault 進行統一初始化
    next_id_val = st.session_state.data['ID'].max() + 1 if not st.session_state.data.empty else 1
    
    initial_values = {
        'next_id': next_id_val,
        'edited_dataframes': {},
        'calculated_delivery_date': today,
        'show_delete_confirm': False,
        'delete_count': 0,
        'show_preview_modal': False, # 新增：控制預覽 Modal
        'preview_url': "",           # 新增：預覽的 URL
    }
    
    for key, value in initial_values.items():
        st.session_state.setdefault(key, value)
        
    # 確保 '標記刪除' 欄位存在
    if '標記刪除' not in st.session_state.data.columns:
        st.session_state.data['標記刪除'] = False
        
    # 確保 '附件URL' 欄位存在
    if '附件URL' not in st.session_state.data.columns:
        st.session_state.data['附件URL'] = ""


# --- 主應用程式核心邏輯 (在登入成功後調用) ---
def run_app():
    """運行應用程式的核心邏輯，在成功登入後調用。"""
    
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}") 
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    initialize_session_state()
    
    # --- 渲染附件預覽 Modal (V2.2.0 UX) ---
    if st.session_state.show_preview_modal:
        url = st.session_state.preview_url
        
        # 使用 Streamlit 內建的 st.modal
        with st.container():
            st.markdown(f"### 附件預覽", unsafe_allow_html=True)
            st.markdown("---")
            
            # 判斷檔案類型進行渲染 (使用 signed URL 判斷)
            # Signed URL 可能有 query params，所以只檢查開頭和結尾
            
            # 檢查是否為圖片
            is_image = any(ext in url.lower() for ext in ('.jpg', '.jpeg', '.png', '.gif'))
            
            if is_image:
                st.image(url, caption="圖片附件", use_column_width=True)
            elif ".pdf" in url.lower():
                st.info("PDF 檔案無法直接嵌入 Streamlit 進行預覽。")
                st.markdown(f"**下載連結:** [點此下載附件]({url})", unsafe_allow_html=True)
            elif url:
                st.warning("無法識別的檔案類型。")
                st.markdown(f"**原始連結:** [點此外部開啟]({url})", unsafe_allow_html=True)
            else:
                st.error("附件 URL 無效或不存在。")
            
            if st.button("關閉預覽", key="close_modal_btn"):
                st.session_state.show_preview_modal = False
                st.session_state.preview_url = ""
                st.rerun() # 關閉 Modal 後重新運行

    # 數據自動計算：在初始化後，計算最慢到貨日
    st.session_state.data = calculate_latest_arrival_dates(
        st.session_state.data, 
        st.session_state.project_metadata
    )
    
    if st.session_state.get('data_load_failed', False):
        st.warning("應用程式無法從 Google Sheets 載入數據，請檢查上方錯誤訊息。")
        
    today = datetime.now().date() 

    # --- UI 核心邏輯開始 ---
    
    # 格式化日期顯示
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
    
    # *** 側邊欄 UI 邏輯 *** <--- 將功能移動到這裡，並添加登出按鈕
    with st.sidebar:
        
        # 顯示登出按鈕 (已從 main() 移動到此處)
        st.button("登出", on_click=logout, type="secondary")
        st.markdown("---")

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
                    
                    if st.button("確認修改專案", type="primary"):
                        handle_project_modification()
                
                elif operation == "刪除專案":
                    st.markdown("##### 🗑️ 專案刪除 (⚠️ 警告)")
                    st.warning(f"您即將永久刪除專案 **{target_proj}** 及其所有相關報價資料。")
                    
                    if st.button(f"確認永久刪除 {target_proj}", type="secondary", help="此操作不可逆，將同時移除所有相關報價"):
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

            if st.button("儲存專案設定", key="btn_save_proj"):
                handle_add_new_project()
        
        st.markdown("---")
        
        # 區塊 3: 新增報價 (新增 file_uploader)
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

                st.caption(f"專案最慢到貨日: {latest_arrival_date.strftime('%Y-%m-%d')}")

            st.markdown("##### 採購項目選擇")
            
            unique_items = sorted(st.session_state.data['專案項目'].unique().tolist())
            item_options = ['新增項目...'] + unique_items

            selected_item = st.selectbox("選擇現有項目", item_options, key="quote_item_select")

            item_name_to_use = None
            if selected_item == '新增項目...':
                item_name_to_use = st.text_input("輸入新的採購項目名稱", key="quote_item_new_input")
            else:
                item_name_to_use = selected_item
            
            st.session_state.item_name_to_use_final = item_name_to_use
            
            st.text_input("供應商名稱", key="quote_supplier")
            st.number_input("單價 (TWD)", min_value=0, key="quote_price")
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
                num_b_days = st.number_input("工作日數", min_value=1, value=5, key="quote_num_b_days_input")
                final_delivery_date = add_business_days(today, int(num_b_days))
                st.session_state.calculated_delivery_date = final_delivery_date
            
            if date_input_type != "1. 指定日期":
                final_delivery_date = st.session_state.calculated_delivery_date
                st.caption(f"計算得出的交期：{final_delivery_date.strftime('%Y-%m-%d')}")

            st.selectbox("目前狀態", STATUS_OPTIONS, key="quote_status")
            
            st.markdown("---")
            st.markdown("##### 📎 上傳附件 (PDF/圖片)")
            # 新增檔案上傳元件
            uploaded_file = st.file_uploader(
                "選取附件",
                type=['pdf', 'jpg', 'jpeg', 'png'],
                key="new_quote_file_uploader"
            )

            if st.button("新增資料", key="btn_add_quote"):
                handle_add_new_quote(latest_arrival_date, uploaded_file)


    # *** 儀表板區塊 ***
    total_projects, total_budget, risk_items, pending_quotes = calculate_dashboard_metrics(df, st.session_state.project_metadata)

    st.subheader("📊 總覽儀表板")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class='metric-box'>
            <div class='metric-title'>專案總數</div>
            <div class='metric-value'>{total_projects}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='metric-box' style='background-color:#21442c;'>
            <div class='metric-title'>預估/已選總預算</div>
            <div class='metric-value'>${total_budget:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div class='metric-box' style='background-color:#5c2d2d;'>
            <div class='metric-title'>交期風險項目</div>
            <div class='metric-value'>{risk_items}</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div class='metric-box' style='background-color:#2a3b5c;'>
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
                cancel_delete_confirmation()

    st.markdown("---")

    # *** 專案 Expander 列表 (核心修改) ***
    
    for proj_name, proj_data in project_groups:
        meta = st.session_state.project_metadata.get(proj_name, {})
        proj_budget = calculate_project_budget(df, proj_name)
        
        last_modified = meta.get('last_modified', 'N/A')
        
        header_html = f"""
        <span class='project-header'>💼 專案: {proj_name}</span> &nbsp;|&nbsp; 
        <span class='project-header'>總預算: ${proj_budget:,.0f}</span> &nbsp;|&nbsp; 
        <span class='meta-info'>交期: {meta.get('due_date')}</span> 
        <span style='float:right; font-size:14px; color:#FFC107;'>🕒 最後修改: {last_modified}</span>
        """
        
        with st.expander(label=f"專案：{proj_name} (點擊展開)", expanded=False):
            st.markdown(header_html, unsafe_allow_html=True)
            
            for item_name, item_data in proj_data.groupby('專案項目'):
                
                has_selection = item_data['選取'].any()
                sub_total = item_data[item_data['選取']]['總價'].sum() if has_selection else item_data['總價'].min()
                calc_method = "(已選)" if has_selection else "(預估)"
                
                # 顯示項目標題和預算
                st.markdown(f"""
                <span class='item-header'>📦 {item_name}</span> 
                <span class='meta-info'> | 計入: ${sub_total:,.0f} {calc_method}</span>
                """, unsafe_allow_html=True)
                
                # 預覽按鈕 (V2.2.0 UX)
                attachment_urls = item_data['附件URL'].tolist()
                
                if any(url and (url.startswith("gs://") or url.startswith("http")) for url in attachment_urls):
                    
                    # 找出第一個有效的 URI/URL 作為預覽對象
                    first_valid_uri = next((url for url in attachment_urls if url and (url.startswith("gs://") or url.startswith("http"))), None)
                    
                    # 判斷是否為圖片，用不同顏色顯示
                    is_image_guess = any(ext in first_valid_uri.lower() for ext in ('.jpg', '.jpeg', '.png', '.gif'))
                    button_type = "secondary" if is_image_guess else "primary"
                    button_text = "圖片預覽" if is_image_guess else "附件預覽"

                    # 為了讓按鈕不與 data_editor 擠在一起，我們將它放在一個專門的 col
                    col_spacer, col_preview_btn = st.columns([0.85, 0.15])
                    with col_preview_btn:
                        if st.button(button_text, key=f"preview_{item_name}_{proj_name}", type=button_type):
                            set_preview_url(first_valid_uri) # 傳遞 GCS URI 或外部 URL
                
                editable_df = item_data.copy()
                editor_key = f"editor_{proj_name}_{item_name}"
                
                edited_df_value = st.data_editor(
                    editable_df[['ID', '選取', '供應商', '單價', '數量', '總價', '交期顯示', '狀態', '附件URL', '標記刪除']],
                    column_config={
                        "ID": st.column_config.Column("ID", disabled=True, width="tiny"), 
                        "選取": st.column_config.CheckboxColumn("選取", width="tiny"), 
                        "供應商": st.column_config.Column("供應商", disabled=False), 
                        "單價": st.column_config.NumberColumn("單價", format="$%d"),
                        "數量": st.column_config.NumberColumn("數量"),
                        "總價": st.column_config.NumberColumn("總價", format="$%d", disabled=True),
                        "交期顯示": st.column_config.TextColumn("預計交貨日 (YYYY-MM-DD)", width="medium", help="可編輯，圖示會自動更新"),
                        "狀態": st.column_config.SelectboxColumn("狀態", options=STATUS_OPTIONS),
                        "標記刪除": st.column_config.CheckboxColumn("刪除?", width="tiny", help="勾選後點擊上方按鈕執行刪除"), 
                        # 附件 URL 欄位：可編輯，儲存 GCS URI (gs://...)
                        "附件URL": st.column_config.TextColumn("附件URL", help="GCS URI 或外部連結", disabled=False, width="medium"), 
                    },
                    key=editor_key,
                    hide_index=True,
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


# --- 程式進入點 ---
def main():
    # 執行登入驗證 (自定義 V1.0.0 邏輯)
    login_form()
    
    # --- 僅在驗證通過後執行後續程式碼 ---
    if st.session_state.authenticated:
        # 顯示登出按鈕 (已移動到 run_app 中的 with st.sidebar 區塊)

        # 執行應用程式核心邏輯
        run_app() 
        
if __name__ == "__main__":
    main()
