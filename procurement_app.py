import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os 
import json
import gspread
import logging

# 配置 Streamlit 日誌，以便將錯誤寫入 journalctl
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# --- 應用程式設定 ---
APP_VERSION = "v2.1.5 (Secure Login)"
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"]

# --- 數據源配置 (GCE/本地通用配置) ---
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
    try:
        GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"] 
    except KeyError:
        logging.error("GCE_SHEET_URL is set, but GSHEETS_CREDENTIALS_PATH is missing.")
        st.error("❌ 錯誤：在 GCE 環境中未找到 GSHEETS_CREDENTIALS_PATH 環境變數。")
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
</style>
"""

# --- 登入與安全函式 (使用 os.environ 安全讀取) ---

def logout():
    """登出函式：清除驗證狀態並重新運行。"""
    st.session_state["authenticated"] = False
    st.rerun()

def login_form():
    """渲染登入表單並處理密碼驗證。"""
    
    # 從 systemd 環境變數中讀取密碼 (安全關鍵!)
    DEFAULT_USERNAME = os.environ.get("AUTH_USERNAME", "dev_user")
    DEFAULT_PASSWORD = os.environ.get("AUTH_PASSWORD", "dev_pwd")
    
    # 這裡我們使用 os.environ，而不是 st.secrets
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
            
            # 用戶名輸入框預設為環境變數的值，不允許用戶更改
            username = st.text_input("用戶名", key="login_username", value=credentials["username"], disabled=True)
            password = st.text_input("密碼", type="password", key="login_password")
            
            if st.button("登入", type="primary"):
                # 驗證用戶名和密碼
                if username.strip() == credentials["username"].strip() and password == credentials["password"]:
                    st.session_state["authenticated"] = True
                    st.toast("✅ 登入成功！")
                    st.rerun()
                else:
                    st.error("用戶名或密碼錯誤。")
            
    # 如果未驗證，阻止執行後續程式碼
    st.stop() 


# --- 數據讀取與寫入函式 (核心修改: 使用 gspread) ---

@st.cache_data(ttl=600, show_spinner="連線 Google Sheets...")
def load_data_from_sheets():
    # ... (Gspread 讀取邏輯保持不變) ...
    if not SHEET_URL:
        st.info("❌ Google Sheets URL 尚未配置。使用空的數據結構。")
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除'])
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
        
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除'])
        st.session_state.data_load_failed = True
        return empty_data, {}


def write_data_to_sheets(df_to_write, metadata_to_write):
    # ... (Sheets 寫入邏輯保持不變) ...
    if st.session_state.get('data_load_failed', False) or not SHEET_URL:
        st.warning("數據載入失敗或 URL 未配置，已禁用寫入 Sheets。")
        return False
        
    try:
        # --- 1. 授權與認證 ---
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # --- 2. 寫入採購總表 (Data) ---
        df_export = df_to_write.drop(columns=['標記刪除', '交期顯示'], errors='ignore')
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_ws.clear()
        data_ws.update([df_export.columns.values.tolist()] + df_export.values.tolist())
        
        # --- 3. 寫入專案設定 (Metadata) ---
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


# --- 輔助函式區 (這裡省略所有 handle_xxx 和 calculate_xxx 函式，假設它們已在檔案中定義) ---
# ... (所有輔助函式定義) ...

# --- Session State 初始化函式 (使用 Gspread 邏輯) ---
def initialize_session_state():
    # ... (保持原邏輯) ...
    today = datetime.now().date()
    
    if 'data' not in st.session_state or 'project_metadata' not in st.session_state:
        data_df, metadata_dict = load_data_from_sheets()
        
        st.session_state.data = data_df
        st.session_state.project_metadata = metadata_dict
        
    if '標記刪除' not in st.session_state.data.columns:
        st.session_state.data['標記刪除'] = False
            
    if 'next_id' not in st.session_state:
        st.session_state.next_id = st.session_state.data['ID'].max() + 1 if not st.session_state.data.empty else 1
    
    # ... (省略其餘初始化邏輯) ...


# --- 主應用程式核心邏輯 (在登入成功後調用) ---
def run_app():
    """運行應用程式的核心邏輯，在成功登入後調用。"""
    
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}") 
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    initialize_session_state()

    # 數據自動計算：在初始化後，計算最慢到貨日
    # st.session_state.data = calculate_latest_arrival_dates(...) 
    
    if st.session_state.get('data_load_failed', False):
        st.warning("應用程式無法從 Google Sheets 載入數據，請檢查上方錯誤訊息。")
        
    # --- V2.1.3 的所有 UI 邏輯 (儀表板、側邊欄、data_editor 等) 貼在這裡 ---
    
    # ... (UI 邏輯，例如 subheader, 儀表板, Expander, data_editor) ...
    
    
# --- 程式進入點 ---
def main():
    # 執行登入驗證 (自定義 V1.0.0 邏輯)
    login_form()
    
    # --- 僅在驗證通過後執行後續程式碼 ---
    if st.session_state.authenticated:
        # 顯示登出按鈕
        st.sidebar.button("登出", on_click=logout) 

        # 執行應用程式核心邏輯
        run_app() 
        
if __name__ == "__main__":
    main()
