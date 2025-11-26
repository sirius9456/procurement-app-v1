import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os 
import json
import gspread

# 引入登入、日誌和配置模組
import logging
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth


# 配置 Streamlit 日誌，以便將錯誤寫入 journalctl
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# --- 應用程式設定 ---
APP_VERSION = "v2.1.2 (Login Integrated)" # 版本更新為 v2.1.2
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"]

# --- 數據源配置 (GCE/本地通用配置) ---
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
    try:
        GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"] 
    except KeyError:
        # 錯誤日誌記錄
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


# 設定頁面標題與寬度 (必須在 Streamlit 程式碼中第一個調用)
st.set_page_config(page_title=f"專案採購小幫手 {APP_VERSION}", layout="wide")

# --- CSS 樣式修正 (保持不變) ---
CUSTOM_CSS = """
<style>
/* 保持原樣，確保頁面風格一致 */
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

# --- 數據讀取與寫入函式 (核心修改: 使用 gspread) ---

@st.cache_data(ttl=600, show_spinner="連線 Google Sheets...")
def load_data_from_sheets():
    """直接使用 gspread 讀取 Google Sheets 中的數據。"""
    
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
        # 記錄完整的錯誤追溯到 systemd journal
        logging.exception("Google Sheets 數據載入時發生致命錯誤！") 
        
        st.error(f"❌ 數據載入失敗！請檢查 Sheets 分享權限、工作表名稱或憑證檔案。")
        st.code(f"錯誤訊息: {e}")
        
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除'])
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


# --- 輔助函式區 ---

def add_business_days(start_date, num_days):
    # 保持原邏輯
    current_date = start_date
    days_added = 0
    while days_added < num_days:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5: days_added += 1
    return current_date

@st.cache_data
def convert_df_to_excel(df):
    # 保持原邏輯
    df_export = df.drop(columns=['標記刪除', '交期顯示'], errors='ignore')
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='採購報價總表')
    
    processed_data = output.getvalue()
    return processed_data


@st.cache_data(show_spinner=False)
def calculate_dashboard_metrics(df_state, project_metadata_state):
    # 保持原邏輯
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
    # 保持原邏輯
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

# --- 專案交期自動計算邏輯 (V2.1.1 優化) ---
@st.cache_data(show_spinner=False)
def calculate_latest_arrival_dates(df, metadata):
    """根據專案設定，計算每個採購項目的採購最慢到貨日。"""
    
    if df.empty or not metadata:
        return df

    metadata_df = pd.DataFrame.from_dict(metadata, orient='index')
    metadata_df = metadata_df.reset_index().rename(columns={'index': '專案名稱'})
    
    metadata_df['due_date'] = metadata_df['due_date'].apply(lambda x: pd.to_datetime(x).date())
    metadata_df['buffer_days'] = metadata_df['buffer_days'].astype(int)

    df = pd.merge(df, metadata_df[['專案名稱', 'due_date', 'buffer_days']], on='專案名稱', how='left')

    df['採購最慢到貨日_NEW'] = (
        df['due_date'] - 
        df['buffer_days'].apply(lambda x: timedelta(days=x))
    ).dt.strftime('%Y-%m-%d')
    
    df['採購最慢到貨日'] = df['採購最慢到貨日_NEW']
    
    df = df.drop(columns=['due_date', 'buffer_days', '採購最慢到貨日_NEW'], errors='ignore')
    return df

# ... (省略 handle_xxx 函式，假設它們已正確定義在檔案中) ...
# 注意：所有 handle_xxx 函式 (如 handle_master_save) 都應在 run_app 之前定義
# --------------------------------------------------------------------------

# --- Session State 初始化函式 (優化) ---
def initialize_session_state():
    """初始化所有 Streamlit Session State 變數。從 Sheets 讀取數據。"""
    # 保持原邏輯
    today = datetime.now().date()
    
    if 'data' not in st.session_state or 'project_metadata' not in st.session_state:
        data_df, metadata_dict = load_data_from_sheets()
        
        st.session_state.data = data_df
        st.session_state.project_metadata = metadata_dict
        
    if '標記刪除' not in st.session_state.data.columns:
        st.session_state.data['標記刪除'] = False
            
    if 'next_id' not in st.session_state:
        st.session_state.next_id = st.session_state.data['ID'].max() + 1 if not st.session_state.data.empty else 1
    
    if 'edited_dataframes' not in st.session_state:
        st.session_state.edited_dataframes = {}

    if 'calculated_delivery_date' not in st.session_state:
        st.session_state.calculated_delivery_date = today
        
    if 'show_delete_confirm' not in st.session_state:
        st.session_state.show_delete_confirm = False
    if 'delete_count' not in st.session_state:
        st.session_state.delete_count = 0


# --- 主應用程式核心邏輯 (原 main 函式，現改名為 run_app) ---
def run_app():
    """運行應用程式的核心邏輯，在成功登入後調用。"""
    
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}") 
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    initialize_session_state()

    # 數據自動計算：在初始化後，計算最慢到貨日
    st.session_state.data = calculate_latest_arrival_dates(
        st.session_state.data, 
        st.session_state.project_metadata
    )
    
    # 如果數據載入失敗，顯示警告
    if st.session_state.get('data_load_failed', False):
        st.warning("應用程式無法從 Google Sheets 載入數據，請檢查上方錯誤訊息。")
        
    today = datetime.now().date() 

    # --- UI 核心邏輯開始 ---
    
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
    
    # ... (此處省略儀表板、批次操作、Expander 和 data_editor 等 UI 代碼，
    # 確保你將 V2.1.0 版本中的所有 UI 代碼貼到這裡，並使用 run_app 函式) ...

    st.subheader("📊 總覽儀表板")
    # 此處應包含儀表板 UI 邏輯
    
    st.markdown("---")
    
    # 此處應包含批次操作按鈕和邏輯
    
    st.markdown("---")
    
    # 此處應包含專案 Expander 列表
    for proj_name, proj_data in project_groups:
        # ... (Expander 和 data_editor 邏輯) ...
        pass
    
    # ... (UI 核心邏輯結束) ...

# --- 登入邏輯 (新的主要入口點) ---
def main():
    # --- 1. 登入配置 ---
    try:
        # 從 config.yaml 載入設定
        with open('config.yaml') as file:
            config = yaml.load(file, Loader=SafeLoader)
    except FileNotFoundError:
        st.error("配置檔案 config.yaml 找不到！請確保檔案存在並命名正確。")
        return
    except Exception as e:
        st.error(f"無法解析 config.yaml 檔案: {e}")
        return

    # 實例化 Authenticator
    authenticator = stauth.Authenticate(
        config['credentials'],
        config['cookie']['name'],
        config['cookie']['key'],
        config['cookie']['expiry_days']
    )

    st.subheader("🛡️ 專案採購管理工具 - 登入驗證") 

    # --- 2. 顯示登入表單 ---
    name, authentication_status, username = authenticator.login('Login')

    # --- 3. 檢查登入狀態並執行應用程式 ---
    if st.session_state["authentication_status"]:
        # 成功登入
        
        # 側邊欄顯示登出按鈕和歡迎訊息
        with st.sidebar:
            # 登出按鈕放在 sidebar
            st.sidebar.markdown("---") # 添加分隔線
            authenticator.logout('登出', 'main') # 保持在主頁面顯示登出
            st.sidebar.write(f'歡迎, {st.session_state["name"]}')

        # 執行應用程式核心邏輯
        run_app() 
        
    elif st.session_state["authentication_status"] is False:
        st.error('用戶名/密碼錯誤')
        
    elif st.session_state["authentication_status"] is None:
        st.warning('請輸入你的用戶名和密碼')


# --- 程式進入點 ---
if __name__ == "__main__":
    main()

