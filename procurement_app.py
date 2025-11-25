import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os # 新增: 導入 os 模組用於讀取環境變數

# --- 應用程式設定 ---
APP_VERSION = "v2.0.0 (Google Sheets Beta)" # 版本更新為 v2.0.0
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"]

# --- 數據源配置 (GCP/本地通用配置) ---
# 檢查環境變數 GCE_SHEET_URL，如果不存在，則使用 st.secrets (用於 Streamlit Cloud)
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
    
    # GCE 部署時，從環境變數讀取憑證檔案路徑
    try:
        # GSHEETS_CREDENTIALS_PATH 是你在 systemd 服務檔案中設置的環境變數
        GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"] 
    except KeyError:
        # 如果在 GCE 環境中找不到路徑，報錯並使用 None (讓連接器使用預設行為)
        st.error("❌ 錯誤：在 GCE 環境中未找到 GSHEETS_CREDENTIALS_PATH 環境變數。")
        GSHEETS_CREDENTIALS = None 
        
else:
    # Streamlit Cloud 或本地 .secrets 部署
    # 這裡假設你在 .streamlit/secrets.toml 中有一個 app_config 區塊
    try:
        SHEET_URL = st.secrets["app_config"]["sheet_url"]
    except KeyError:
        st.error("❌ 錯誤：請在 secrets.toml 檔案中配置 [app_config] sheet_url。")
        SHEET_URL = None
    GSHEETS_CREDENTIALS = None # st.connection 會自動處理 st.secrets

DATA_SHEET_NAME = "採購總表"     # 包含報價數據的工作表名稱
METADATA_SHEET_NAME = "專案設定" # 包含專案設定的工作表名稱


# 設定頁面標題與寬度
st.set_page_config(page_title=f"專案採購小幫手 {APP_VERSION}", layout="wide")

# --- CSS 樣式修正 (不變) ---
CUSTOM_CSS = """
<style>
/* 1. 基礎樣式與顏色 */
.streamlit-expanderContent { padding-left: 1rem !important; padding-right: 1rem !important; padding-bottom: 1rem !important; }
.project-header { font-size: 20px !important; font-weight: bold !important; color: #FAFAFA; }
.item-header { font-size: 16px !important; font-weight: 600 !important; color: #E0E0E0; }
.meta-info { font-size: 14px !important; color: #9E9E9E; font-weight: normal; }
div[data-baseweb="select"] > div, div[data-baseweb="base-input"] > input, div[data-baseweb="input"] > div { background-color: #262730 !important; color: white !important; -webkit-text-fill-color: white !important; }
div[data-baseweb="popover"], div[data-baseweb="menu"] { background-color: #262730 !important; }
div[data-baseweb="option"] { color: white !important; }
li[aria-selected="true"] { background-color: #FF4B4B !important; color: white !important; }

/* 儀表板樣式 */
.metric-box {
    padding: 10px 15px;
    border-radius: 8px;
    margin-bottom: 10px;
    background-color: #262730;
    text-align: center;
}
.metric-title {
    font-size: 14px;
    color: #9E9E9E;
    margin-bottom: 5px;
}
.metric-value {
    font-size: 24px;
    font-weight: bold;
}
</style>
"""

# --- 數據讀取與寫入函式 (核心修改) ---

@st.cache_data(ttl=600, show_spinner="連線 Google Sheets...") # 緩存數據，每 10 分鐘檢查一次 Sheets
def load_data_from_sheets():
    """使用 st.connection 讀取 Google Sheets 中的採購總表和專案設定。"""
    
    if not SHEET_URL:
        st.stop() # 如果 URL 為空，則停止運行
    
    try:
        # 根據部署環境動態調整連接器參數
        if GSHEETS_CREDENTIALS and os.path.exists(GSHEETS_CREDENTIALS):
            # GCE 模式: 使用憑證檔案路徑 (在 systemd 中設定)
            conn = st.connection("gsheets", type=st.connection_factory.google_sheets, 
                                 credentials=GSHEETS_CREDENTIALS) 
        else:
            # Streamlit Cloud/本地模式: 使用 st.secrets
            conn = st.connection("gsheets", type=st.connection_factory.google_sheets)
        
        # 1. 讀取採購總表 (Data)
        data_df = conn.read(spreadsheet=SHEET_URL, worksheet=DATA_SHEET_NAME, ttl=5)
        
        # 數據類型轉換與處理
        data_df = data_df.astype({
            'ID': 'Int64', '選取': 'bool', '單價': 'float', '數量': 'Int64', '總價': 'float'
        })
        # 確保 '標記刪除' 欄位存在
        if '標記刪除' not in data_df.columns:
            data_df['標記刪除'] = False
        
        # 2. 讀取專案設定 (Metadata)
        metadata_df = conn.read(spreadsheet=SHEET_URL, worksheet=METADATA_SHEET_NAME, ttl=5)
        
        # 轉換為 dictionary 格式
        project_metadata = {}
        if not metadata_df.empty:
            for index, row in metadata_df.iterrows():
                project_metadata[row['專案名稱']] = {
                    'due_date': datetime.strptime(str(row['專案交貨日']), '%Y-%m-%d').date(),
                    'buffer_days': int(row['緩衝天數']),
                    'last_modified': str(row['最後修改'])
                }

        return data_df, project_metadata

    except Exception as e:
        # 如果載入失敗，我們將使用空白 DataFrame 防止應用程式崩潰
        st.error(f"❌ 數據載入失敗！請檢查 Google Sheets 憑證、URL/工作表名稱和 API 權限。")
        st.code(f"錯誤訊息: {e}")
        
        # 設置空的 DataFrame 結構以防止後續程式碼崩潰
        empty_data = pd.DataFrame(columns=['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '標記刪除'])
        empty_meta = {}
        st.session_state.data_load_failed = True
        return empty_data, empty_meta

def write_data_to_sheets(df_to_write, metadata_to_write):
    """將採購總表和專案設定寫回 Google Sheets。"""
    if st.session_state.get('data_load_failed', False):
        st.warning("數據載入失敗，已禁用寫入 Sheets。")
        return False
        
    try:
        if GSHEETS_CREDENTIALS and os.path.exists(GSHEETS_CREDENTIALS):
            conn = st.connection("gsheets", type=st.connection_factory.google_sheets, credentials=GSHEETS_CREDENTIALS) 
        else:
            conn = st.connection("gsheets", type=st.connection_factory.google_sheets)
        
        # 1. 寫入採購總表 (Data) - 需先移除 '標記刪除' 和 '交期顯示' 欄位
        df_export = df_to_write.drop(columns=['標記刪除', '交期顯示'], errors='ignore')
        conn.write(df_export, spreadsheet=SHEET_URL, worksheet=DATA_SHEET_NAME)
        
        # 2. 寫入專案設定 (Metadata)
        metadata_list = [
            {'專案名稱': name, 
             '專案交貨日': data['due_date'].strftime('%Y-%m-%d'),
             '緩衝天數': data['buffer_days'], 
             '最後修改': data['last_modified']}
            for name, data in metadata_to_write.items()
        ]
        metadata_df = pd.DataFrame(metadata_list)
        conn.write(metadata_df, spreadsheet=SHEET_URL, worksheet=METADATA_SHEET_NAME)
        
        st.cache_data.clear() # 清除緩存強制下次重新載入
        return True
    except Exception as e:
        st.error(f"❌ 數據寫回 Google Sheets 失敗！請檢查 Sheets 權限。")
        st.code(f"錯誤訊息: {e}")
        return False


# --- Session State 初始化函式 (優化) ---
def initialize_session_state():
    """初始化所有 Streamlit Session State 變數。從 Sheets 讀取數據。"""
    today = datetime.now().date()
    
    # *** 數據初始化 (從 Google Sheets 讀取) ***
    # 這裡我們只在 session state 中沒有 'data' 時才嘗試載入
    if 'data' not in st.session_state or 'project_metadata' not in st.session_state:
        # 使用 load_data_from_sheets 讀取數據
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


# --- 輔助函式區 (add_business_days, calculate_dashboard_metrics, calculate_project_budget 不變) ---

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


# 刪除報價邏輯：批次刪除 (增加 Sheets 寫入)
def handle_batch_delete_quotes():
    """根據 '標記刪除' 欄位，批次刪除報價。"""
    
    main_df = st.session_state.data.copy()
    
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
        return

    st.session_state.data = main_df[main_df['標記刪除'] == False].drop(columns=['標記刪除'], errors='ignore')
    
    # *** 數據寫回 Sheets ***
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.session_state.show_delete_confirm = False
        st.success(f"✅ 已成功刪除 {len(ids_to_delete)} 筆報價，並同步到 Google Sheets！")
        st.rerun()

# 批次刪除的觸發函式 (不變)
def trigger_delete_confirmation():
    """點擊 '刪除已標記項目' 按鈕時，觸發確認流程。"""
    
    temp_df = st.session_state.data.copy()
    
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


# 處理總表儲存邏輯 - 增加 Sheets 寫入
def handle_master_save():
    """批次處理所有 data_editor 的修改，並重新計算總價與預算。"""
    
    if not st.session_state.edited_dataframes:
        st.info("沒有偵測到表格修改。")
        return

    main_df = st.session_state.data.copy()
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
            if main_df.loc[main_idx, '選取'] != new_row['選取']:
                main_df.loc[main_idx, '選取'] = new_row['選取']
                changes_detected = True
                
            updatable_cols = ['供應商', '單價', '數量', '狀態']
            for col in updatable_cols:
                if main_df.loc[main_idx, col] != new_row[col]:
                    main_df.loc[main_idx, col] = new_row[col]
                    changes_detected = True
            
            # 處理日期解析
            try:
                date_str_parts = str(new_row['交期顯示']).strip().split(' ') 
                date_part = date_str_parts[0] 
                
                if str(main_df.loc[main_idx, '預計交貨日']) != date_part:
                    datetime.strptime(date_part, "%Y-%m-%d")
                    main_df.loc[main_idx, '預計交貨日'] = date_part
                    changes_detected = True
            except:
                st.warning(f"ID {original_id} 日期格式錯誤，請使用 YYYY-MM-DD") 
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
        st.session_state.data = main_df.copy()
        
        for proj in affected_projects:
            if proj in st.session_state.project_metadata:
                st.session_state.project_metadata[proj]['last_modified'] = current_time_str
        
        # *** 數據寫回 Sheets ***
        if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
            st.success("✅ 資料已儲存！總價與總預算已更新，並同步到 Google Sheets！")
            st.rerun()

    else:
        st.info("沒有偵測到表格修改。")

# 處理專案修改邏輯 - 增加 Sheets 寫入
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
    
    # *** 數據寫回 Sheets ***
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"專案已更新：{new_name}，並同步到 Google Sheets！")
        st.rerun()

# 處理專案刪除邏輯 - 增加 Sheets 寫入
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

    # *** 數據寫回 Sheets ***
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"🗑️ 專案 **{project_to_delete}** 及其相關的 {deleted_count} 筆報價已成功刪除，並同步到 Google Sheets！")
        st.rerun()

# 處理新增專案邏輯 - 增加 Sheets 寫入
def handle_add_new_project():
    """處理新增專案設定的邏輯"""
    project_name = st.session_state.new_proj_name
    project_due_date = st.session_state.new_proj_due_date
    buffer_days = st.session_state.new_proj_buffer_days
    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    if not project_name:
        st.error("專案名稱不能為空。")
        return
        
    is_update = project_name in st.session_state.project_metadata

    st.session_state.project_metadata[project_name] = {
        'due_date': project_due_date, 
        'buffer_days': buffer_days,
        'last_modified': current_time_str
    }
    
    # *** 數據寫回 Sheets ***
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        if is_update:
            st.success(f"已更新專案設定：{project_name}，並同步到 Google Sheets！")
        else:
            st.success(f"已新增專案設定：{project_name}，並同步到 Google Sheets！")
        st.rerun()

# 處理新增報價邏輯 - 增加 Sheets 寫入
def handle_add_new_quote(latest_arrival_date):
    """處理新增報價的邏輯"""
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
    
    st.session_state.project_metadata[project_name]['last_modified'] = current_time_str

    new_row = {
        'ID': st.session_state.next_id, '選取': False, '專案名稱': project_name, 
        '專案項目': item_name_to_use, '供應商': supplier, '單價': price, '數量': qty, 
        '總價': total_price, '預計交貨日': final_delivery_date.strftime('%Y-%m-%d'), 
        '狀態': status, '採購最慢到貨日': latest_arrival_date.strftime('%Y-%m-%d'), 
        '標記刪除': False
    }
    
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    st.session_state.next_id += 1
    
    # *** 數據寫回 Sheets ***
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.success(f"已新增報價至 {project_name}，並同步到 Google Sheets！")
        st.rerun()


# --- 主要應用程式 (main 函式保持不變) ---
def main():
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    initialize_session_state()

    today = datetime.now().date() 

    # --- 側邊欄 ---
    with st.sidebar:
        
        # --- 區塊 1: 修改/刪除專案 ---
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
        
        # --- 區塊 2 & 3 (不變) ---
        with st.expander("➕ 新增/設定專案時程", expanded=False):
            st.text_input("專案名稱 (Project Name)", key="new_proj_name")
            
            project_due_date = st.date_input("專案交貨日 (Project Due Date)", value=today + timedelta(days=30), key="new_proj_due_date")
            buffer_days = st.number_input("採購緩衝天數 (天)", min_value=0, value=7, key="new_proj_buffer_days")
            
            latest_arrival_date_proj = project_due_date - timedelta(days=int(buffer_days))
            st.caption(f"計算得出最慢到貨日：{latest_arrival_date_proj.strftime('%Y年%m月%d日')}")

            if st.button("儲存專案設定", key="btn_save_proj"):
                handle_add_new_project()
        
        st.markdown("---")
        
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
            
            if st.button("新增資料", key="btn_add_quote"):
                handle_add_new_quote(latest_arrival_date)


    # --- 主介面 ---
    df = st.session_state.data
    
    def format_date_with_icon(row):
        date_str = str(row['預計交貨日'])
        try:
            v_date = pd.to_datetime(row['預計交貨日']).date()
            l_date = pd.to_datetime(row['採購最慢到貨日']).date()
            icon = "🔴" if v_date > l_date else "✅"
            return f"{date_str} {icon}"
        except:
            return date_str

    if not df.empty:
        df['交期顯示'] = df.apply(format_date_with_icon, axis=1)

    project_groups = df.groupby('專案名稱')
    
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
                
                st.markdown(f"""
                <span class='item-header'>📦 {item_name}</span> 
                <span class='meta-info'> | 計入: ${sub_total:,.0f} {calc_method}</span>
                """, unsafe_allow_html=True)

                editable_df = item_data.copy()
                editor_key = f"editor_{proj_name}_{item_name}"
                
                edited_df_value = st.data_editor(
                    editable_df[['ID', '選取', '供應商', '單價', '數量', '總價', '交期顯示', '狀態', '標記刪除']],
                    column_config={
                        "ID": st.column_config.Column("ID", disabled=True, width="tiny"), 
                        "選取": st.column_config.CheckboxColumn("選取", width="tiny"), 
                        "供應商": st.column_config.Column("供應商", disabled=True),
                        "單價": st.column_config.NumberColumn("單價", format="$%d"),
                        "數量": st.column_config.NumberColumn("數量"),
                        "總價": st.column_config.NumberColumn("總價", format="$%d", disabled=True),
                        "交期顯示": st.column_config.TextColumn("預計交貨日 (YYYY-MM-DD)", width="medium", help="可編輯，圖示會自動更新"),
                        "狀態": st.column_config.SelectboxColumn("狀態", options=STATUS_OPTIONS),
                        "標記刪除": st.column_config.CheckboxColumn("刪除?", width="tiny", help="勾選後點擊上方按鈕執行刪除"), 
                    },
                    key=editor_key,
                    hide_index=True,
                    use_container_width=True,
                    height=150 + (len(item_data) * 35) if len(item_data) > 3 else 150,
                    disabled=is_locked
                )
                
                st.session_state.edited_dataframes[item_name] = edited_df_value 
                st.markdown("---")

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("💾 資料匯出")
    st.download_button("📥 下載 Excel 報表", 
                      convert_df_to_excel(df), 
                      f'procurement_report_{datetime.now().strftime("%Y%m%d")}.xlsx', 
                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    main()
