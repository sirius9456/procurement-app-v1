import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date 
from io import BytesIO
import os 
import json
import gspread
import logging
import time
import base64 
# GCS 相關導入
from google.cloud import storage 
from google.oauth2 import service_account

# ******************************
# *--- 1. 全域設定與常數 ---*
# ******************************

# 配置 Streamlit 日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 版本號
APP_VERSION = "V2.4.0 (UI Polish)" 

# 時間格式
DATE_FORMAT = "%Y-%m-%d"
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# --- Google Sheets URL 設定 ---
if "GCE_SHEET_URL" in os.environ:
    SHEET_URL = os.environ["GCE_SHEET_URL"]
else:
    try:
        SHEET_URL = st.secrets["spreadsheet"]["url"]
    except:
        SHEET_URL = "https://docs.google.com/spreadsheets/d/16vSMLx-GYcIpV2cuyGIeZctvA2sI8zcqh9NKKyrs-uY/edit?usp=sharing"

# 工作表名稱
DATA_SHEET_NAME = '採購總表_測試'
METADATA_SHEET_NAME = '專案設定_測試'

# --- GCS 設定 ---
GCS_BUCKET_NAME = "procurement-attachments-bucket"
GCS_FOLDER_PATH = "attachments"
GCS_BASE_URL = f"https://storage.googleapis.com/{GCS_BUCKET_NAME}"

# --- 憑證路徑設定 ---
if "GSHEETS_CREDENTIALS_PATH" in os.environ:
    GSHEETS_CREDENTIALS = os.environ["GSHEETS_CREDENTIALS_PATH"]
elif os.path.exists("secrets/google_sheets_credentials.json"):
    GSHEETS_CREDENTIALS = "secrets/google_sheets_credentials.json"
elif os.path.exists("google_sheets_credentials.json"):
    GSHEETS_CREDENTIALS = "google_sheets_credentials.json"
else:
    GSHEETS_CREDENTIALS = "secrets/google_sheets_credentials.json"

st.set_page_config(
    page_title=f"專案採購小幫手 {APP_VERSION}", 
    page_icon="🧪", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS 樣式
CUSTOM_CSS = """
<style>
    html, body, [class*="css"] {
        font-family: "Microsoft JhengHei", "Noto Sans TC", "PingFang TC", sans-serif;
    }
    .metric-box {
        padding: 15px; border-radius: 8px; text-align: center; color: white; margin-bottom: 10px;
    }
    .metric-title { font-size: 14px; opacity: 0.8; }
    .metric-value { font-size: 24px; font-weight: bold; }
    .project-header { font-size: 18px; font-weight: bold; color: #FF9800; }
    .item-header { font-size: 16px; font-weight: 600; color: #2196F3; margin-left: 10px; }
    .meta-info { font-size: 13px; color: #888; }
    
    div[data-baseweb="select"] > div, div[data-baseweb="base-input"] > input, div[data-baseweb="input"] > div { 
        background-color: #262730 !important; color: white !important; -webkit-text-fill-color: white !important; 
    }
    [data-testid="stDataFrame"] input[type="date"]::-webkit-calendar-picker-indicator {
        filter: invert(1); cursor: pointer;
    }
    input[type="date"]::-webkit-calendar-picker-indicator {
        filter: invert(1); cursor: pointer;
    }
</style>
"""

STATUS_OPTIONS = ["詢價中", "已報價", "待採購", "已採購", "運送中", "已到貨", "已驗收", "取消"]


# ******************************
# *--- 2. 認證與安全 ---*
# ******************************

def logout():
    st.session_state["authenticated"] = False
    st.rerun()

def login_form():
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
            st.title("🧪 測試版登入")
            st.markdown("---")
            username = st.text_input("用戶名", value=credentials["username"], disabled=True)
            password = st.text_input("密碼", type="password")
            if st.button("登入", type="primary", use_container_width=True):
                if username.strip() == credentials["username"].strip() and password == credentials["password"]:
                    st.session_state["authenticated"] = True
                    st.toast("✅ 登入成功！")
                    st.rerun()
                else:
                    st.error("密碼錯誤。")
    st.stop() 


# ******************************
# *--- 3. 外部服務 (GCS & Utils) ---*
# ******************************

@st.cache_resource
def get_gcs_signing_client():
    """獲取 GCS Client (含私鑰)，用於 Signed URL。"""
    try:
        sa_info = st.secrets["gcs_sa"]
        credentials = service_account.Credentials.from_service_account_info(sa_info)
        return storage.Client(credentials=credentials)
    except KeyError:
        st.error("GCS 憑證錯誤：secrets.toml 中缺少 [gcs_sa] 設定。")
        raise
    except Exception as e:
        st.error(f"GCS Client 載入失敗：{e}")
        raise

def get_gcs_client_standard():
    """獲取標準 GCS Client (用於一般上傳/刪除)。"""
    return storage.Client()

def upload_file_to_gcs(uploaded_file, quote_id):
    """上傳檔案至 GCS。"""
    if uploaded_file is None: return None
    try:
        client = get_gcs_client_standard()
        bucket = client.bucket(GCS_BUCKET_NAME)
        destination_blob_name = f"{GCS_FOLDER_PATH}/{quote_id}_{uploaded_file.name}"
        blob = bucket.blob(destination_blob_name)
        blob.upload_from_string(uploaded_file.getvalue(), content_type=uploaded_file.type)
        return destination_blob_name
    except Exception as e:
        logging.error(f"GCS 上傳失敗: {e}")
        st.error(f"❌ 上傳失敗：{e}")
        return None

def delete_file_from_gcs(gcs_object_name):
    """刪除 GCS 檔案。"""
    if not gcs_object_name: return True
    try:
        client = get_gcs_client_standard()
        bucket = client.bucket(GCS_BUCKET_NAME)
        blob = bucket.blob(gcs_object_name)
        if blob.exists(): blob.delete()
        return True
    except Exception as e:
        logging.error(f"GCS 刪除失敗: {e}")
        return False

def add_business_days(start_date, num_days):
    """計算工作日。"""
    current_date = start_date
    days_added = 0
    while days_added < num_days:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5: days_added += 1
    return current_date

@st.cache_data
def convert_df_to_excel(df):
    """DataFrame 轉 Excel。"""
    df_export = df.drop(columns=['標記刪除', '交期顯示', '預覽', '附件名稱'], errors='ignore') 
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='採購報價總表')
    return output.getvalue()


# ******************************
# *--- 4. 數據處理 (Google Sheets) ---*
# ******************************

def load_data_from_sheets():
    expected_cols = ['ID', '選取', '專案名稱', '專案項目', '供應商', '單價', '數量', '總價', '預計交貨日', '狀態', '採購最慢到貨日', '最後修改時間', '附件', '標記刪除']
    
    if not SHEET_URL: return pd.DataFrame(columns=expected_cols), {}

    try:
        if not GSHEETS_CREDENTIALS or not os.path.exists(GSHEETS_CREDENTIALS):
             raise FileNotFoundError("憑證檔案不存在")
             
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # 讀取 Data
        try:
            data_ws = sh.worksheet(DATA_SHEET_NAME)
            data_df = pd.DataFrame(data_ws.get_all_records())
        except:
            data_df = pd.DataFrame(columns=expected_cols)

        # 欄位補齊與清洗
        if data_df.empty: data_df = pd.DataFrame(columns=expected_cols)
        else:
            for col in expected_cols:
                if col not in data_df.columns:
                    if col in ['ID', '數量']: data_df[col] = 0
                    elif col in ['單價', '總價']: data_df[col] = 0.0
                    elif col in ['選取', '標記刪除']: data_df[col] = False
                    else: data_df[col] = ''

        def clean_bool(x):
            if isinstance(x, bool): return x
            return str(x).strip().upper() == 'TRUE'

        for col in ['選取', '標記刪除']:
            if col in data_df.columns: data_df[col] = data_df[col].apply(clean_bool)

        dtype_map = {'ID': 'Int64', '單價': 'float', '數量': 'Int64', '總價': 'float'}
        data_df = data_df.astype({k: v for k, v in dtype_map.items() if k in data_df.columns}, errors='ignore')
        
        if '附件' in data_df.columns: data_df['附件'] = data_df['附件'].astype(str)
        for col in ['預計交貨日', '採購最慢到貨日']:
            if col in data_df.columns: data_df[col] = pd.to_datetime(data_df[col], errors='coerce', format=DATE_FORMAT)

        # 讀取 Metadata
        try:
            metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
            metadata_records = metadata_ws.get_all_records()
        except:
            metadata_records = []
            
        project_metadata = {}
        for row in metadata_records:
            try: due_date = pd.to_datetime(str(row['專案交貨日'])).date()
            except: due_date = datetime.now().date()
            project_metadata[row['專案名稱']] = {
                'due_date': due_date,
                'buffer_days': int(row.get('緩衝天數', 7)),
                'last_modified': str(row.get('最後修改', ''))
            }

        st.success(f"🧪 數據載入成功！") 
        return data_df, project_metadata

    except Exception as e:
        st.error(f"❌ 數據載入失敗: {e}")
        st.session_state.data_load_failed = True
        return pd.DataFrame(columns=expected_cols), {}

def write_data_to_sheets(df_to_write, metadata_to_write):
    if st.session_state.get('data_load_failed', False) or not SHEET_URL: return False
        
    try:
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # 寫入 Data
        df_export = df_to_write.drop(columns=['交期判定', '交期顯示', '預覽', '附件名稱'], errors='ignore')
        for col in ['預計交貨日', '採購最慢到貨日']:
            if col in df_export.columns:
                df_export[col] = pd.to_datetime(df_export[col], errors='coerce').dt.strftime(DATE_FORMAT).fillna("")
        
        df_export = df_export.fillna("")
        for col in ['選取', '標記刪除']:
            if col in df_export.columns: df_export[col] = df_export[col].apply(bool)
        if '附件' in df_export.columns: df_export['附件'] = df_export['附件'].astype(str)
        
        try: data_ws = sh.worksheet(DATA_SHEET_NAME)
        except: return False

        data_ws.clear()
        data_ws.update([df_export.columns.values.tolist()] + df_export.astype(object).values.tolist())
        
        # 寫入 Metadata
        metadata_list = [
            {'專案名稱': name, 
             '專案交貨日': data['due_date'].strftime(DATE_FORMAT) if isinstance(data['due_date'], (datetime, date)) else str(data['due_date']),
             '緩衝天數': int(data['buffer_days']), 
             '最後修改': str(data['last_modified'])}
            for name, data in metadata_to_write.items()
        ]
        metadata_df = pd.DataFrame(metadata_list)
        try: metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        except: return False

        metadata_ws.clear()
        if not metadata_df.empty:
            metadata_ws.update([metadata_df.columns.values.tolist()] + metadata_df.values.tolist())
            
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ 寫入失敗: {e}")
        return False


# ******************************
# *--- 5. 邏輯處理 (Actions) ---*
# ******************************

def calculate_dashboard_metrics(df_state, project_metadata_state):
    total_projects = len(project_metadata_state)
    total_budget = 0
    df = df_state.copy()
    if df.empty: return 0, 0, 0, 0
    
    # 預算計算
    for _, proj_data in df.groupby('專案名稱'):
        if proj_data['專案名稱'].iloc[0] not in project_metadata_state: continue
        for _, item_df in proj_data.groupby('專案項目'):
            selected = item_df[item_df['選取'] == True]
            total_budget += selected['總價'].sum() if not selected.empty else item_df['總價'].min()
    
    # 風險計算
    temp = df.copy()
    temp['d'] = pd.to_datetime(temp['預計交貨日'], errors='coerce')
    temp['l'] = pd.to_datetime(temp['採購最慢到貨日'], errors='coerce')
    risk_items = (temp['d'] > temp['l']).sum()
    
    # 待處理
    pending = df[~df['狀態'].isin(['已收貨', '取消'])].shape[0]
    return total_projects, total_budget, risk_items, pending

def calculate_project_budget(df, project_name):
    proj_df = df[df['專案名稱'] == project_name]
    total = 0
    for _, item_df in proj_df.groupby('專案項目'):
        sel = item_df[item_df['選取'] == True]
        total += sel['總價'].sum() if not sel.empty else item_df['總價'].min()
    return total

def calculate_latest_arrival_dates(df, metadata):
    if df.empty or not metadata: return df
    meta_df = pd.DataFrame.from_dict(metadata, orient='index').reset_index().rename(columns={'index': '專案名稱'})
    meta_df['due_date'] = meta_df['due_date'].apply(lambda x: pd.to_datetime(x).date())
    meta_df['buffer_days'] = meta_df['buffer_days'].astype(int)
    
    df = pd.merge(df, meta_df[['專案名稱', 'due_date', 'buffer_days']], on='專案名稱', how='left')
    df['due_date_ts'] = pd.to_datetime(df['due_date'])
    df['採購最慢到貨日'] = (df['due_date_ts'] - df['buffer_days'].apply(lambda x: timedelta(days=x))).dt.strftime('%Y-%m-%d')
    return df.drop(columns=['due_date', 'buffer_days', 'due_date_ts'], errors='ignore')

def handle_master_save():
    if not st.session_state.edited_dataframes:
        st.info("無修改。")
        return
    main_df = st.session_state.data.copy()
    changed = False
    
    for _, edited_df in st.session_state.edited_dataframes.items():
        if edited_df.empty: continue
        for _, new_row in edited_df.iterrows():
            idx_in_main = main_df[main_df['ID'] == new_row['ID']].index
            if idx_in_main.empty: continue
            main_idx = idx_in_main[0]
            
            row_changed = False
            # 日期更新
            new_date = new_row['預計交貨日']
            if pd.notna(new_date):
                 new_date = pd.to_datetime(new_date).normalize()
                 if main_df.loc[main_idx, '預計交貨日'] != new_date:
                     main_df.loc[main_idx, '預計交貨日'] = new_date
                     row_changed = True
            
            # 其他欄位
            for col in ['選取', '供應商', '單價', '數量', '狀態', '標記刪除']:
                if str(main_df.loc[main_idx, col]) != str(new_row[col]):
                    main_df.loc[main_idx, col] = new_row[col]
                    row_changed = True
            
            # 總價重算
            new_total = float(main_df.loc[main_idx, '單價']) * float(main_df.loc[main_idx, '數量'])
            if main_df.loc[main_idx, '總價'] != new_total:
                main_df.loc[main_idx, '總價'] = new_total
                row_changed = True
            
            if row_changed:
                main_df.loc[main_idx, '最後修改時間'] = datetime.now().strftime(DATETIME_FORMAT)
                changed = True

    if changed:
        st.session_state.data = main_df
        if write_data_to_sheets(main_df, st.session_state.project_metadata):
            st.session_state.edited_dataframes = {}
            st.success("✅ 儲存成功！")
            st.rerun()
    else:
        st.info("無修改。")

def trigger_delete_confirmation():
    ids = []
    for edited_df in st.session_state.edited_dataframes.values():
        if edited_df is not None:
            for _, row in edited_df.iterrows():
                if row['標記刪除']: ids.append(row['ID'])
    
    if not ids:
        st.warning("請先勾選 '刪除?'。")
        return
        
    st.session_state.pending_delete_ids = ids
    st.session_state.delete_count = len(ids)
    st.session_state.show_delete_confirm = True
    st.rerun()

def handle_batch_delete_quotes():
    ids = st.session_state.get('pending_delete_ids', [])
    if not ids:
        st.session_state.show_delete_confirm = False
        st.rerun()
        return

    main_df = st.session_state.data
    quotes_to_del = main_df[main_df['ID'].isin(ids)]
    
    success = True
    for _, row in quotes_to_del.iterrows():
        if str(row.get('附件', '')).strip():
            if not delete_file_from_gcs(str(row.get('附件', '')).strip()): success = False
    
    st.session_state.data = main_df[~main_df['ID'].isin(ids)].reset_index(drop=True)
    
    if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
        st.session_state.show_delete_confirm = False
        st.session_state.edited_dataframes = {}
        msg = f"✅ 已刪除 {len(ids)} 筆。"
        if not success: msg += " (部分附件刪除失敗)"
        st.success(msg)
        st.rerun()

def handle_add_new_project():
    name = st.session_state.new_proj_name
    if not name:
        st.error("名稱不能為空")
        return
    st.session_state.project_metadata[name] = {
        'due_date': st.session_state.new_proj_due_date,
        'buffer_days': st.session_state.new_proj_buffer_days,
        'last_modified': datetime.now().strftime(DATETIME_FORMAT)
    }
    write_data_to_sheets(st.session_state.data, st.session_state.project_metadata)
    st.success(f"✅ 專案 {name} 設定已儲存")
    st.rerun()

def handle_add_new_quote(latest_arrival):
    proj = st.session_state.quote_project_select
    item = st.session_state.item_name_to_use_final
    if not proj or not item:
        st.error("請填寫專案與項目")
        return
    
    delivery = st.session_state.quote_delivery_date if st.session_state.quote_date_type == "1. 指定日期" else st.session_state.calculated_delivery_date
    
    new_row = {
        'ID': st.session_state.next_id, '選取': False, '專案名稱': proj, 
        '專案項目': item, '供應商': st.session_state.quote_supplier, 
        '單價': st.session_state.quote_price, '數量': st.session_state.quote_qty, 
        '總價': st.session_state.quote_price * st.session_state.quote_qty, 
        '預計交貨日': pd.to_datetime(delivery).normalize(), 
        '狀態': st.session_state.quote_status, 
        '採購最慢到貨日': pd.to_datetime(latest_arrival).normalize(), 
        '標記刪除': False, '最後修改時間': datetime.now().strftime(DATETIME_FORMAT), '附件': ""
    }
    st.session_state.next_id += 1
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    write_data_to_sheets(st.session_state.data, st.session_state.project_metadata)
    st.success("✅ 報價新增成功")
    st.rerun()

def handle_project_modification():
    old = st.session_state.edit_target_project
    new = st.session_state.edit_new_name
    if not new: return
    
    meta = st.session_state.project_metadata.pop(old)
    st.session_state.project_metadata[new] = meta
    st.session_state.data.loc[st.session_state.data['專案名稱'] == old, '專案名稱'] = new
    write_data_to_sheets(st.session_state.data, st.session_state.project_metadata)
    st.rerun()

def handle_delete_project(proj):
    # 刪附件
    for _, row in st.session_state.data[st.session_state.data['專案名稱'] == proj].iterrows():
        delete_file_from_gcs(str(row.get('附件', '')).strip())
    
    if proj in st.session_state.project_metadata: del st.session_state.project_metadata[proj]
    st.session_state.data = st.session_state.data[st.session_state.data['專案名稱'] != proj].reset_index(drop=True)
    write_data_to_sheets(st.session_state.data, st.session_state.project_metadata)
    st.rerun()


# ******************************
# *--- 6. UI 渲染 (Components) ---*
# ******************************

def render_sidebar_ui(df, project_metadata, today):
    with st.sidebar:
        with st.expander("✏️ 修改/刪除專案", expanded=False):
            projs = sorted(list(project_metadata.keys()))
            if projs:
                t = st.selectbox("目標專案", projs, key="edit_target_project")
                op = st.selectbox("操作", ("修改專案資訊", "刪除專案"), key="project_operation_select")
                if op == "修改專案資訊":
                    st.text_input("新名稱", value=t, key="edit_new_name")
                    if st.button("確認修改"): handle_project_modification()
                else:
                    st.warning(f"將刪除 {t} 及所有報價")
                    if st.button("確認刪除", type="secondary"): handle_delete_project(t)
            else: st.info("無專案")

        with st.expander("➕ 設定專案時程", expanded=False):
            st.text_input("專案名稱", key="new_proj_name")
            d = st.date_input("交貨日", value=today+timedelta(30), key="new_proj_due_date")
            b = st.number_input("緩衝天數", 0, value=7, key="new_proj_buffer_days")
            st.caption(f"最慢到貨: {(d - timedelta(b)).strftime(DATE_FORMAT)}")
            if st.button("儲存設定"): handle_add_new_project()

        with st.expander("➕ 新增報價", expanded=False):
            projs = sorted(list(project_metadata.keys()))
            if not projs: st.warning("請先新增專案")
            else:
                p = st.selectbox("專案", projs, key="quote_project_select")
                meta = project_metadata.get(p, {'due_date': today, 'buffer_days': 7})
                latest = meta['due_date'] - timedelta(int(meta['buffer_days']))
                st.caption(f"最慢: {latest}")
                
                items = sorted(df['專案項目'].unique().tolist())
                sel_i = st.selectbox("項目", ['🆕 新增...'] + items, key="quote_item_select")
                if sel_i == '🆕 新增...': st.text_input("新項目名稱", key="quote_item_new_input")
                else: st.session_state.item_name_to_use_final = sel_i
                
                st.text_input("供應商", key="quote_supplier")
                st.number_input("單價", 0, step=1, key="quote_price")
                st.number_input("數量", 1, value=1, key="quote_qty")
                
                dt_type = st.radio("交期", ("1. 指定日期", "2. 自然日數", "3. 工作日數"), horizontal=True, key="quote_date_type")
                if dt_type == "1. 指定日期": st.date_input("日期", today, key="quote_delivery_date")
                elif dt_type == "2. 自然日數": 
                    n = st.number_input("天數", 1, value=7, key="quote_num_days_input")
                    st.session_state.calculated_delivery_date = today + timedelta(n)
                else:
                    n = st.number_input("天數", 1, value=5, key="quote_num_days_input")
                    st.session_state.calculated_delivery_date = add_business_days(today, n)
                
                st.selectbox("狀態", STATUS_OPTIONS, key="quote_status")
                if st.button("新增資料", type="primary"): handle_add_new_quote(latest)

        st.button("🚪 登出", on_click=logout, type="secondary")

def render_dashboard(df, project_metadata):
    tp, tb, ri, pq = calculate_dashboard_metrics(df, project_metadata)
    st.subheader("📊 總覽儀表板")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='metric-box' style='background:#33343c;'><div class='metric-title'>專案數</div><div class='metric-value'>{tp}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='metric-box' style='background:#1b4d3e;'><div class='metric-title'>總預算</div><div class='metric-value'>${tb:,.0f}</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='metric-box' style='background:#5a2a2a;'><div class='metric-title'>風險項</div><div class='metric-value'>{ri}</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='metric-box' style='background:#2a3b5a;'><div class='metric-title'>待處理</div><div class='metric-value'>{pq}</div></div>", unsafe_allow_html=True)
    st.markdown("---")

def render_batch_operations():
    c1, c2 = st.columns([0.8, 0.2])
    locked = st.session_state.show_delete_confirm
    if c1.button("💾 儲存修改", type="primary", disabled=locked): handle_master_save()
    if c2.button("🔴 刪除標記", type="secondary", disabled=locked): trigger_delete_confirmation()
    
    if st.session_state.show_delete_confirm:
        st.error(f"確認刪除 {st.session_state.delete_count} 筆？")
        c_y, c_n = st.columns([0.2, 0.8])
        if c_y.button("✅ 確認"): handle_batch_delete_quotes()
        if c_n.button("❌ 取消"): 
            st.session_state.show_delete_confirm = False
            st.rerun()
    st.markdown("---")

def render_project_tables(df, project_metadata):
    if df.empty:
        st.info("無數據")
        return

    # 初始化預覽 ID
    if 'preview_from_table_id' not in st.session_state:
        st.session_state.preview_from_table_id = None
        
    current_preview_id = st.session_state.preview_from_table_id

    for proj_name, proj_data in df.groupby('專案名稱'):
        meta = project_metadata.get(proj_name, {})
        budget = calculate_project_budget(df, proj_name)
        
        try: dd = meta.get('due_date').strftime(DATE_FORMAT)
        except: dd = str(meta.get('due_date'))
        
        try: ld = (meta.get('due_date') - timedelta(int(meta.get('buffer_days', 7)))).strftime(DATE_FORMAT)
        except: ld = "N/A"

        with st.expander(f"專案：{proj_name}", expanded=False):
            st.markdown(f"<span class='project-header'>預算: ${budget:,.0f} | 交期: {dd} | 最慢: {ld}</span>", unsafe_allow_html=True)
            st.caption("💡 提示：勾選 **「預覽」** 欄位可查看附件 (單選)。")

            for item_name, item_data in proj_data.groupby('專案項目'):
                st.markdown(f"<span class='item-header'>📦 {item_name}</span>", unsafe_allow_html=True)
                
                edf = item_data.copy()
                if '預計交貨日' in edf.columns:
                    edf['預計交貨日'] = pd.to_datetime(edf['預計交貨日'], errors='coerce').apply(lambda x: x.date() if pd.notnull(x) else None)
                if '採購最慢到貨日' in edf.columns:
                    edf['採購最慢到貨日'] = pd.to_datetime(edf['採購最慢到貨日'], errors='coerce').apply(lambda x: x.date() if pd.notnull(x) else None)
                
                if '最後修改時間' not in edf: edf['最後修改時間'] = ''
                
                # 附件顯示處理
                edf['附件名稱'] = edf['附件'].apply(lambda x: os.path.basename(x) if x else '')
                
                # *** 單選核心邏輯：根據 State 設定 Checkbox ***
                edf['預覽'] = edf['ID'].apply(lambda x: True if x == current_preview_id else False)

                col_cfg = {
                    "選取": st.column_config.CheckboxColumn(width="tiny"),
                    "單價": st.column_config.NumberColumn(format="$%d"),
                    "總價": st.column_config.NumberColumn(format="$%d", disabled=True),
                    "預計交貨日": st.column_config.DateColumn(format="YYYY-MM-DD", step=1),
                    "狀態": st.column_config.SelectboxColumn(options=STATUS_OPTIONS),
                    "最後修改時間": st.column_config.TextColumn(disabled=True),
                    "附件名稱": st.column_config.TextColumn(disabled=True, width="medium"),
                    "預覽": st.column_config.CheckboxColumn(width="small", label="預覽(單選)"),
                    "標記刪除": st.column_config.CheckboxColumn(width="tiny")
                }
                
                editor_key = f"editor_{proj_name}_{item_name}"
                edited_val = st.data_editor(
                    edf, 
                    column_order=['選取', '供應商', '單價', '數量', '總價', '預計交貨日', '交期判定', '狀態', '最後修改時間', '附件名稱', '預覽', '標記刪除'],
                    column_config=col_cfg,
                    key=editor_key,
                    hide_index=True,
                    use_container_width=True,
                    disabled=st.session_state.show_delete_confirm
                )
                st.session_state.edited_dataframes[item_name] = edited_val

                # *** 單選核心邏輯：偵測點擊並重整 ***
                if '預覽' in edited_val.columns:
                    checked = edited_val[edited_val['預覽'] == True]
                    if not checked.empty:
                        for _, row in checked.iterrows():
                            # 如果這個 ID 與當前不同，代表是新點擊的 -> 更新 State 並重整
                            if row['ID'] != current_preview_id:
                                st.session_state.preview_from_table_id = row['ID']
                                st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("💾 資料匯出")
    st.download_button("📥 下載 Excel", convert_df_to_excel(df), f'report_{datetime.now().strftime("%Y%m%d")}.xlsx')


def render_attachment_module(df):
    """
    【UI 優化版】附件管理中心
    使用卡片式佈局，移除複雜下拉選單，以表格點選為主。
    """
    st.markdown("---")
    st.subheader("📎 報價附件管理中心")

    # 1. 獲取當前選取的報價 ID
    selected_id = st.session_state.get('preview_from_table_id', None)
    
    # 2. 如果沒有選取，顯示提示畫面
    if selected_id is None:
        st.info("👆 請在上方表格中勾選 **「預覽」** 欄位以檢視或上傳附件。")
        return

    # 3. 獲取該筆資料
    try:
        row = df[df['ID'] == selected_id].iloc[0]
        proj_name = row['專案名稱']
        item_name = row['專案項目']
        supplier = row['供應商']
        current_file_path = str(row.get('附件', '')).strip()
        current_filename = os.path.basename(current_file_path) if current_file_path else None
    except IndexError:
        st.error("找不到該筆資料，可能已被刪除。")
        st.session_state.preview_from_table_id = None
        st.rerun()
        return

    # 4. UI 呈現：使用容器框住
    with st.container(border=True):
        # 標題列
        col_header, col_close = st.columns([0.9, 0.1])
        with col_header:
            st.markdown(f"### 📦 {item_name} <span style='font-size:0.8em; color:gray'>({supplier})</span>", unsafe_allow_html=True)
            st.caption(f"專案：{proj_name} | ID: {selected_id}")
        with col_close:
            if st.button("❌", help="關閉預覽"):
                st.session_state.preview_from_table_id = None
                st.rerun()

        st.markdown("---")

        # 內容區：左側操作，右側預覽
        col_action, col_preview = st.columns([1, 1.5], gap="large")

        with col_action:
            st.markdown("#### 📤 附件操作")
            
            # 狀態顯示
            if current_filename:
                st.success(f"✅ 現有附件：**{current_filename}**")
            else:
                st.warning("⚠️ 目前尚無附件")

            # 上傳區
            uploaded_file = st.file_uploader("上傳/更換附件 (JPG, PNG, PDF)", type=['png', 'jpg', 'jpeg', 'pdf'], key=f"uploader_{selected_id}")
            
            if uploaded_file:
                if st.button("☁️ 確認上傳至 GCS", type="primary", use_container_width=True):
                    new_path = save_uploaded_file(uploaded_file, selected_id)
                    if new_path:
                        # 更新 Session Data
                        idx = st.session_state.data[st.session_state.data['ID'] == selected_id].index[0]
                        st.session_state.data.loc[idx, '附件'] = new_path
                        st.session_state.data.loc[idx, '最後修改時間'] = datetime.now().strftime(DATETIME_FORMAT)
                        
                        # 寫入 Sheets
                        if write_data_to_sheets(st.session_state.data, st.session_state.project_metadata):
                            st.toast("上傳成功！")
                            time.sleep(1)
                            st.rerun()

        with col_preview:
            st.markdown("#### 👁️ 內容預覽")
            
            if current_file_path:
                try:
                    # 使用 Signed URL 獲取安全連結
                    client = get_gcs_signing_client()
                    bucket = client.bucket(GCS_BUCKET_NAME)
                    blob = bucket.blob(current_file_path)
                    signed_url = blob.generate_signed_url(version="v4", expiration=timedelta(minutes=10), method="GET")
                    
                    ext = os.path.splitext(current_filename)[1].lower()
                    if ext in ['.png', '.jpg', '.jpeg']:
                        st.image(signed_url, caption=current_filename, use_container_width=True)
                    elif ext == '.pdf':
                        st.markdown(f'<iframe src="{signed_url}" width="100%" height="600"></iframe>', unsafe_allow_html=True)
                    else:
                        st.markdown(f"📄 [下載檔案]({signed_url})")
                        
                except Exception as e:
                    st.error(f"預覽失敗: {e}")
            else:
                st.info("無檔案可預覽")


# ******************************
# *--- 7. 主程式邏輯 ---*
# ******************************

def run_app():
    # 初始化
    if 'data' not in st.session_state:
        df, meta = load_data_from_sheets()
        st.session_state.data = df
        st.session_state.project_metadata = meta
        st.session_state.next_id = df['ID'].max() + 1 if not df.empty else 1
        st.session_state.edited_dataframes = {}
        st.session_state.show_delete_confirm = False
        st.session_state.preview_from_table_id = None

    # 自動計算
    st.session_state.data = calculate_latest_arrival_dates(st.session_state.data, st.session_state.project_metadata)
    
    # 判斷交期
    def judge(row):
        try:
            d = pd.to_datetime(row['預計交貨日'])
            l = pd.to_datetime(row['採購最慢到貨日'])
            if pd.isna(d) or pd.isna(l): return ""
            return "❌" if d.date() > l.date() else "✅"
        except: return ""
    
    if not st.session_state.data.empty:
        st.session_state.data['交期判定'] = st.session_state.data.apply(judge, axis=1)

    # 渲染畫面
    df = st.session_state.data
    meta = st.session_state.project_metadata
    today = datetime.now().date()
    
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}") 
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    
    render_sidebar_ui(df, meta, today)
    render_dashboard(df, meta)
    render_batch_operations()
    render_project_tables(df, meta)
    render_attachment_module(df) # 新版 UI

def main():
    login_form()
    if st.session_state.get("authenticated", False):
        run_app()

if __name__ == "__main__":
    main()
