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
APP_VERSION = "v2.2.2 (Hyperlink & GCS)" # <--- 版本號更新
STATUS_OPTIONS = ["待採購", "已下單", "已收貨", "取消"]

# --- Google Cloud Storage 配置 ---
# ⚠️ WARNING: 請替換為您在 GCP 上建立的儲存桶名稱！
GCS_BUCKET_NAME = "procurement-attachments-bucket"
GCS_ATTACHMENT_FOLDER = "attachments"

# --- 數據源配置 ---
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

# --- CSS ---
CUSTOM_CSS = """
<style>
.streamlit-expanderContent { padding: 1rem !important; }
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

# --- 登入函式 ---
def logout():
    st.session_state["authenticated"] = False
    st.rerun()

def login_form():
    DEFAULT_USERNAME = os.environ.get("AUTH_USERNAME", "dev_user")
    DEFAULT_PASSWORD = os.environ.get("AUTH_PASSWORD", "dev_pwd")
    credentials = {"username": DEFAULT_USERNAME, "password": DEFAULT_PASSWORD}

    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
        
    if st.session_state["authenticated"]:
        return

    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.container(border=True):
            st.title("🔐 請登入")
            username = st.text_input("用戶名", value=credentials["username"], disabled=True)
            password = st.text_input("密碼", type="password")
            if st.button("登入", type="primary"):
                if username.strip() == credentials["username"].strip() and password == credentials["password"]:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("密碼錯誤")
    st.stop() 


# --- GCS 服務函式 (強化版) ---

def upload_attachment_to_gcs(file_obj, next_id):
    try:
        storage_client = storage.Client()
        bucket = storage_client.bucket(GCS_BUCKET_NAME)
        file_extension = os.path.splitext(file_obj.name)[1]
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        blob_name = f"{GCS_ATTACHMENT_FOLDER}/{next_id}_{timestamp}{file_extension}"
        blob = bucket.blob(blob_name)
        file_obj.seek(0)
        blob.upload_from_file(file_obj, content_type=file_obj.type)
        return f"gs://{GCS_BUCKET_NAME}/{blob_name}"
    except Exception as e:
        logging.error(f"GCS Upload Error: {e}")
        st.error("❌ 附件上傳失敗，請檢查 GCP 權限。")
        return None

def generate_signed_url_cached(gcs_uri):
    """生成簽章 URL，用於在表格中直接顯示超連結。"""
    if not gcs_uri or not isinstance(gcs_uri, str):
        return None
    if gcs_uri.startswith("http"):
        return gcs_uri
    if not gcs_uri.startswith("gs://"):
        return None

    try:
        # 解析 gs://bucket/path
        parts = gcs_uri[5:].split('/', 1)
        bucket_name = parts[0]
        blob_name = parts[1]
        
        storage_client = storage.Client()
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        
        # 簽署 URL (有效期 60 分鐘)
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(minutes=60),
            method="GET"
        )
        return url
    except Exception as e:
        # 如果權限不足 (缺少 Token Creator)，這裡會報錯
        # 為了不讓整個 App 崩潰，我們返回 None 並記錄日誌
        logging.error(f"Failed to sign URL for {gcs_uri}: {e}")
        return None


# --- Gspread 函式 ---

@st.cache_data(ttl=600, show_spinner="連線 Sheets...")
def load_data_from_sheets():
    if not SHEET_URL:
        return pd.DataFrame(), {}

    try:
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_df = pd.DataFrame(data_ws.get_all_records())

        if '附件URL' not in data_df.columns:
            data_df['附件URL'] = ""
            
        data_df = data_df.astype({'ID': 'Int64', '選取': 'bool', '單價': 'float', '數量': 'Int64', '總價': 'float'})
        if '標記刪除' not in data_df.columns: data_df['標記刪除'] = False

        metadata_ws = sh.worksheet(METADATA_SHEET_NAME)
        meta_records = metadata_ws.get_all_records()
        project_metadata = {}
        for row in meta_records:
            try: due = pd.to_datetime(str(row['專案交貨日'])).date()
            except: due = datetime.now().date()
            project_metadata[row['專案名稱']] = {'due_date': due, 'buffer_days': int(row['緩衝天數']), 'last_modified': str(row['最後修改'])}

        return data_df, project_metadata
    except Exception as e:
        st.error(f"數據載入失敗: {e}")
        st.session_state.data_load_failed = True
        return pd.DataFrame(), {}


def write_data_to_sheets(df_to_write, metadata_to_write):
    if st.session_state.get('data_load_failed', False) or not SHEET_URL: return False
    try:
        gc = gspread.service_account(filename=GSHEETS_CREDENTIALS)
        sh = gc.open_by_url(SHEET_URL)
        
        # ⚠️ 重要：移除 '附件連結' 等輔助欄位，只儲存原始數據
        cols_to_drop = ['標記刪除', '交期顯示', '附件連結'] 
        df_export = df_to_write.drop(columns=cols_to_drop, errors='ignore')
        
        data_ws = sh.worksheet(DATA_SHEET_NAME)
        data_ws.clear()
        data_ws.update([df_export.columns.values.tolist()] + df_export.values.tolist())
        
        meta_list = [{'專案名稱': k, '專案交貨日': v['due_date'].strftime('%Y-%m-%d'), '緩衝天數': v['buffer_days'], '最後修改': v['last_modified']} for k,v in metadata_to_write.items()]
        meta_df = pd.DataFrame(meta_list)
        meta_ws = sh.worksheet(METADATA_SHEET_NAME)
        meta_ws.clear()
        meta_ws.update([meta_df.columns.values.tolist()] + meta_df.values.tolist())
        
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"寫入失敗: {e}")
        return False

# --- 輔助函式 ---
def calculate_latest_arrival_dates(df, metadata):
    if df.empty or not metadata: return df
    meta_df = pd.DataFrame.from_dict(metadata, orient='index').reset_index().rename(columns={'index': '專案名稱'})
    meta_df['due_date'] = pd.to_datetime(meta_df['due_date']).dt.date
    df = pd.merge(df, meta_df[['專案名稱', 'due_date', 'buffer_days']], on='專案名稱', how='left')
    df['採購最慢到貨日'] = (pd.to_datetime(df['due_date']) - pd.to_timedelta(df['buffer_days'], unit='D')).dt.strftime('%Y-%m-%d')
    return df.drop(columns=['due_date', 'buffer_days'], errors='ignore')

def calculate_metrics(df, meta):
    if df.empty: return 0, 0, 0, 0
    total_budget = 0
    for _, proj in df.groupby('專案名稱'):
        for _, item in proj.groupby('專案項目'):
            sel = item[item['選取'] == True]
            total_budget += sel['總價'].sum() if not sel.empty else item['總價'].min()
    risk = (pd.to_datetime(df['預計交貨日'], errors='coerce') > pd.to_datetime(df['採購最慢到貨日'], errors='coerce')).sum()
    pending = df[~df['狀態'].isin(['已收貨', '取消'])].shape[0]
    return len(meta), total_budget, risk, pending

def save_and_rerun(df, meta, msg=""):
    if write_data_to_sheets(df, meta):
        st.session_state.edited_dataframes = {}
        if msg: st.success(msg)
        st.rerun()

def handle_master_save():
    if not st.session_state.edited_dataframes: return
    main_df = st.session_state.data
    changed = False
    
    for _, edited_df in st.session_state.edited_dataframes.items():
        if edited_df.empty: continue
        for _, new_row in edited_df.iterrows():
            idx = main_df[main_df['ID'] == new_row['ID']].index
            if idx.empty: continue
            idx = idx[0]
            
            # 僅更新可編輯欄位，忽略 '附件連結'
            for col in ['選取', '供應商', '單價', '數量', '狀態', '標記刪除', '附件URL']:
                if col in new_row and main_df.at[idx, col] != new_row[col]:
                    main_df.at[idx, col] = new_row[col]
                    changed = True
            
            # 日期與總價邏輯
            try:
                date_val = str(new_row['交期顯示']).split(' ')[0]
                if main_df.at[idx, '預計交貨日'] != date_val:
                    main_df.at[idx, '預計交貨日'] = date_val
                    changed = True
            except: pass
            
            new_total = float(new_row['單價']) * float(new_row['數量'])
            if main_df.at[idx, '總價'] != new_total:
                main_df.at[idx, '總價'] = new_total
                changed = True

    if changed:
        st.session_state.data = main_df.copy()
        save_and_rerun(st.session_state.data, st.session_state.project_metadata, "✅ 儲存成功！")
    else:
        st.info("無變更")

def handle_add_quote(date, file):
    proj = st.session_state.quote_project
    item = st.session_state.quote_item_final
    if not proj or not item:
        st.error("請填寫完整資訊")
        return
        
    uri = ""
    if file:
        with st.spinner("上傳附件中..."):
            uri = upload_attachment_to_gcs(file, st.session_state.next_id)
            if not uri: return
            
    new_row = {
        'ID': st.session_state.next_id, '選取': False, '專案名稱': proj, '專案項目': item,
        '供應商': st.session_state.quote_supplier, '單價': st.session_state.quote_price,
        '數量': st.session_state.quote_qty, '總價': st.session_state.quote_price * st.session_state.quote_qty,
        '預計交貨日': st.session_state.quote_date.strftime('%Y-%m-%d'), '狀態': '待採購',
        '採購最慢到貨日': date.strftime('%Y-%m-%d'), '標記刪除': False, '附件URL': uri
    }
    st.session_state.next_id += 1
    st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
    save_and_rerun(st.session_state.data, st.session_state.project_metadata, "✅ 新增成功！")

# --- 初始化 ---
def init_state():
    if 'data' not in st.session_state:
        d, m = load_data_from_sheets()
        st.session_state.data = d
        st.session_state.project_metadata = m
    
    defaults = {'next_id': 1, 'edited_dataframes': {}}
    if not st.session_state.data.empty:
        defaults['next_id'] = st.session_state.data['ID'].max() + 1
        
    for k, v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v

# --- App ---
def run_app():
    st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    init_state()
    
    st.session_state.data = calculate_latest_arrival_dates(st.session_state.data, st.session_state.project_metadata)
    
    # 日期顯示處理
    def fmt_date(r):
        d = str(r['預計交貨日'])
        try: 
            return f"{d} 🔴" if pd.to_datetime(d).date() > pd.to_datetime(r['採購最慢到貨日']).date() else f"{d} ✅"
        except: return d
    
    if not st.session_state.data.empty:
        st.session_state.data['交期顯示'] = st.session_state.data.apply(fmt_date, axis=1)

    # 側邊欄
    with st.sidebar:
        st.button("登出", on_click=logout)
        # (簡化側邊欄邏輯以聚焦核心功能，保留新增報價)
        with st.expander("➕ 新增報價", expanded=True):
            projs = sorted(st.session_state.project_metadata.keys())
            if not projs: st.warning("請先設定專案")
            else:
                st.session_state.quote_project = st.selectbox("專案", projs)
                meta = st.session_state.project_metadata[st.session_state.quote_project]
                limit_date = meta['due_date'] - timedelta(days=meta['buffer_days'])
                st.caption(f"最慢到貨: {limit_date}")
                
                exist_items = sorted(st.session_state.data['專案項目'].unique())
                sel_item = st.selectbox("項目", ["新項目..."] + exist_items)
                st.session_state.quote_item_final = st.text_input("項目名稱") if sel_item == "新項目..." else sel_item
                
                st.session_state.quote_supplier = st.text_input("供應商")
                st.session_state.quote_price = st.number_input("單價", 0)
                st.session_state.quote_qty = st.number_input("數量", 1, value=1)
                st.session_state.quote_date = st.date_input("預計交貨")
                
                f = st.file_uploader("附件")
                if st.button("新增"): handle_add_quote(limit_date, f)

    # 儀表板
    n_proj, bud, risk, pend = calculate_metrics(st.session_state.data, st.session_state.project_metadata)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("專案數", n_proj)
    c2.metric("總預算", f"${bud:,.0f}")
    c3.metric("風險項", risk)
    c4.metric("待處理", pend)
    
    st.button("💾 儲存所有變更", type="primary", on_click=handle_master_save)

    # 專案列表 (核心：附件連結轉換)
    for proj_name, proj_df in st.session_state.data.groupby('專案名稱'):
        with st.expander(f"專案: {proj_name}", expanded=True):
            for item_name, item_df in proj_df.groupby('專案項目'):
                st.markdown(f"**{item_name}**")
                
                # --- [關鍵修改] 動態生成簽章連結 ---
                display_df = item_df.copy()
                
                # 1. 建立 '附件連結' 欄位，預設為空
                display_df['附件連結'] = None 
                
                # 2. 針對有 gs:// 的列生成簽章
                for idx, row in display_df.iterrows():
                    uri = row.get('附件URL', '')
                    if uri:
                        signed_url = generate_signed_url_cached(uri)
                        if signed_url:
                            # 成功簽署，設為 URL
                            display_df.at[idx, '附件連結'] = signed_url
                        else:
                            # 簽署失敗 (通常是權限問題)，設為 None 或錯誤提示
                            # LinkColumn 如果是 None 就不會顯示連結
                            pass

                editor_key = f"ed_{proj_name}_{item_name}"
                edited = st.data_editor(
                    display_df[[
                        'ID', '選取', '供應商', '單價', '數量', '總價', 
                        '交期顯示', '狀態', '附件連結', '附件URL', '標記刪除' # 包含新舊欄位
                    ]],
                    column_config={
                        "ID": st.column_config.NumberColumn(disabled=True, width="small"),
                        "附件連結": st.column_config.LinkColumn(
                            "附件 (點擊開啟)", 
                            display_text="📄 開啟附件", 
                            help="點擊即可在新分頁開啟附件。若無法開啟，請檢查 GCP 權限。",
                            width="medium"
                        ),
                        "附件URL": st.column_config.TextColumn(
                            "原始路徑 (gs://)", 
                            disabled=True, 
                            help="系統內部儲存路徑，不可編輯"
                        ),
                        "交期顯示": st.column_config.TextColumn("交貨日", disabled=False),
                        "總價": st.column_config.NumberColumn(disabled=True),
                    },
                    hide_index=True,
                    key=editor_key
                )
                st.session_state.edited_dataframes[item_name] = edited

def main():
    login_form()
    if st.session_state.authenticated: run_app()

if __name__ == "__main__":
    main()
