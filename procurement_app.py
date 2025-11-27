# --- 移除所有 streamlit_authenticator, yaml, config 導入 ---
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import os 
import json
import gspread # <-- 保持 Gspread 連線導入

# 移除 logging, yaml, streamlit_authenticator 導入
# 確保你的程式碼沒有這些舊的/衝突的導入。

# --- 將 V1.0.0 登入函式和 logout 函式貼到這裡 ---

def logout():
    """登出函式：清除驗證狀態並重新運行。"""
    st.session_state.authenticated = False
    st.rerun()

def login_form():
    """渲染登入表單並處理密碼驗證。"""
    
    # 設置預設的用戶名和密碼
    DEFAULT_USERNAME = "tajung"
    DEFAULT_PASSWORD = "tjdfb24676881"

    # 這裡我們使用硬編碼或 secrets（如果存在）來檢查密碼
    try:
        credentials = st.secrets["auth"]
    except (KeyError, FileNotFoundError):
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
            
            # 注意：這裡使用硬編碼的用戶名 "tajung"，確保與你的 V1.0.0 程式碼一致
            username = st.text_input("用戶名", key="login_username", value=credentials["username"]) # 這裡不再允許用戶修改用戶名，只允許輸入密碼
            password = st.text_input("密碼", type="password", key="login_password")
            
            if st.button("登入", type="primary"):
                # 注意：這裡的比較應該是基於用戶輸入的密碼
                if username.strip() == credentials["username"].strip() and password == credentials["password"]:
                    st.session_state["authenticated"] = True
                    st.toast("✅ 登入成功！")
                    st.rerun()
                else:
                    st.error("用戶名或密碼錯誤。")
            
    # 如果未驗證，阻止執行後續程式碼
    st.stop() 


# --- 將 V2.1.3 的所有核心邏輯包裝在 if 區塊內 ---

def main():
    # 執行登入驗證
    login_form()
    
    # --- 僅在驗證通過後執行後續程式碼 ---
    if st.session_state.authenticated:
        # 顯示登出按鈕
        st.sidebar.button("登出", on_click=logout) 

        # --- 以下是原 V2.1.3 的所有核心邏輯 ---
        
        initialize_session_state()
        
        # 執行 V2.1.3 的數據自動計算邏輯
        st.session_state.data = calculate_latest_arrival_dates(
            st.session_state.data, 
            st.session_state.project_metadata
        )
        
        # 確保所有 V2.1.3 的 UI 邏輯在此處運行
        # 這裡應該是原 run_app() 的內容，但我們將其直接整合到 main()
        # ... (V2.1.3 的所有 UI 邏輯，從 st.title 開始) ...
        
        # 由於 V1.0.0 中沒有 run_app()，我們將 V2.1.3 的 UI 邏輯直接放在這裡
        
        st.title(f"🛠️ 專案採購管理工具 {APP_VERSION}")
        st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

        # ... (V2.1.3 的儀表板、側邊欄、Expander、data_editor 等所有邏輯) ...
        
# ------------------------------------------------------------------

if __name__ == "__main__":
    main()
