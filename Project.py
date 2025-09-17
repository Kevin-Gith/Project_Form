import streamlit as st
import pandas as pd
import datetime
import gspread
import json
from google.oauth2.service_account import Credentials
from io import BytesIO

# ------------ CONFIG -------------
SHEET_NAME = "Project_Form"       # Google Sheet 名稱
WORKSHEET_NAME = "Python"         # 工作表名稱

USERS = {
    "sam@kipotec.com.tw":   {"password": "Kipo-0926969586$$$", "name": "Sam"},
    "sale1@kipotec.com.tw": {"password": "Kipo-0917369466$$$", "name": "Vivian"},
    "sale2@kipotec.com.tw": {"password": "Kipo-0905038111$$$", "name": "Lillian"},
    "sale5@kipotec.com.tw": {"password": "Kipo-0925698417$$$", "name": "Wendy"},
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ---------- Google Sheet ----------
def get_gc():
    # 使用本地的 service_account_key.json
    with open("service_account_key.json", "r") as f:
        service_account_info = json.load(f)
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    return gspread.authorize(creds)

def open_main_ws():
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=2000, cols=20)
        ws.append_row([
            "ODM_Customers","Brand_Customers","Application_Purpose",
            "Product_Application","Cooling_Solution","Delivery_Location",
            "Applicant","Application Deadline"
        ])
    return ws

def append_row(ws, row_dict):
    headers = ws.row_values(1)
    row = [str(row_dict.get(h, "")) for h in headers]
    ws.append_row(row, value_input_option="RAW")

# ---------- UI ----------
st.set_page_config(page_title="專案申請系統", layout="centered")

# 初始化 session_state
for key, default in [("logged_in", False), ("username", ""), ("page", "login"), ("form_data", {})]:
    if key not in st.session_state:
        st.session_state[key] = default

# ---------- 登入頁 ----------
def login():
    st.title("專案申請系統 - 登入")
    username = st.text_input("帳號 (email)")
    password = st.text_input("密碼", type="password")
    if st.button("登入"):
        if username in USERS and USERS[username]["password"] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.page = "form"
            st.rerun()
        else:
            st.error("帳號或密碼輸入錯誤，請重新輸入")

# ---------- 表單頁 ----------
def form_page():
    st.title("填寫專案表單")
    applicant = USERS[st.session_state.username]["name"]

    st.header("A. 客戶資訊")
    odm = st.selectbox("ODM客戶(RD)", ["", "(01)仁寶","(02)廣達","(03)緯創","(04)華勤","(05)光寶","(06)技嘉","(07)智邦","(00)其他"])
    brand = st.selectbox("品牌客戶(RD)", ["", "(01)惠普","(02)聯想","(03)高通","(04)華碩","(05)宏碁","(06)微星","(07)技嘉","(00)其他"])
    purpose = st.selectbox("申請目的", ["", "(01)客戶專案開發","(02)內部新產品開發","(03)技術平台預研","(04)其他"])

    st.header("B. 開案資訊")
    product = st.selectbox("產品應用", ["", "(01)NB CPU","(02)NB GPU","(03)Server","(04)Automotive(Car)","(05)Other"])
    cooling = st.selectbox("散熱方式", ["", "(01)Air Cooling","(02)Fan","(03)Cooler(含Fan)","(04)Liquid Cooling","(05)Other"])
    location = st.selectbox("交貨地點", ["", "(01)Taiwan","(02)China","(03)Thailand","(04)Vietnam","(05)Other"])

    # ----------------- C. 規格資訊 -----------------
    st.header("C. 規格資訊 (可複選，不會存到 Google)")
    spec_options = ["Air Cooling", "Fan", "Liquid Cooling"]
    selected_specs = st.multiselect("請選擇規格類型", spec_options)

    spec_details = {}
    for spec in selected_specs:
        st.subheader(f"{spec} 規格")
        detail = st.text_area(f"請輸入 {spec} 的詳細說明")
        spec_details[spec] = detail
    # ------------------------------------------------

    if st.button("完成填寫"):
        if not odm or not brand or not purpose or not product or not cooling or not location:
            st.error("❌ 客戶資訊或開案資訊欄位未填寫完畢，請重新確認")
            return

        st.session_state.form_data = {
            "ODM_Customers": odm,
            "Brand_Customers": brand,
            "Application_Purpose": purpose,
            "Product_Application": product,
            "Cooling_Solution": cooling,
            "Delivery_Location": location,
            "Applicant": applicant,
            "Application Deadline": datetime.datetime.now().strftime("%Y/%m/%d %H:%M"),
            # 規格資訊只給 preview/excel 用，不送 Google
            "Spec_Info": json.dumps(spec_details, ensure_ascii=False),
        }
        st.session_state.page = "preview"
        st.rerun()

# ---------- 預覽頁 ----------
def preview_page():
    st.title("表單預覽")

    # 拿到資料
    form_data = st.session_state.form_data.copy()

    # 展開規格資訊
    spec_details = json.loads(form_data.get("Spec_Info", "{}"))
    if spec_details:
        for spec, detail in spec_details.items():
            form_data[f"Spec_{spec}"] = detail

    # 預覽表格 (含規格)
    st.table(pd.DataFrame([form_data]))

    # 匯出 Excel 也包含規格
    buffer = BytesIO()
    pd.DataFrame([form_data]).to_excel(buffer, index=False)
    st.download_button("下載 Excel", buffer.getvalue(), file_name="project_application.xlsx")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("送出"):
            ws = open_main_ws()
            # ✅ 送 Google Sheet 前刪掉 Spec_Info
            google_data = st.session_state.form_data.copy()
            google_data.pop("Spec_Info", None)
            append_row(ws, google_data)
            st.success("✅ 已送出並記錄到 Google Sheet")
            st.session_state.page = "form"
            st.rerun()
    with col2:
        if st.button("取消"):
            st.session_state.page = "form"
            st.rerun()

# ---------- 主程式 ----------
def main():
    if not st.session_state.logged_in:
        login()
    elif st.session_state.page == "form":
        form_page()
    elif st.session_state.page == "preview":
        preview_page()

if __name__ == "__main__":
    main()