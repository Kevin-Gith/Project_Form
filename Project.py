import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
import datetime

# ========== Google Sheet 設定 ==========
SHEET_NAME = "Project_Form"
WORKSHEET_NAME = "Python"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES
)
client = gspread.authorize(creds)
sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)

# ========== 使用者帳號密碼 ==========
USER_CREDENTIALS = {
    "sam@kipotec.com.tw": "Kipo-0926969586$$$",
    "sale1@kipotec.com.tw": "Kipo-0917369466$$$",
    "sale5@kipotec.com.tw": "Kipo-0925698417$$$",
    "sale2@kipotec.com.tw": "Kipo-0905038111$$$"
}

USER_MAPPING = {
    "sam@kipotec.com.tw": "Sam",
    "sale1@kipotec.com.tw": "Vivian",
    "sale5@kipotec.com.tw": "Wendy",
    "sale2@kipotec.com.tw": "Lillian"
}

# ========== 登入頁 ==========
def login_page():
    st.title("💻 Kipo專案申請系統")

    username = st.text_input("帳號")
    password = st.text_input("密碼", type="password")

    if st.button("登入"):
        if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
            st.session_state["logged_in"] = True
            st.session_state["user"] = username
            st.rerun()
        else:
            st.error("帳號或密碼錯誤，請重新輸入！")

# ========== A. 客戶資訊 ==========
def render_customer_info():
    st.header("A. 客戶資訊")

    current_user = st.session_state.get("user", "")
    sales_user = USER_MAPPING.get(current_user, "Unknown")
    st.text_input("北辦業務", value=sales_user, disabled=True)

    odm = st.selectbox("ODM客戶 (RD)", [
        "(01)仁寶", "(02)廣達", "(03)緯創", "(04)華勤",
        "(05)光寶", "(06)技嘉", "(07)智邦", "(08)其他"
    ])
    if odm.endswith("其他"):
        odm = st.text_input("請輸入 ODM 客戶")

    brand = st.selectbox("品牌客戶 (RD)", [
        "(01)惠普", "(02)聯想", "(03)高通", "(04)華碩",
        "(05)宏碁", "(06)微星", "(07)技嘉", "(08)其他"
    ])
    if brand.endswith("其他"):
        brand = st.text_input("請輸入品牌客戶")

    purpose = st.selectbox("申請目的", [
        "(01)客戶專案開發", "(02)內部新產品開發", "(03)技術平台預研", "(04)其他"
    ])
    if purpose.endswith("其他"):
        purpose = st.text_input("請輸入申請目的")

    project_name = st.text_input("客戶專案名稱")
    proposal_date = st.date_input("客戶提案日期", value=datetime.date.today())

    return {
        "Sales_User": sales_user,
        "ODM_Customers": odm,
        "Brand_Customers": brand,
        "Application_Purpose": purpose,
        "Project_Name": project_name,
        "Proposal_Date": proposal_date.strftime("%Y/%m/%d")
    }

# ========== B. 開案資訊 ==========
def render_project_info():
    st.header("B. 開案資訊")

    product_app = st.selectbox("產品應用", [
        "(01)NB CPU", "(02)NB GPU", "(03)Server",
        "(04)Automotive(Car)", "(05)Other"
    ])
    if product_app.endswith("Other"):
        product_app = st.text_input("請輸入產品應用")

    cooling = st.selectbox("散熱方式", [
        "(01)Air Cooling氣冷", "(02)Fan風扇", "(03)Cooler(含Fan)",
        "(04)Liquid Cooling水冷", "(05)Other"
    ])
    if cooling.endswith("Other"):
        cooling = st.text_input("請輸入散熱方式")

    delivery = st.selectbox("交貨地點", [
        "(01)Taiwan", "(02)China", "(03)Thailand", "(04)Vietnam", "(05)Other"
    ])
    if delivery.endswith("Other"):
        delivery = st.text_input("請輸入交貨地點")

    sample_date = st.date_input("樣品需求日期", value=datetime.date.today())
    sample_qty = st.text_input("樣品需求數量")
    demand_qty = st.text_input("需求量 (預估數量/總年數)")

    st.text("Schedule")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        si = st.text_input("SI")
    with col2:
        pv = st.text_input("PV")
    with col3:
        mv = st.text_input("MV")
    with col4:
        mp = st.text_input("MP")

    return {
        "Product_Application": product_app,
        "Cooling_Solution": cooling,
        "Delivery_Location": delivery,
        "Sample_Date": sample_date.strftime("%Y/%m/%d"),
        "Sample_Qty": sample_qty,
        "Demand_Qty": demand_qty,
        "Schedule_SI": si,
        "Schedule_PV": pv,
        "Schedule_MV": mv,
        "Schedule_MP": mp
    }

# ========== C. 規格資訊 ==========
def render_spec_info():
    st.header("C. 規格資訊")

    spec_option = st.selectbox("散熱方案", ["Air Cooling氣冷", "Fan風扇", "Liquid Cooling水冷"])
    spec_data = {"Spec_Type": spec_option}

    if spec_option == "Air Cooling氣冷":
        spec_data["Air_Flow"] = st.text_input("Air Flow (RPM/Voltage/CFM)")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)")
        spec_data["Max_Power"] = st.text_input("Max Power (W)")
        spec_data["Length"] = st.text_input("Length (mm)")
        spec_data["Width"] = st.text_input("Width (mm)")
        spec_data["Height"] = st.text_input("Height (mm)")

    elif spec_option == "Fan風扇":
        spec_data["Length"] = st.text_input("Length (mm)")
        spec_data["Width"] = st.text_input("Width (mm)")
        spec_data["Height"] = st.text_input("Height (mm)")
        spec_data["Max_Power"] = st.text_input("Max Power (W)")
        spec_data["Input_Voltage"] = st.text_input("Input voltage (V)")
        spec_data["Input_Current"] = st.text_input("Input current (A)")
        spec_data["PQ"] = st.text_input("P-Q")
        spec_data["Speed"] = st.text_input("Rotational speed (RPM)")
        spec_data["Noise"] = st.text_input("Noise (dB)")
        spec_data["Tone"] = st.text_input("Tone")
        spec_data["Sone"] = st.text_input("Sone")
        spec_data["Weight"] = st.text_input("Weight (g)")
        spec_data["Connector_Type"] = st.text_input("端子頭型號")
        spec_data["Connector_Pin"] = st.text_input("線序")
        spec_data["Connector_Length"] = st.text_input("出框線長")

    elif spec_option == "Liquid Cooling水冷":
        spec_data["Plate_Form"] = st.text_input("Plate Form")
        spec_data["Max_Power"] = st.text_input("Max Power (W)")
        spec_data["Tj_Max"] = st.text_input("Tj_Max (°C)")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)")
        spec_data["T_Inlet"] = st.text_input("T_Inlet (°C)")
        spec_data["Chip_Size"] = st.text_input("Chip contact size LxWxH (mm)")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)")
        spec_data["Flow_Rate"] = st.text_input("Flow rate (LPM)")
        spec_data["Impedance"] = st.text_input("Impedance (KPa)")
        spec_data["Max_Loading"] = st.text_input("Max loading (lbs)")

    return spec_data

# ========== 預覽頁 ==========
def preview_page():
    st.title("📑 預覽申請內容")

    form_data = st.session_state["form_data"]
    df = pd.DataFrame([form_data])

    st.table(df)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ 下載並送出"):
            filename = f"Project_Form_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df.to_excel(filename, index=False)

            sheet.append_row(list(form_data.values()))

            with open(filename, "rb") as f:
                st.download_button("⬇️ 下載 Excel", f, file_name=filename)

            st.success("✅ 已送出並記錄到 Google Sheet！")

    with col2:
        if st.button("❌ 取消"):
            st.session_state["preview_mode"] = False
            st.rerun()

# ========== 表單頁 ==========
def form_page():
    st.title("📝 專案申請表單")

    customer_info = render_customer_info()
    project_info = render_project_info()
    spec_info = render_spec_info()

    if st.button("完成"):
        form_data = {**customer_info, **project_info, **spec_info}
        form_data["建立時間"] = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
        st.session_state["form_data"] = form_data
        st.session_state["preview_mode"] = True
        st.rerun()

# ========== 主程式 ==========
def main():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if "preview_mode" not in st.session_state:
        st.session_state["preview_mode"] = False
    if "form_data" not in st.session_state:
        st.session_state["form_data"] = {}

    if not st.session_state["logged_in"]:
        login_page()
    else:
        if st.session_state["preview_mode"]:
            preview_page()
        else:
            form_page()

if __name__ == "__main__":
    main()