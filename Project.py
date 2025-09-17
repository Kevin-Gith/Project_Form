import streamlit as st
import gspread
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
    "sam@kipotec.com.tw": "1234",
    "sale1@kipotec.com.tw": "abcd",
    "sale5@kipotec.com.tw": "pass123",
    "sale2@kipotec.com.tw": "0000"
}

USER_MAPPING = {
    "sam@kipotec.com.tw": "Sam",
    "sale1@kipotec.com.tw": "Vivian",
    "sale5@kipotec.com.tw": "Wendy",
    "sale2@kipotec.com.tw": "Lillian"
}

# ========== Function：登入頁 ==========
def login_page():
    st.title("🔐 使用者登入")

    username = st.text_input("帳號")
    password = st.text_input("密碼", type="password")

    if st.button("登入"):
        if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
            st.session_state["logged_in"] = True
            st.session_state["user"] = username
            st.success("登入成功！")
            st.rerun()
        else:
            st.error("帳號或密碼錯誤！")

# ========== Function：A. 客戶資訊 ==========
def render_customer_info():
    st.header("A. 客戶資訊")

    # 自動帶出北辦業務
    current_user = st.session_state.get("user", "")
    sales_user = USER_MAPPING.get(current_user, "Unknown")
    st.text_input("北辦業務", value=sales_user, disabled=True)

    odm = st.selectbox("ODM客戶 (RD)", [
        "(01)仁寶", "(02)廣達", "(03)緯創", "(04)華勤",
        "(05)光寶", "(06)技嘉", "(07)智邦", "(08)其他"
    ])

    brand = st.selectbox("品牌客戶 (RD)", [
        "(01)惠普", "(02)聯想", "(03)高通", "(04)華碩",
        "(05)宏碁", "(06)微星", "(07)技嘉", "(08)其他"
    ])

    purpose = st.selectbox("申請目的", [
        "(01)客戶專案開發", "(02)內部新產品開發", "(03)技術平台預研", "(04)其他"
    ])

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

# ========== Function：B. 開案資訊 ==========
def render_project_info():
    st.header("B. 開案資訊")

    product_app = st.selectbox("產品應用", [
        "(01)NB CPU", "(02)NB GPU", "(03)Server",
        "(04)Automotive(Car)", "(05)Other"
    ])

    cooling = st.selectbox("散熱方式", [
        "(01)Air Cooling", "(02)Fan", "(03)Cooler(含Fan)",
        "(04)Liquid Cooling", "(05)Other"
    ])

    delivery = st.selectbox("交貨地點", [
        "(01)Taiwan", "(02)China", "(03)Thailand", "(04)Vietnam", "(05)Other"
    ])

    sample_date = st.date_input("樣品需求日期", value=datetime.date.today())
    sample_qty = st.text_input("樣品需求數量")
    demand_qty = st.text_input("需求量 (預估數量/總年數)")

    st.subheader("Schedule")
    si = st.text_input("SI")
    pv = st.text_input("PV")
    mv = st.text_input("MV")
    mp = st.text_input("MP")

    return {
        "Product_Application": product_app,
        "Cooling_Solution": cooling,
        "Delivery_Location": delivery,
        "Sample_Date": sample_date.strftime("%Y/%m/%d"),
        "Sample_Qty": sample_qty,
        "Demand_Qty": demand_qty,
        "SI": si,
        "PV": pv,
        "MV": mv,
        "MP": mp
    }

# ========== Function：C. 規格資訊 ==========
def render_spec_info():
    st.header("C. 規格資訊")

    spec_option = st.selectbox("散熱方案", ["Air Cooling", "Fan", "Liquid Cooling"])
    spec_data = {"Spec_Type": spec_option}

    if spec_option == "Air Cooling":
        spec_data["Air_Flow"] = st.text_input("Air Flow (RPM/Voltage/CFM)")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)")
        spec_data["Max_Power"] = st.text_input("Max Power (W)")
        spec_data["Length"] = st.text_input("Length (mm)")
        spec_data["Width"] = st.text_input("Width (mm)")
        spec_data["Height"] = st.text_input("Height (mm)")

    elif spec_option == "Fan":
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

    elif spec_option == "Liquid Cooling":
        spec_data["Plate_Form"] = st.text_input("Plate Form")
        spec_data["Max_Power"] = st.text_input("Max Power (W)")
        spec_data["Tj_Max"] = st.text_input("Tj_Max (°C)")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)")
        spec_data["T_Inlet"] = st.text_input("T_Inlet (°C)")
        spec_data["Chip_Size"] = st.text_input("Chip contact size LxWxH (mm)")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)")
        spec_data["Flow_Rate"] = st.text_input("Flow rate (LPM)")
        spec_data["Impedance"] = st.text_input("Impedance (KPa)")  # ✅ 修正大小寫
        spec_data["Max_Loading"] = st.text_input("Max loading (lbs)")

    return spec_data

# ========== 主程式 ==========
def main():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        login_page()
    else:
        st.title("📌 Project Form 系統")

        customer_info = render_customer_info()
        project_info = render_project_info()
        spec_info = render_spec_info()

        if st.button("完成"):
            record = {**customer_info, **project_info, **spec_info}
            record["Application_Deadline"] = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
            sheet.append_row(list(record.values()))
            st.success("✅ 表單已送出並記錄到 Google Sheet！")

if __name__ == "__main__":
    main()
