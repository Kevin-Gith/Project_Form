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
    "sam": "1234",
    "vivian": "abcd",
    "wendy": "pass123",
    "lillian": "0000"
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
            st.experimental_rerun()
        else:
            st.error("帳號或密碼錯誤！")

# ========== Function：A. 客戶資訊 ==========
def render_customer_info():
    st.header("A. 客戶資訊")

    user_mapping = {
        "sam": "Sam",
        "vivian": "Vivian",
        "wendy": "Wendy",
        "lillian": "Lillian"
    }
    sales_user = user_mapping.get(st.session_state.get("user"), "Unknown")

    odm = st.text_input("ODM 客戶 (RD)")
    brand = st.text_input("品牌客戶 (RD)")
    applicant = st.text_input("申請人", sales_user)

    return {
        "ODM_Customers": odm,
        "Brand_Customers": brand,
        "Applicant": applicant
    }

# ========== Function：B. 開案資訊 ==========
def render_project_info():
    st.header("B. 開案資訊")

    purpose = st.selectbox("申請目的", ["(01)開案", "(02)量產", "(00)Other"])
    if purpose == "(00)Other":
        purpose = st.text_input("請輸入其他申請目的")

    product_app = st.text_input("產品應用")
    cooling = st.selectbox("散熱方式", ["Air Cooling", "Fan", "Liquid Cooling", "(00)Other"])
    if cooling == "(00)Other":
        cooling = st.text_input("請輸入其他散熱方式")

    delivery = st.text_input("交貨地點")

    sample_date = st.date_input("樣品需求日期", value=datetime.date.today())
    sample_qty = st.number_input("樣品需求數量", min_value=1, step=1)

    st.subheader("Schedule")
    si = st.text_input("SI")
    pv = st.text_input("PV")
    mv = st.text_input("MV")
    mp = st.text_input("MP")

    demand_qty = st.text_input("需求量 (預估數量/總年數)")

    return {
        "Application_Purpose": purpose,
        "Product_Application": product_app,
        "Cooling_Solution": cooling,
        "Delivery_Location": delivery,
        "Sample_Date": sample_date.strftime("%Y/%m/%d"),
        "Sample_Qty": sample_qty,
        "SI": si,
        "PV": pv,
        "MV": mv,
        "MP": mp,
        "Demand_Qty": demand_qty
    }

# ========== Function：C. 規格資訊 ==========
def render_spec_info():
    st.header("C. 規格資訊")
    spec_option = st.selectbox("Cooling Solution", ["Air Cooling", "Fan", "Liquid Cooling"])
    spec_data = {"Spec_Type": spec_option}

    if spec_option == "Air Cooling":
        spec_data["Air_Flow"] = st.text_input("Air Flow (RPM/Voltage/CFM)")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)")
        spec_data["Max_Power"] = st.text_input("Max Power (W)")
        spec_data["Size"] = st.text_input("Size LxWxH (mm)")

    elif spec_option == "Fan":
        spec_data["Size"] = st.text_input("Size LxWxH (mm)")
        spec_data["Max_Power"] = st.text_input("Max Power (W)")
        spec_data["Input_Voltage"] = st.text_input("Input voltage (V)")
        spec_data["Input_Current"] = st.text_input("Input current (A)")
        spec_data["PQ"] = st.text_input("P-Q")
        spec_data["Speed"] = st.text_input("Rotational speed (RPM)")
        spec_data["Noise"] = st.text_input("Noise (dB)")
        spec_data["Tone"] = st.text_input("Tone")
        spec_data["Sone"] = st.text_input("Sone")
        spec_data["Weight"] = st.text_input("Weight (g)")
        spec_data["Connector"] = st.text_input("端子頭型號、線序、出框線長")

    elif spec_option == "Liquid Cooling":
        spec_data["Plate_Form"] = st.text_input("Plate Form")
        spec_data["Max_Power"] = st.text_input("Max Power (W)")
        spec_data["Tj_Max"] = st.text_input("Tj_Max (°C)")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)")
        spec_data["T_Inlet"] = st.text_input("T_Inlet (°C)")
        spec_data["Chip_Size"] = st.text_input("Chip contact size LxWxH (mm)")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)")
        spec_data["Flow_Rate"] = st.text_input("Flow rate (LPM)")
        spec_data["Impedance"] = st.text_input("Impedance (Kpa)")
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
