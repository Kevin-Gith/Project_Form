import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import datetime

# ========== Google Sheet 設定 ==========
SHEET_NAME = "Project_Form"
WORKSHEET_NAME = "Python"

# 明確設定 scope
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES
)

client = gspread.authorize(creds)
sheet = client.open("Project_Form").worksheet("Python")

try:
    first_row = sheet.row_values(1)
    st.success("✅ 已成功連線到 Google Sheet")
    st.write("第一列資料：", first_row)
except Exception as e:
    st.error(f"❌ 連線失敗: {e}")

# ========== Function：A. 客戶資訊 ==========
def render_customer_info():
    st.header("A. 客戶資訊")

    # 模擬登入帳號對應業務
    user_mapping = {
        "sam@company.com": "Sam",
        "vivian@company.com": "Vivian",
        "wendy@company.com": "Wendy",
        "lillian@company.com": "Lillian"
    }
    current_user = "sam@company.com"  # ⚠️ 未來可改成實際登入帳號
    sales_user = user_mapping.get(current_user, "Unknown")

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

    # 新增需求
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


# ========== 主程式入口 ==========
def main():
    st.title("📌 Project Form 系統")

    # 三個區塊
    customer_info = render_customer_info()
    project_info = render_project_info()
    spec_info = render_spec_info()

    if st.button("完成"):
        # 組合所有資料
        record = {**customer_info, **project_info, **spec_info}
        record["Application_Deadline"] = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")

        # 寫入 Google Sheet
        sheet.append_row(list(record.values()))

        st.success("✅ 表單已送出並記錄到 Google Sheet！")


if __name__ == "__main__":
    main()
