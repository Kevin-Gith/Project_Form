import streamlit as st
import gspread
import datetime
from google.oauth2.service_account import Credentials

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

# 固定 Google Sheet 欄位順序
SHEET_HEADERS = [
    "Sales_User", "ODM_Customers", "Brand_Customers", "Application_Purpose",
    "Project_Name", "Proposal_Date", "Product_Application", "Cooling_Solution",
    "Delivery_Location", "Sample_Date", "Sample_Qty", "Demand_Qty",
    "SI", "PV", "MV", "MP", "Spec_Type", "Update_Time"
]

# 使用者帳號密碼
USER_CREDENTIALS = {
    "sam@kipotec.com.tw": {"password": "Kipo-0926969586$$$", "name": "Sam"},
    "sale1@kipotec.com.tw": {"password": "Kipo-0917369466$$$", "name": "Vivian"},
    "sale5@kipotec.com.tw": {"password": "Kipo-0925698417$$$", "name": "Wendy"},
    "sale2@kipotec.com.tw": {"password": "Kipo-0905038111$$$", "name": "Lillian"},
}

# ========== 登出 ==========
def logout():
    keep_keys = {"page", "logged_in"}
    for key in list(st.session_state.keys()):
        if key not in keep_keys:
            del st.session_state[key]
    st.session_state["page"] = "login"
    st.session_state["logged_in"] = False

# ========== 儲存到 Google Sheet ==========
def save_to_google_sheet(record):
    # Spec_Type 只存方案名稱
    record["Spec_Type"] = ", ".join(record.get("Spec_Type", []))
    record["Update_Time"] = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")

    row = [record.get(col, "") for col in SHEET_HEADERS]
    sheet.append_row(row)

# ========== 頁面：登入 ==========
def login_page():
    st.title("💻 Kipo專案申請系統")

    username = st.text_input("帳號", key="login_username")
    password = st.text_input("密碼", type="password", key="login_password")

    if st.button("🔑 登入"):
        if username in USER_CREDENTIALS and USER_CREDENTIALS[username]["password"] == password:
            st.session_state["logged_in"] = True
            st.session_state["user"] = USER_CREDENTIALS[username]["name"]
            st.session_state["page"] = "form"
        else:
            st.error("帳號或密碼錯誤，請重新輸入")

# ========== 頁面：A. 客戶資訊 ==========
def render_customer_info():
    st.header("A. 客戶資訊")
    st.write(f"**北辦業務：{st.session_state.get('user','')}**")

    odm = st.selectbox("ODM客戶 (RD)", ["", "(01)仁寶", "(02)廣達", "(03)緯創", "(04)華勤", "(05)光寶", "(06)技嘉", "(07)智邦", "(08)其他"], key="odm")
    if odm == "(08)其他":
        odm = st.text_input("請輸入ODM客戶", key="odm_other")

    brand = st.selectbox("品牌客戶 (RD)", ["", "(01)惠普", "(02)聯想", "(03)高通", "(04)華碩", "(05)宏碁", "(06)微星", "(07)技嘉", "(08)其他"], key="brand")
    if brand == "(08)其他":
        brand = st.text_input("請輸入品牌客戶", key="brand_other")

    purpose = st.selectbox("申請目的", ["", "(01)客戶專案開發", "(02)內部新產品開發", "(03)技術平台預研", "(04)其他"], key="purpose")
    if purpose == "(04)其他":
        purpose = st.text_input("請輸入申請目的", key="purpose_other")

    project_name = st.text_input("客戶專案名稱", key="project_name")
    proposal_date = st.date_input("客戶提案日期", value=datetime.date.today(), key="proposal_date")

    return {
        "Sales_User": st.session_state["user"],
        "ODM_Customers": odm,
        "Brand_Customers": brand,
        "Application_Purpose": purpose,
        "Project_Name": project_name,
        "Proposal_Date": proposal_date.strftime("%Y/%m/%d")
    }

# ========== 頁面：B. 開案資訊 ==========
def render_project_info():
    st.header("B. 開案資訊")

    product_app = st.selectbox("產品應用", ["", "(01)NB CPU", "(02)NB GPU", "(03)Server", "(04)Automotive(Car)", "(05)Other"], key="product_app")
    if product_app == "(05)Other":
        product_app = st.text_input("請輸入產品應用", key="product_app_other")

    cooling = st.selectbox("散熱方式", ["", "(01)Air Cooling", "(02)Fan", "(03)Cooler(含Fan)", "(04)Liquid Cooling", "(05)Other"], key="cooling")
    if cooling == "(05)Other":
        cooling = st.text_input("請輸入散熱方式", key="cooling_other")

    delivery = st.selectbox("交貨地點", ["", "(01)Taiwan", "(02)China", "(03)Thailand", "(04)Vietnam", "(05)Other"], key="delivery")
    if delivery == "(05)Other":
        delivery = st.text_input("請輸入交貨地點", key="delivery_other")

    sample_date = st.date_input("樣品需求日期", value=datetime.date.today(), key="sample_date")
    sample_qty = st.text_input("樣品需求數量", key="sample_qty")
    demand_qty = st.text_input("需求量 (預估數量/總年數)", key="demand_qty")

    st.text("Schedule")  # 和樣品需求數量一樣樣式
    col1, col2, col3, col4 = st.columns(4)
    si = col1.text_input("SI", key="si")
    pv = col2.text_input("PV", key="pv")
    mv = col3.text_input("MV", key="mv")
    mp = col4.text_input("MP", key="mp")

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

# ========== 頁面：C. 規格資訊 ==========
def render_spec_info():
    st.header("C. 規格資訊")
    spec_options = st.multiselect("選擇散熱方案", ["Air Cooling", "Fan", "Liquid Cooling"], key="spec_options")

    spec_data = {"Spec_Type": spec_options}

    if "Air Cooling" in spec_options:
        st.subheader("Air Cooling 氣冷")
        spec_data["Air_Flow"] = st.text_input("Air Flow (RPM/Voltage/CFM)", key="air_flow")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)", key="air_tcase")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)", key="air_res")
        spec_data["Max_Power_Air"] = st.text_input("Max Power (W)", key="air_power")
        spec_data["Size_Air"] = st.text_input("Size LxWxH (mm)", key="air_size")

    if "Fan" in spec_options:
        st.subheader("Fan 風扇")
        spec_data["Size_Fan"] = st.text_input("Size LxWxH (mm)", key="fan_size")
        spec_data["Max_Power_Fan"] = st.text_input("Max Power (W)", key="fan_power")
        spec_data["Input_Voltage"] = st.text_input("Input voltage (V)", key="fan_volt")
        spec_data["Input_Current"] = st.text_input("Input current (A)", key="fan_curr")
        spec_data["PQ"] = st.text_input("P-Q", key="fan_pq")
        spec_data["Speed"] = st.text_input("Rotational speed (RPM)", key="fan_speed")
        spec_data["Noise"] = st.text_input("Noise (dB)", key="fan_noise")
        spec_data["Tone"] = st.text_input("Tone", key="fan_tone")
        spec_data["Sone"] = st.text_input("Sone", key="fan_sone")
        spec_data["Weight"] = st.text_input("Weight (g)", key="fan_weight")
        spec_data["Connector"] = st.text_input("端子頭型號、線序、出框線長", key="fan_connector")

    if "Liquid Cooling" in spec_options:
        st.subheader("Liquid Cooling 水冷")
        spec_data["Plate_Form"] = st.text_input("Plate Form", key="liq_plate")
        spec_data["Max_Power_Liquid"] = st.text_input("Max Power (W)", key="liq_power")
        spec_data["Tj_Max"] = st.text_input("Tj_Max (°C)", key="liq_tj")
        spec_data["Tcase_Max_Liquid"] = st.text_input("Tcase_Max (°C)", key="liq_tcase")
        spec_data["T_Inlet"] = st.text_input("T_Inlet (°C)", key="liq_inlet")
        spec_data["Chip_Length"] = st.text_input("Chip contact Length (mm)", key="liq_len")
        spec_data["Chip_Width"] = st.text_input("Chip contact Width (mm)", key="liq_wid")
        spec_data["Chip_Height"] = st.text_input("Chip contact Height (mm)", key="liq_hei")
        spec_data["Thermal_Resistance_Liquid"] = st.text_input("Thermal Resistance (°C/W)", key="liq_res")
        spec_data["Flow_Rate"] = st.text_input("Flow rate (LPM)", key="liq_flow")
        spec_data["Impedance"] = st.text_input("Impedance (KPa)", key="liq_imp")
        spec_data["Max_Loading"] = st.text_input("Max loading (lbs)", key="liq_load")

    return spec_data

# ========== 頁面：表單 ==========
def form_page():
    st.title("💻 Kipo專案申請系統")
    if st.button("🚪 登出"):
        logout()

    customer_info = render_customer_info()
    project_info = render_project_info()
    spec_info = render_spec_info()

    if st.button("✅ 完成"):
        if not customer_info["ODM_Customers"] or not customer_info["Brand_Customers"] or not customer_info["Application_Purpose"] or not customer_info["Project_Name"]:
            st.error("客戶資訊未完成填寫，請重新確認")
        elif not project_info["Product_Application"] or not project_info["Cooling_Solution"] or not project_info["Delivery_Location"] or not project_info["Sample_Qty"] or not project_info["Demand_Qty"]:
            st.error("開案資訊未完成填寫，請重新確認")
        else:
            st.session_state["record"] = {**customer_info, **project_info, **spec_info}
            st.session_state["page"] = "preview"

# ========== 頁面：預覽 ==========
def preview_page():
    st.title("📋 預覽填寫內容")

    record = st.session_state.get("record", {})

    st.subheader("A. 客戶資訊")
    for k in ["Sales_User", "ODM_Customers", "Brand_Customers", "Application_Purpose", "Project_Name", "Proposal_Date"]:
        st.write(f"**{k}：** {record.get(k, '')}")

    st.subheader("B. 開案資訊")
    for k in ["Product_Application", "Cooling_Solution", "Delivery_Location", "Sample_Date", "Sample_Qty", "Demand_Qty", "SI", "PV", "MV", "MP"]:
        st.write(f"**{k}：** {record.get(k, '')}")

    st.subheader("C. 規格資訊")
    for option in record.get("Spec_Type", []):
        st.write(f"### {option}")
        for k, v in record.items():
            if option in k or k.startswith(option[:3].lower()):  # 簡單對應顯示
                st.write(f"- {k}：{v}")

    col1, col2 = st.columns(2)
    if col1.button("🔙 返回修改"):
        st.session_state["page"] = "form"
    if col2.button("💾 確認送出"):
        save_to_google_sheet(record)
        st.success("✅ 已寫入 Google Sheet！")

# ========== 主程式 ==========
def main():
    if "page" not in st.session_state:
        st.session_state["page"] = "login"
        st.session_state["logged_in"] = False

    if st.session_state["page"] == "login":
        login_page()
    elif st.session_state["page"] == "form":
        form_page()
    elif st.session_state["page"] == "preview":
        preview_page()

if __name__ == "__main__":
    main()
