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
    "sam@kipotec.com.tw": {"password": "Kipo-0926969586$$$", "name": "Sam"},
    "sale1@kipotec.com.tw": {"password": "Kipo-0917369466$$$", "name": "Vivian"},
    "sale5@kipotec.com.tw": {"password": "Kipo-0925698417$$$", "name": "Wendy"},
    "sale2@kipotec.com.tw": {"password": "Kipo-0905038111$$$", "name": "Lillian"},
}


# ========== 登入頁面 ==========
def login_page():
    st.title("💻 Kipo專案申請系統")

    username = st.text_input("帳號")
    password = st.text_input("密碼", type="password")

    if st.button("登入"):
        if username in USER_CREDENTIALS and USER_CREDENTIALS[username]["password"] == password:
            st.session_state["logged_in"] = True
            st.session_state["current_user"] = username
            st.success("✅ 登入成功！")
            st.rerun()
        else:
            st.error("❌ 帳號或密碼錯誤，請重新輸入")


# ========== A. 客戶資訊 ==========
def render_customer_info(current_user):
    st.header("A. 客戶資訊")

    sales_user = USER_CREDENTIALS.get(current_user, {}).get("name", "Unknown")
    st.write(f"**北辦業務：{sales_user}**")

    odm = st.selectbox("ODM 客戶 (RD)",
                       ["(01)仁寶", "(02)廣達", "(03)緯創", "(04)華勤",
                        "(05)光寶", "(06)技嘉", "(07)智邦", "(08)其他"],
                       key="ODM_Customers")
    if odm == "(08)其他":
        odm = st.text_input("請輸入 ODM 客戶", key="ODM_Customers_Other")

    brand = st.selectbox("品牌客戶 (RD)",
                         ["(01)惠普", "(02)聯想", "(03)高通", "(04)華碩",
                          "(05)宏碁", "(06)微星", "(07)技嘉", "(08)其他"],
                         key="Brand_Customers")
    if brand == "(08)其他":
        brand = st.text_input("請輸入品牌客戶", key="Brand_Customers_Other")

    purpose = st.selectbox("申請目的",
                           ["(01)客戶專案開發", "(02)內部新產品開發", "(03)技術平台預研", "(04)其他"],
                           key="Application_Purpose")
    if purpose == "(04)其他":
        purpose = st.text_input("請輸入申請目的", key="Application_Purpose_Other")

    project_name = st.text_input("客戶專案名稱", key="Project_Name")
    proposal_date = st.date_input("客戶提案日期", value=datetime.date.today(), key="Proposal_Date")

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

    product_app = st.selectbox("產品應用",
                               ["(01)NB CPU", "(02)NB GPU", "(03)Server", "(04)Automotive(Car)", "(05)Other"],
                               key="Product_Application")
    if product_app == "(05)Other":
        product_app = st.text_input("請輸入產品應用", key="Product_Application_Other")

    cooling = st.selectbox("散熱方式",
                           ["(01)Air Cooling", "(02)Fan", "(03)Cooler(含Fan)", "(04)Liquid Cooling", "(05)Other"],
                           key="Cooling_Solution")
    if cooling == "(05)Other":
        cooling = st.text_input("請輸入散熱方式", key="Cooling_Solution_Other")

    delivery = st.selectbox("交貨地點",
                            ["(01)Taiwan", "(02)China", "(03)Thailand", "(04)Vietnam", "(05)Other"],
                            key="Delivery_Location")
    if delivery == "(05)Other":
        delivery = st.text_input("請輸入交貨地點", key="Delivery_Location_Other")

    sample_date = st.date_input("樣品需求日期", value=datetime.date.today(), key="Sample_Date")
    sample_qty = st.text_input("樣品需求數量", key="Sample_Qty")
    demand_qty = st.text_input("需求量 (預估數量/總年數)", key="Demand_Qty")

    st.subheader("Schedule")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        si = st.text_input("SI", key="Schedule_SI")
    with col2:
        pv = st.text_input("PV", key="Schedule_PV")
    with col3:
        mv = st.text_input("MV", key="Schedule_MV")
    with col4:
        mp = st.text_input("MP", key="Schedule_MP")

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

    spec_options = st.multiselect("選擇散熱方案", ["Air Cooling氣冷", "Fan風扇", "Liquid Cooling水冷"],
                                  key="Selected_Specs")
    spec_data = {"Selected_Specs": spec_options}

    if not spec_options:
        st.info("👉 請先選擇散熱方案，才會顯示對應的輸入欄位")
        return spec_data

    if "Air Cooling氣冷" in spec_options:
        st.subheader("Air Cooling氣冷")
        spec_data["Air_Flow"] = st.text_input("Air Flow (RPM/Voltage/CFM)", key="Air_Flow")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)", key="Tcase_Max")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)", key="Thermal_Resistance")
        spec_data["Max_Power"] = st.text_input("Max Power (W)", key="Air_Max_Power")
        spec_data["Length_Air"] = st.text_input("Length (mm)", key="Length_Air")
        spec_data["Width_Air"] = st.text_input("Width (mm)", key="Width_Air")
        spec_data["Height_Air"] = st.text_input("Height (mm)", key="Height_Air")

    if "Fan風扇" in spec_options:
        st.subheader("Fan風扇")
        spec_data["Length_Fan"] = st.text_input("Length (mm)", key="Length_Fan")
        spec_data["Width_Fan"] = st.text_input("Width (mm)", key="Width_Fan")
        spec_data["Height_Fan"] = st.text_input("Height (mm)", key="Height_Fan")
        spec_data["Max_Power_Fan"] = st.text_input("Max Power (W)", key="Fan_Max_Power")
        spec_data["Input_Voltage"] = st.text_input("Input voltage (V)", key="Input_Voltage")
        spec_data["Input_Current"] = st.text_input("Input current (A)", key="Input_Current")
        spec_data["PQ"] = st.text_input("P-Q", key="PQ")
        spec_data["Speed"] = st.text_input("Rotational speed (RPM)", key="Speed")
        spec_data["Noise"] = st.text_input("Noise (dB)", key="Noise")
        spec_data["Tone"] = st.text_input("Tone", key="Tone")
        spec_data["Sone"] = st.text_input("Sone", key="Sone")
        spec_data["Weight"] = st.text_input("Weight (g)", key="Weight")
        spec_data["Connector_Type"] = st.text_input("端子頭型號", key="Connector_Type")
        spec_data["Connector_Pin"] = st.text_input("線序", key="Connector_Pin")
        spec_data["Connector_Length"] = st.text_input("出框線長", key="Connector_Length")

    if "Liquid Cooling水冷" in spec_options:
        st.subheader("Liquid Cooling水冷")
        spec_data["Plate_Form"] = st.text_input("Plate Form", key="Plate_Form")
        spec_data["Max_Power_Liquid"] = st.text_input("Max Power (W)", key="Liquid_Max_Power")
        spec_data["Tj_Max"] = st.text_input("Tj_Max (°C)", key="Tj_Max")
        spec_data["Tcase_Max_Liquid"] = st.text_input("Tcase_Max (°C)", key="Tcase_Max_Liquid")
        spec_data["T_Inlet"] = st.text_input("T_Inlet (°C)", key="T_Inlet")
        spec_data["Chip_Size"] = st.text_input("Chip contact size LxWxH (mm)", key="Chip_Size")
        spec_data["Thermal_Resistance_Liquid"] = st.text_input("Thermal Resistance (°C/W)", key="Liquid_Resistance")
        spec_data["Flow_Rate"] = st.text_input("Flow rate (LPM)", key="Flow_Rate")
        spec_data["Impedance"] = st.text_input("Impedance (KPa)", key="Impedance")
        spec_data["Max_Loading"] = st.text_input("Max loading (lbs)", key="Max_Loading")

    return spec_data


# ========== 表單頁面 ==========
def form_page(current_user):
    st.title("📝 Kipo專案申請系統")

    customer_info = render_customer_info(current_user)
    project_info = render_project_info()
    spec_info = render_spec_info()

    if st.button("完成"):
        missing_fields = [k for k, v in {**customer_info, **project_info}.items() if not v]
        if missing_fields:
            st.error("❌ 客戶資訊或開案資訊未完成填寫，請重新確認")
        else:
            form_data = {**customer_info, **project_info, **spec_info}
            form_data["建立時間"] = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
            st.session_state["form_data"] = form_data
            st.session_state["preview_mode"] = True
            st.rerun()


# ========== 預覽頁面 ==========
def preview_page():
    st.title("📑 預覽申請內容")
    form_data = st.session_state["form_data"]

    # A. 客戶資訊
    st.subheader("A. 客戶資訊")
    st.write(form_data["Sales_User"], form_data["ODM_Customers"], form_data["Brand_Customers"],
             form_data["Application_Purpose"], form_data["Project_Name"], form_data["Proposal_Date"])

    # B. 開案資訊
    st.subheader("B. 開案資訊")
    st.write(form_data["Product_Application"], form_data["Cooling_Solution"], form_data["Delivery_Location"],
             form_data["Sample_Date"], form_data["Sample_Qty"], form_data["Demand_Qty"],
             form_data["Schedule_SI"], form_data["Schedule_PV"], form_data["Schedule_MV"], form_data["Schedule_MP"])

    # C. 規格資訊 (即使沒填也顯示)
    st.subheader("C. 規格資訊")
    sections = {
        "Air Cooling氣冷": ["Air_Flow", "Tcase_Max", "Thermal_Resistance", "Max_Power",
                          "Length_Air", "Width_Air", "Height_Air"],
        "Fan風扇": ["Length_Fan", "Width_Fan", "Height_Fan", "Max_Power_Fan", "Input_Voltage", "Input_Current",
                   "PQ", "Speed", "Noise", "Tone", "Sone", "Weight", "Connector_Type", "Connector_Pin", "Connector_Length"],
        "Liquid Cooling水冷": ["Plate_Form", "Max_Power_Liquid", "Tj_Max", "Tcase_Max_Liquid", "T_Inlet",
                              "Chip_Size", "Thermal_Resistance_Liquid", "Flow_Rate", "Impedance", "Max_Loading"]
    }

    for section, keys in sections.items():
        st.markdown(f"**{section}**")
        for k in keys:
            st.write(f"- {k}: {form_data.get(k, '')}")

    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ 下載並送出"):
            filename = f"Project_Form_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            pd.DataFrame([form_data]).to_excel(filename, index=False)
            sheet.append_row(list(form_data.values()))
            with open(filename, "rb") as f:
                st.download_button("⬇️ 下載 Excel", f, file_name=filename)
            st.success("✅ 已送出並記錄到 Google Sheet！")
    with col2:
        if st.button("🔙 返回修改"):
            st.session_state["preview_mode"] = False
            st.rerun()


# ========== 主程式 ==========
def main():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if "preview_mode" not in st.session_state:
        st.session_state["preview_mode"] = False

    if not st.session_state["logged_in"]:
        login_page()
    else:
        if st.session_state["preview_mode"]:
            preview_page()
        else:
            form_page(st.session_state["current_user"])


if __name__ == "__main__":
    main()
