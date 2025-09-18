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


# ========== A. 客戶資訊 ==========
def render_customer_info(current_user):
    st.header("A. 客戶資訊")

    user_mapping = {
        "sam@kipotec.com.tw": "Sam",
        "sale1@kipotec.com.tw": "Vivian",
        "sale5@kipotec.com.tw": "Wendy",
        "sale2@kipotec.com.tw": "Lillian"
    }
    sales_user = user_mapping.get(current_user, "Unknown")

    st.write(f"**北辦業務：{sales_user}**")

    odm = st.selectbox("ODM 客戶 (RD)", ["(01)仁寶", "(02)廣達", "(03)緯創", "(04)華勤", "(05)光寶", "(06)技嘉", "(07)智邦", "(08)其他"])
    if odm == "(08)其他":
        odm = st.text_input("請輸入 ODM 客戶")

    brand = st.selectbox("品牌客戶 (RD)", ["(01)惠普", "(02)聯想", "(03)高通", "(04)華碩", "(05)宏碁", "(06)微星", "(07)技嘉", "(08)其他"])
    if brand == "(08)其他":
        brand = st.text_input("請輸入品牌客戶")

    purpose = st.selectbox("申請目的", ["(01)客戶專案開發", "(02)內部新產品開發", "(03)技術平台預研", "(04)其他"])
    if purpose == "(04)其他":
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

    product_app = st.selectbox("產品應用", ["(01)NB CPU", "(02)NB GPU", "(03)Server", "(04)Automotive(Car)", "(05)Other"])
    if product_app == "(05)Other":
        product_app = st.text_input("請輸入產品應用")

    cooling = st.selectbox("散熱方式", ["(01)Air Cooling", "(02)Fan", "(03)Cooler(含Fan)", "(04)Liquid Cooling", "(05)Other"])
    if cooling == "(05)Other":
        cooling = st.text_input("請輸入散熱方式")

    delivery = st.selectbox("交貨地點", ["(01)Taiwan", "(02)China", "(03)Thailand", "(04)Vietnam", "(05)Other"])
    if delivery == "(05)Other":
        delivery = st.text_input("請輸入交貨地點")

    sample_date = st.date_input("樣品需求日期", value=datetime.date.today())
    sample_qty = st.text_input("樣品需求數量")
    demand_qty = st.text_input("需求量 (預估數量/總年數)")

    st.subheader("Schedule")
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

    spec_options = st.multiselect("選擇散熱方案", ["Air Cooling氣冷", "Fan風扇", "Liquid Cooling水冷"])
    spec_data = {"Selected_Specs": spec_options}

    if "Air Cooling氣冷" in spec_options:
        st.subheader("Air Cooling氣冷")
        spec_data["Air_Flow"] = st.text_input("Air Flow (RPM/Voltage/CFM)", key="air_flow")
        spec_data["Tcase_Max"] = st.text_input("Tcase_Max (°C)", key="air_tcase")
        spec_data["Thermal_Resistance"] = st.text_input("Thermal Resistance (°C/W)", key="air_res")
        spec_data["Max_Power"] = st.text_input("Max Power (W)", key="air_max_power")
        spec_data["Length_Air"] = st.text_input("Length (mm)", key="air_length")
        spec_data["Width_Air"] = st.text_input("Width (mm)", key="air_width")
        spec_data["Height_Air"] = st.text_input("Height (mm)", key="air_height")

    if "Fan風扇" in spec_options:
        st.subheader("Fan風扇")
        spec_data["Length_Fan"] = st.text_input("Length (mm)", key="fan_length")
        spec_data["Width_Fan"] = st.text_input("Width (mm)", key="fan_width")
        spec_data["Height_Fan"] = st.text_input("Height (mm)", key="fan_height")
        spec_data["Max_Power_Fan"] = st.text_input("Max Power (W)", key="fan_max_power")
        spec_data["Input_Voltage"] = st.text_input("Input voltage (V)", key="fan_voltage")
        spec_data["Input_Current"] = st.text_input("Input current (A)", key="fan_current")
        spec_data["PQ"] = st.text_input("P-Q", key="fan_pq")
        spec_data["Speed"] = st.text_input("Rotational speed (RPM)", key="fan_speed")
        spec_data["Noise"] = st.text_input("Noise (dB)", key="fan_noise")
        spec_data["Tone"] = st.text_input("Tone", key="fan_tone")
        spec_data["Sone"] = st.text_input("Sone", key="fan_sone")
        spec_data["Weight"] = st.text_input("Weight (g)", key="fan_weight")
        spec_data["Connector_Type"] = st.text_input("端子頭型號", key="fan_conn_type")
        spec_data["Connector_Pin"] = st.text_input("線序", key="fan_conn_pin")
        spec_data["Connector_Length"] = st.text_input("出框線長", key="fan_conn_len")

    if "Liquid Cooling水冷" in spec_options:
        st.subheader("Liquid Cooling水冷")
        spec_data["Plate_Form"] = st.text_input("Plate Form", key="liq_plate")
        spec_data["Max_Power_Liquid"] = st.text_input("Max Power (W)", key="liq_max_power")
        spec_data["Tj_Max"] = st.text_input("Tj_Max (°C)", key="liq_tj")
        spec_data["Tcase_Max_Liquid"] = st.text_input("Tcase_Max (°C)", key="liq_tcase")
        spec_data["T_Inlet"] = st.text_input("T_Inlet (°C)", key="liq_inlet")
        spec_data["Chip_Size"] = st.text_input("Chip contact size LxWxH (mm)", key="liq_chip")
        spec_data["Thermal_Resistance_Liquid"] = st.text_input("Thermal Resistance (°C/W)", key="liq_res")
        spec_data["Flow_Rate"] = st.text_input("Flow rate (LPM)", key="liq_flow")
        spec_data["Impedance"] = st.text_input("Impedance (KPa)", key="liq_imp")
        spec_data["Max_Loading"] = st.text_input("Max loading (lbs)", key="liq_load")

    return spec_data


# ========== 表單頁面 ==========
def form_page(current_user):
    st.title("📝 Kipo專案申請系統")

    customer_info = render_customer_info(current_user)
    project_info = render_project_info()
    spec_info = render_spec_info()

    if st.button("完成"):
        # 檢查必填欄位
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
    col1, col2 = st.columns(2)
    with col1:
        st.write("**北辦業務**:", form_data["Sales_User"])
        st.write("**ODM客戶 (RD)**:", form_data["ODM_Customers"])
        st.write("**品牌客戶 (RD)**:", form_data["Brand_Customers"])
    with col2:
        st.write("**申請目的**:", form_data["Application_Purpose"])
        st.write("**客戶專案名稱**:", form_data["Project_Name"])
        st.write("**客戶提案日期**:", form_data["Proposal_Date"])

    # B. 開案資訊
    st.subheader("B. 開案資訊")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**產品應用**:", form_data["Product_Application"])
        st.write("**散熱方式**:", form_data["Cooling_Solution"])
        st.write("**交貨地點**:", form_data["Delivery_Location"])
        st.write("**樣品需求日期**:", form_data["Sample_Date"])
    with col2:
        st.write("**樣品需求數量**:", form_data["Sample_Qty"])
        st.write("**需求量**:", form_data["Demand_Qty"])
        st.write("**Schedule SI**:", form_data["Schedule_SI"])
        st.write("**Schedule PV**:", form_data["Schedule_PV"])
        st.write("**Schedule MV**:", form_data["Schedule_MV"])
        st.write("**Schedule MP**:", form_data["Schedule_MP"])

    # C. 規格資訊
    st.subheader("C. 規格資訊")
    if "Air Cooling氣冷" in form_data.get("Selected_Specs", []):
        st.markdown("**Air Cooling氣冷**")
        for k in ["Air_Flow", "Tcase_Max", "Thermal_Resistance", "Max_Power", "Length_Air", "Width_Air", "Height_Air"]:
            if form_data.get(k):
                st.write(f"- {k}: {form_data[k]}")

    if "Fan風扇" in form_data.get("Selected_Specs", []):
        st.markdown("**Fan風扇**")
        for k in ["Length_Fan", "Width_Fan", "Height_Fan", "Max_Power_Fan", "Input_Voltage", "Input_Current",
                  "PQ", "Speed", "Noise", "Tone", "Sone", "Weight", "Connector_Type", "Connector_Pin", "Connector_Length"]:
            if form_data.get(k):
                st.write(f"- {k}: {form_data[k]}")

    if "Liquid Cooling水冷" in form_data.get("Selected_Specs", []):
        st.markdown("**Liquid Cooling水冷**")
        for k in ["Plate_Form", "Max_Power_Liquid", "Tj_Max", "Tcase_Max_Liquid", "T_Inlet", "Chip_Size",
                  "Thermal_Resistance_Liquid", "Flow_Rate", "Impedance", "Max_Loading"]:
            if form_data.get(k):
                st.write(f"- {k}: {form_data[k]}")

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
    current_user = "sam@kipotec.com.tw"  # 模擬登入帳號
    if "preview_mode" not in st.session_state:
        st.session_state["preview_mode"] = False

    if st.session_state["preview_mode"]:
        preview_page()
    else:
        form_page(current_user)


if __name__ == "__main__":
    main()
