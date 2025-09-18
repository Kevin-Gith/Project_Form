import streamlit as st
import gspread
import pandas as pd
import io
import os
from openpyxl import load_workbook
from google.oauth2.service_account import Credentials
from openpyxl.utils import column_index_from_string
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

# 固定 Google Sheet 欄位順序
SHEET_HEADERS = [
    "Sales_User", "ODM_Customers", "Brand_Customers", "Application_Purpose",
    "Project_Name", "Proposal_Date", "Product_Application", "Cooling_Solution",
    "Delivery_Location", "Sample_Date", "Sample_Qty", "Demand_Qty",
    "SI", "PV", "MV", "MP", "Spec_Type", "Update_Time"
]

# ========== 使用者帳號密碼 ==========
USER_CREDENTIALS = {
    "sam@kipotec.com.tw": {"password": "Kipo-0926969586$$$", "name": "Sam"},
    "sale1@kipotec.com.tw": {"password": "Kipo-0917369466$$$", "name": "Vivian"},
    "sale5@kipotec.com.tw": {"password": "Kipo-0925698417$$$", "name": "Wendy"},
    "sale2@kipotec.com.tw": {"password": "Kipo-0905038111$$$", "name": "Lillian"},
}

# ========== 登出功能 ==========
def logout():
    st.session_state.clear()
    st.session_state["page"] = "login"
    st.session_state["logged_in"] = False

# ========== 儲存到 Google Sheet ==========
def save_to_google_sheet(record):
    record_for_sheet = record.copy()
    record_for_sheet["Spec_Type"] = ", ".join(record.get("Spec_Type", {}).keys())
    record_for_sheet["Update_Time"] = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
    row = [record_for_sheet.get(col, "") for col in SHEET_HEADERS]
    sheet.append_row(row)

# ========== 匯出到 Excel 模板 ==========
def export_to_template(record):
    template_path = os.path.join(os.path.dirname(__file__), "Kipo_Project_Form.xlsx")
    wb = load_workbook(template_path)
    ws = wb.active  # 預設第一個工作表

    # A. 客戶資訊
    ws["B5"] = record.get("Sales_User", "")
    ws["B7"] = record.get("ODM_Customers", "")
    ws["E7"] = record.get("Brand_Customers", "")
    ws["B8"] = record.get("Project_Name", "")
    ws["E8"] = record.get("Proposal_Date", "")
    ws["B9"] = record.get("Application_Purpose", "")

    # B. 開案資訊
    ws["B11"] = record.get("Product_Application", "")
    ws["E11"] = record.get("Cooling_Solution", "")
    ws["E13"] = record.get("Delivery_Location", "")
    ws["B12"] = record.get("Sample_Date", "")
    ws["E12"] = record.get("Sample_Qty", "")
    ws["B13"] = record.get("Demand_Qty", "")
    ws["B14"] = record.get("SI", "")
    ws["E14"] = record.get("PV", "")
    ws["B15"] = record.get("MV", "")
    ws["E15"] = record.get("MP", "")

    # ====== C. 規格資訊 ======
    spec_map = {
        "Air Cooling氣冷": "A17",
        "Fan風扇": "C17",
        "Liquid Cooling水冷": "E17"
    }

    for section, fields in record.get("Spec_Type", {}).items():
        if section in spec_map:
            # 把整個區塊組成文字
            lines = [section]  # 先放標題
            for k, v in fields.items():
                value = v if v not in ["", None] else ""
                lines.append(f"{k}: {value}")
            text_value = "\n".join(lines)

            # 一次寫入起始格（合併區左上角）
            ws[spec_map[section]] = text_value

    # 匯出
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

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
    st.write(f"### 北辦業務：{st.session_state.get('user','')}")
    st.header("A. 客戶資訊")

    odm = st.selectbox("ODM客戶 (RD)", ["(01)仁寶", "(02)廣達", "(03)緯創", "(04)華勤", "(05)光寶", "(06)技嘉", "(07)智邦", "(08)其他"], key="odm")
    if odm == "(08)其他":
        odm = st.text_input("請輸入ODM客戶", key="odm_other")

    brand = st.selectbox("品牌客戶 (RD)", ["(01)惠普", "(02)聯想", "(03)高通", "(04)華碩", "(05)宏碁", "(06)微星", "(07)技嘉", "(08)其他"], key="brand")
    if brand == "(08)其他":
        brand = st.text_input("請輸入品牌客戶", key="brand_other")

    purpose = st.selectbox("申請目的", ["(01)客戶專案開發", "(02)內部新產品開發", "(03)技術平台預研", "(04)其他"], key="purpose")
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

    product_app = st.selectbox("產品應用", ["(01)NB CPU", "(02)NB GPU", "(03)Server", "(04)Automotive(Car)", "(05)Other"], key="product_app")
    if product_app == "(05)Other":
        product_app = st.text_input("請輸入產品應用", key="product_app_other")

    cooling = st.selectbox("散熱方式", ["(01)Air Cooling", "(02)Fan", "(03)Cooler(含Fan)", "(04)Liquid Cooling", "(05)Other"], key="cooling")
    if cooling == "(05)Other":
        cooling = st.text_input("請輸入散熱方式", key="cooling_other")

    delivery = st.selectbox("交貨地點", ["(01)Taiwan", "(02)China", "(03)Thailand", "(04)Vietnam", "(05)Other"], key="delivery")
    if delivery == "(05)Other":
        delivery = st.text_input("請輸入交貨地點", key="delivery_other")

    sample_date = st.date_input("樣品需求日期", value=datetime.date.today(), key="sample_date")
    sample_qty = st.text_input("樣品需求數量", key="sample_qty")
    demand_qty = st.text_input("需求量 (預估數量/總年數)", key="demand_qty")

    st.text("需求進度 (Schedule)")
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
    spec_options = st.multiselect("選擇散熱方案", ["Air Cooling氣冷", "Fan風扇", "Liquid Cooling水冷"], key="spec_options")
    spec_data = {}

    if "Air Cooling氣冷" in spec_options:
        st.subheader("Air Cooling氣冷")
        spec_data["Air Cooling氣冷"] = {
            "Air_Flow": st.text_input("Air Flow (RPM/Voltage/CFM)", key="air_flow"),
            "Tcase_Max": st.text_input("Tcase_Max (°C)", key="air_tcase"),
            "Thermal_Resistance": st.text_input("Thermal Resistance (°C/W)", key="air_res"),
            "Max_Power": st.text_input("Max Power (W)", key="air_power"),
            "Length": st.text_input("Length (mm)", key="air_len"),
            "Width": st.text_input("Width (mm)", key="air_wid"),
            "Height": st.text_input("Height (mm)", key="air_hei"),
        }

    if "Fan風扇" in spec_options:
        st.subheader("Fan風扇")
        spec_data["Fan風扇"] = {
            "Length": st.text_input("Length (mm)", key="fan_len"),
            "Width": st.text_input("Width (mm)", key="fan_wid"),
            "Height": st.text_input("Height (mm)", key="fan_hei"),
            "Max_Power": st.text_input("Max Power (W)", key="fan_power"),
            "Input_Voltage": st.text_input("Input voltage (V)", key="fan_volt"),
            "Input_Current": st.text_input("Input current (A)", key="fan_curr"),
            "PQ": st.text_input("P-Q", key="fan_pq"),
            "Speed": st.text_input("Rotational speed (RPM)", key="fan_speed"),
            "Noise": st.text_input("Noise (dB)", key="fan_noise"),
            "Tone": st.text_input("Tone", key="fan_tone"),
            "Sone": st.text_input("Sone", key="fan_sone"),
            "Weight": st.text_input("Weight (g)", key="fan_weight"),
            "Connector": st.text_input("端子頭型號", key="fan_con"),
            "Wiring": st.text_input("線序", key="fan_wire"),
            "Cable_Length": st.text_input("出框線長", key="fan_cable"),
        }

    if "Liquid Cooling水冷" in spec_options:
        st.subheader("Liquid Cooling水冷")
        spec_data["Liquid Cooling水冷"] = {
            "Plate_Form": st.text_input("Plate Form", key="liq_plate"),
            "Max_Power": st.text_input("Max Power (W)", key="liq_max_power"),
            "Tj_Max": st.text_input("Tj_Max (°C)", key="liq_tj"),
            "Tcase_Max": st.text_input("Tcase_Max (°C)", key="liq_tcase"),
            "T_Inlet": st.text_input("T_Inlet (°C)", key="liq_inlet"),
            "Chip_Length": st.text_input("Chip contact Length (mm)", key="liq_chip_length"),
            "Chip_Width": st.text_input("Chip contact Width (mm)", key="liq_chip_width"),
            "Chip_Height": st.text_input("Chip contact Height (mm)", key="liq_chip_height"),
            "Thermal_Resistance": st.text_input("Thermal Resistance (°C/W)", key="liq_res"),
            "Flow_Rate": st.text_input("Flow rate (LPM)", key="liq_flow"),
            "Impedance": st.text_input("Impedance (KPa)", key="liq_imp"),
            "Max_Loading": st.text_input("Max loading (lbs)", key="liq_load")
        }

    return spec_data

# ========== 頁面：表單 ==========
def form_page():
    if not st.session_state.get("logged_in", False):
        st.session_state["page"] = "login"
        return

    st.title("💻 Kipo專案申請系統")
    if st.button("🚪 登出"):
        logout()

    customer_info = render_customer_info()
    project_info = render_project_info()
    spec_info = render_spec_info()

    if st.button("✅ 完成"):
        if any(v in ["", None] for v in customer_info.values()):
            st.error("客戶資訊未完成填寫，請重新確認")
        elif any(v in ["", None] for v in project_info.values()):
            st.error("開案資訊未完成填寫，請重新確認")
        elif not spec_info:
            st.error("規格資訊請至少選擇一種方案")
        else:
            st.session_state["record"] = {
                **customer_info, **project_info, "Spec_Type": spec_info
            }
            st.session_state["page"] = "preview"

# ========== 頁面：預覽 ==========
def preview_page():
    st.title("📋 預覽填寫內容")

    record = st.session_state.get("record", {})
    st.write(f"### 北辦業務：{record.get('Sales_User','')}")

    st.subheader("A. 客戶資訊")
    field_map_a = {
        "ODM_Customers": "ODM客戶(RD)",
        "Brand_Customers": "品牌客戶(RD)",
        "Application_Purpose": "申請目的",
        "Project_Name": "客戶專案名稱",
        "Proposal_Date": "客戶提案日期"
    }
    for k, v in field_map_a.items():
        st.write(f"**{v}：** {record.get(k, '')}")

    st.subheader("B. 開案資訊")
    field_map_b = {
        "Product_Application": "產品應用",
        "Cooling_Solution": "散熱方式",
        "Delivery_Location": "交貨地點",
        "Sample_Date": "樣品需求日期",
        "Sample_Qty": "樣品需求數量",
        "Demand_Qty": "需求量(預估數量/總年數)",
    }
    for k, v in field_map_b.items():
        st.write(f"**{v}：** {record.get(k, '')}")

    st.markdown("#### Schedule")
    col1, col2, col3, col4 = st.columns(4)
    col1.write(f"**SI：** {record.get('SI','')}")
    col2.write(f"**PV：** {record.get('PV','')}")
    col3.write(f"**MV：** {record.get('MV','')}")
    col4.write(f"**MP：** {record.get('MP','')}")

    st.subheader("C. 規格資訊")
    for section, fields in record.get("Spec_Type", {}).items():
        st.markdown(f"**{section}**")
        for k, v in fields.items():
            st.write(f"{k}：{v}")

    col1, col2 = st.columns(2)
    if col1.button("🔙 返回修改"):
        st.session_state["page"] = "form"

    if col2.button("💾 確認送出"):
        save_to_google_sheet(record)

        excel_data = export_to_template(record)

        st.download_button(
            label="⬇️ 下載Excel檔案",
            data=excel_data,
            file_name=f"ProjectForm_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("✅ 已下載Excel檔案並記錄到Google Sheet！")

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