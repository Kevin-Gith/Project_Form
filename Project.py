import streamlit as st
import gspread
import pandas as pd
import io
import os
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment
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

# ========== 欄位單位對照表 ==========
UNIT_MAP = {
    "Air_Flow": "RPM/Voltage/CFM",
    "Tcase_Max": "°C",
    "Thermal_Resistance": "°C/W",
    "Max_Power": "W",
    "Length": "mm",
    "Width": "mm",
    "Height": "mm",

    "Input_Voltage": "V",
    "Input_Current": "A",
    "PQ": "",
    "Speed": "RPM",
    "Noise": "dB",
    "Tone": "",
    "Sone": "",
    "Weight": "g",
    "Connector": "",
    "Wiring": "",
    "Cable_Length": "mm",

    "Plate_Form": "",
    "Tj_Max": "°C",
    "T_Inlet": "°C",
    "Chip_Length": "mm",
    "Chip_Width": "mm",
    "Chip_Height": "mm",
    "Flow_Rate": "LPM",
    "Impedance": "KPa",
    "Max_Loading": "lbs"
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
            lines = [section, ""]  # 標題 + 空行
            # 先一般欄位，最後三個欄位排最後
            if section == "Air Cooling氣冷":
                last_keys = ["Length", "Width", "Height"]
            elif section == "Fan風扇":
                last_keys = ["Length", "Width", "Height"]
            elif section == "Liquid Cooling水冷":
                last_keys = ["Chip_Length", "Chip_Width", "Chip_Height"]
            else:
                last_keys = []

            for k, v in fields.items():
                if k not in last_keys:
                    value = v if v not in ["", None] else ""
                    unit = UNIT_MAP.get(k, "")
                    unit_str = f" ({unit})" if unit else ""
                    lines.append(f"{k}{unit_str}: {value}")

            for k in last_keys:
                if k in fields:
                    value = fields.get(k, "")
                    value = value if value not in ["", None] else ""
                    unit = UNIT_MAP.get(k, "")
                    unit_str = f" ({unit})" if unit else ""
                    lines.append(f"{k}{unit_str}: {value}")

            text_value = "\n".join(lines)
            cell = ws[spec_map[section]]
            cell.value = text_value
            cell.alignment = Alignment(wrapText=True, vertical="top")

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

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
            "Length": st.text_input("Length (mm)", key="fan_len"),
            "Width": st.text_input("Width (mm)", key="fan_wid"),
            "Height": st.text_input("Height (mm)", key="fan_hei"),
        }

    if "Liquid Cooling水冷" in spec_options:
        st.subheader("Liquid Cooling水冷")
        spec_data["Liquid Cooling水冷"] = {
            "Plate_Form": st.text_input("Plate Form", key="liq_plate"),
            "Max_Power": st.text_input("Max Power (W)", key="liq_max_power"),
            "Tj_Max": st.text_input("Tj_Max (°C)", key="liq_tj"),
            "Tcase_Max": st.text_input("Tcase_Max (°C)", key="liq_tcase"),
            "T_Inlet": st.text_input("T_Inlet (°C)", key="liq_inlet"),
            "Thermal_Resistance": st.text_input("Thermal Resistance (°C/W)", key="liq_res"),
            "Flow_Rate": st.text_input("Flow rate (LPM)", key="liq_flow"),
            "Impedance": st.text_input("Impedance (KPa)", key="liq_imp"),
            "Max_Loading": st.text_input("Max loading (lbs)", key="liq_load"),
            "Chip_Length": st.text_input("Chip contact Length (mm)", key="liq_chip_length"),
            "Chip_Width": st.text_input("Chip contact Width (mm)", key="liq_chip_width"),
            "Chip_Height": st.text_input("Chip contact Height (mm)", key="liq_chip_height"),
        }

    return spec_data

# ========== 預覽 ==========
def preview_page():
    st.title("📋 預覽填寫內容")

    record = st.session_state.get("record", {})
    st.write(f"### 北辦業務：{record.get('Sales_User','')}")

    st.subheader("C. 規格資訊")
    for section, fields in record.get("Spec_Type", {}).items():
        st.markdown(f"**{section}**")
        for k, v in fields.items():
            unit = UNIT_MAP.get(k, "")
            unit_str = f" ({unit})" if unit else ""
            st.write(f"{k}{unit_str}: {v}")