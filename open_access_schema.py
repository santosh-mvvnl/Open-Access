r"""
Open Access & Billing Revenue Schema, Taxonomy & Transformation Engine
Dataset: UPPCL Monthly Master Billing Registers (5 DISCOMs, 155 Authentic Columns)
Authoritative Mapping: D:\Division Name Mapping\UPPCL Division Code Mapping.xlsx
"""

import os
import re
import math
from datetime import datetime
from typing import Dict, Any, Optional, Tuple
import pandas as pd

# The authentic 155 columns from UPPCL Master Billing Register
MASTER_COLUMNS_155 = [
    'ACCT_ID', 'DIV_CODE', 'SDO_CODE', 'TOWN', 'KNO', 'SCNO', 'MOBILE_NO', 'LANDLINE_NO',
    'BOOK_NO', 'NAME', 'FATHER_NAME', 'ADDRESS', 'SUPPLY_TYPE', 'LOAD', 'LOAD_UNIT',
    'DOC', 'CON_STATUS', 'PROJECT_AREA', 'CONNECTION_TYPE', 'NO_OF_AC', 'OPR_FLG',
    'BILL_AFTER_DATE', 'INSTALLATION_DATE', 'SERVICE_CYCLE_CODE', 'BILL_CYC_CD',
    'HEALTHY_CONSUMER', 'GOVT_CODE', 'VILLAGE_CODE', 'VILLAGE_NAME', 'HABITAT_CODE',
    'HABITAT_NAME', 'DISTRICT', 'PAYOFF_BAL', 'NETWID', 'CT_RATIO', 'PT_RATIO',
    'MULTIPLY_FACTOR', 'LAT', 'LON', 'LOOMS_GRTRTHN_60', 'LOOMS_LESSTHN_60',
    'SUPPLY_VOLTAGE', 'SERIAL_NBR', 'METER_BADGE_NO', 'MTR_MAKE', 'METER_VOLTAGE',
    'MTR_TYPE_CD', 'METER_STATUS', 'MANUL_BILL_EXCEP', 'MTR_CONFIG_TY_CD',
    'MR_SOURCE_CD', 'MTR_EXCEP_CD', 'MTR_NO_RECORDED', 'MTR_MAKE_RECORDED',
    'METER_READ_REMARK', 'MTR_RDR_AGENCY', 'MR_SOURCE_CD_MTR', 'MR_READ',
    'OPENING_READING_KWH', 'OPENING_READING_KVAH', 'OPENING_READING_DATE',
    'CLOSING_READING_KWH', 'CLOSING_READING_KVAH', 'CLOSING_READING_DATE',
    'CLOSE_READING', 'CONSUMPTION_KWH', 'CONSUMPTION_KVAH', 'CONSUMPTION_CURR_MNTH',
    'CONSUMPTION_PREV_MNTH', 'CONSUMPTION_PREV_TO_PREV_MNTH', 'LAST_OK_READING',
    'LAST_OK_STATUS', 'MET_FAULTY_CNT', 'SBM_BILL_DATE', 'SBM_MACHINE_ID',
    'LAST_BILL_DATE', 'MDI', 'BILL_BASIS', 'BILL_TYP', 'INF_BILL',
    'TEMP_DISCON_DATE', 'ARREAR', 'LPSC', 'EC', 'FC', 'REBATES', 'MIN_CHARGES',
    'FUEL_SURCHARGE', 'DEMAND_AMT', 'LTMETERING_CHARGES', 'CAP_CHARGES',
    'REG_SURCHARGE', 'REG_SURCHARGE_2', 'ELECTRICITY_DUTY', 'TARIFF_ADJUSTMENTS',
    'PROV_ADJUSTMENTS', 'CURRENT_CYCLE_LPSC', 'CURRENT_ASSESSMENT', 'ENERGY_AMT',
    'NON_ENERGY_AMT', 'BILLED_AMOUNT', 'DUE_DATE_REBATE', 'CROSS_SUBS', 'PTW_SUBS',
    'TARIFF_SUBS', 'POWERLOOM_SUBS', 'WMCR_SUBS', 'WMCR_UI', 'PAY_AMT', 'PAY_DATE',
    'TOTAL_PAY_AMT', 'PAYMENT_MODE', 'TOTAL_OUTSTANDING', 'INOP_AMT', 'SUBSTATION',
    'FEEDER', 'DT', 'POLE_NO', 'SS_NAME', 'SS_GIS_CODE', 'FEEDER_NAME',
    'FEEDER_GIS_CODE', 'DT_NAME', 'DT_GIS_ID', 'AMISPCD_CODE', 'CMPRESYS_CODE',
    'DT_GIS_CODE', 'KW_CONSUMPTION', 'KW_OP_RDG', 'KW_CL_READING',
    'KVAH_CONSUMPTION', 'KVAH_OP_RDG', 'KVAH_CL_READING', 'KVE_CONSUMPTION',
    'KVE_OP_RDG', 'KVE_CL_READING', 'KWH_CONSUMPTION', 'KWH_OP_RDG',
    'KWH_CL_READING', 'KWHE_CONSUMPTION', 'KWHE_OP_RDG', 'KWHE_CL_READING',
    'KVHX_CONSUMPTION', 'KVHX_OP_RDG', 'KVHX_CL_READING', 'KWHX_CONSUMPTION',
    'KWHX_OP_RDG', 'KWHX_CL_READING', 'SAUBHAGYA_FLAG', 'CA_EX_SUBS',
    'BILL_CRE_DTTM', 'POWER_FACTOR', 'PTW_SCHEME_REGD', 'SM_MTR_INSTALL_DT', 'ADVANCE'
]

# DISCOMs present in UPPCL Master Data
DISCOMS = ['DVVNL', 'KESCO', 'MVVNL', 'PuVVNL', 'PVVNL']

# Administrative Hierarchy Columns
ADMIN_COLUMNS = ['DISCOM', 'ZONE_NAME', 'CIRCLE_NAME', 'DIV_CODE', 'DIVISION_NAME', 'DISTRICT']

# Month dictionary
MONTH_NAMES = {
    '01': 'January', '02': 'February', '03': 'March', '04': 'April',
    '05': 'May', '06': 'June', '07': 'July', '08': 'August',
    '09': 'September', '10': 'October', '11': 'November', '12': 'December'
}

MONTH_SHORT_TO_NUM = {
    'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04',
    'MAY': '05', 'JUN': '06', 'JUL': '07', 'AUG': '08',
    'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'
}

# ----------------- INDIAN NUMBER FORMATTING UTILITY (RULE #8) -----------------
def format_in(val: Any, decimals: int = 0, prefix: str = "") -> str:
    """
    Formats integers and floats using the Indian numbering system:
    1,000, 10,000, 1,00,000, 10,00,000, 1,00,00,000.
    Preserves precision and gracefully handles nulls, negatives, and currency.
    """
    if val is None or pd.isna(val) or val == "":
        return "-"
    try:
        num = float(val)
    except (ValueError, TypeError):
        return str(val)
    
    is_neg = num < 0
    num = abs(num)
    
    if decimals > 0:
        dec_fmt = f"{num:.{decimals}f}"
        integer_part, decimal_part = dec_fmt.split(".")
        dec_suffix = "." + decimal_part
    else:
        integer_part = str(int(round(num)))
        dec_suffix = ""
        
    if len(integer_part) <= 3:
        res = integer_part
    else:
        last3 = integer_part[-3:]
        remaining = integer_part[:-3]
        groups = []
        while len(remaining) > 2:
            groups.insert(0, remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            groups.insert(0, remaining)
        res = ",".join(groups) + "," + last3
        
    sign = "-" if is_neg else ""
    return f"{sign}{prefix}{res}{dec_suffix}"

def format_currency_dual(val: Any) -> str:
    """Formats large financial figures with ₹ and Crore representation when >= 1 Cr."""
    if val is None or pd.isna(val) or val == "":
        return "₹ 0.00"
    try:
        num = float(val)
    except (ValueError, TypeError):
        return str(val)
    
    abs_num = abs(num)
    sign = "-" if num < 0 else ""
    if abs_num >= 10000000:
        cr = num / 10000000.0
        return f"{sign}₹ {abs(cr):.2f} Cr ({format_in(num, decimals=2, prefix='₹ ')})"
    return format_in(num, decimals=2, prefix="₹ ")

# ----------------- MASTER PERIOD EXTRACTION -----------------
def parse_folder_period(folder_path: str) -> Dict[str, str]:
    """
    Extracts Master Month (first two digits) and Master Year (last four digits)
    from folder name, e.g. '102026' -> month='10', year='2026'.
    """
    if not folder_path:
        return {'month': '10', 'month_name': 'October', 'year': '2026', 'display': 'October 2026'}
    
    folder_name = os.path.basename(os.path.normpath(folder_path)).strip()
    # Check if folder name has at least 6 characters with digits
    digits = re.findall(r'\d+', folder_name)
    raw = "".join(digits)
    if len(raw) >= 6:
        m_str = raw[:2]
        y_str = raw[-4:]
    else:
        m_str = '10'
        y_str = '2026'
        
    m_int = int(m_str) if m_str.isdigit() else 10
    if m_int < 1 or m_int > 12:
        m_str = '10'
        m_int = 10
    m_padded = f"{m_int:02d}"
    m_name = MONTH_NAMES.get(m_padded, 'Unknown')
    
    return {
        'month': m_padded,
        'month_name': m_name,
        'year': y_str,
        'display': f"{m_name} {y_str}"
    }

# ----------------- AUTHORITATIVE DIVISION HIERARCHY MAPPING -----------------
def load_division_mapping(mapping_path: str = r"D:\Division Name Mapping\UPPCL Division Code Mapping.xlsx") -> Dict[str, Dict[str, str]]:
    """
    Loads authoritative division hierarchy mapping:
    Div_Code -> {DIVISION_NAME, CIRCLE_NAME, ZONE_NAME, DISCOM, DISTRICT}
    """
    div_map = {}
    if not os.path.exists(mapping_path):
        return div_map
    
    try:
        # Sheet 1: Major Divisions
        df1 = pd.read_excel(mapping_path, sheet_name="Major Divisions")
        for _, row in df1.iterrows():
            code = str(row.get("Div_Code", "")).strip().upper()
            if code and code != "NAN":
                div_map[code] = {
                    "DIVISION_NAME": str(row.get("Division Name", "")).strip(),
                    "CIRCLE_NAME": str(row.get("Circle Name", "")).strip(),
                    "ZONE_NAME": str(row.get("Zone Name", "")).strip(),
                    "DISCOM": str(row.get("DISCOM", "")).strip().upper(),
                    "DISTRICT": str(row.get("District", "")).strip()
                }
    except Exception as e:
        print(f"Warning: Could not load division mapping from Major Divisions: {e}")
        
    try:
        # Fallback / augment from other sheets if present
        xl = pd.ExcelFile(mapping_path)
        for s in xl.sheet_names:
            if s != "Major Divisions":
                try:
                    df = xl.parse(s)
                    for _, row in df.iterrows():
                        code = str(row.get("Div_Code", row.get("DIV_CODE", ""))).strip().upper()
                        if code and code != "NAN" and code not in div_map:
                            div_map[code] = {
                                "DIVISION_NAME": str(row.get("Division Name", row.get("DIVISION_NAME", ""))).strip(),
                                "CIRCLE_NAME": str(row.get("Circle Name", row.get("CIRCLE_NAME", ""))).strip(),
                                "ZONE_NAME": str(row.get("Zone Name", row.get("ZONE_NAME", ""))).strip(),
                                "DISCOM": str(row.get("DISCOM", "")).strip().upper(),
                                "DISTRICT": str(row.get("District", row.get("DISTRICT", ""))).strip()
                            }
                except Exception:
                    pass
    except Exception:
        pass

    return div_map

# ----------------- DERIVED COLUMN TRANSFORMATIONS (SECTION 3) -----------------

def derive_billed_flag(last_bill_date_str: str, master_month: str, master_year: str) -> str:
    """
    3.1 Billed Flag:
    - Billed: If LAST_BILL_DATE falls within the selected Master Month and Master Year.
    - Unbilled: Otherwise (blank, invalid, or outside month/year).
    """
    if not last_bill_date_str or pd.isna(last_bill_date_str):
        return "Unbilled"
    s = str(last_bill_date_str).strip().upper()
    if not s or s in ("-", "NAN", "NULL", "NONE"):
        return "Unbilled"
    
    # Common format: '05-OCT-2026' or '05-OCT-26' or '2026-10-05' or '05/10/2026'
    # Check 3-letter month
    m_target_num = f"{int(master_month):02d}"
    m_target_short = [k for k, v in MONTH_SHORT_TO_NUM.items() if v == m_target_num]
    m_short = m_target_short[0] if m_target_short else ""
    y_target = str(master_year).strip()
    y_short = y_target[-2:]

    # Match format like DD-MMM-YYYY (e.g. 05-OCT-2026)
    match_dmy = re.search(r'(\d{1,2})[-/ ]([A-Z]{3})[-/ ](\d{2,4})', s)
    if match_dmy:
        d_mon = match_dmy.group(2)
        d_yr = match_dmy.group(3)
        if len(d_yr) == 2:
            yr_match = (d_yr == y_short)
        else:
            yr_match = (d_yr == y_target)
        if d_mon == m_short and yr_match:
            return "Billed"
        return "Unbilled"
    
    # Match format like YYYY-MM-DD or DD-MM-YYYY
    match_num = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if match_num:
        d_yr = match_num.group(1)
        d_mon = f"{int(match_num.group(2)):02d}"
        if d_yr == y_target and d_mon == m_target_num:
            return "Billed"
        return "Unbilled"

    match_num2 = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', s)
    if match_num2:
        d_mon = f"{int(match_num2.group(2)):02d}"
        d_yr = match_num2.group(3)
        if d_yr == y_target and d_mon == m_target_num:
            return "Billed"
        return "Unbilled"

    return "Unbilled"

def derive_paid_flag(total_pay_amt_val: Any) -> str:
    """
    3.2 Paid Flag:
    - Paid: If TOTAL_PAY_AMT > 0.
    - Unpaid: If TOTAL_PAY_AMT <= 0, blank, or null.
    """
    if total_pay_amt_val is None or pd.isna(total_pay_amt_val):
        return "Unpaid"
    try:
        val = float(total_pay_amt_val)
        return "Paid" if val > 0 else "Unpaid"
    except (ValueError, TypeError):
        return "Unpaid"

def derive_tariff_category(supply_type_val: Any) -> str:
    """
    3.3 Tariff Category:
    Evaluated with strict precedence:
    - LMV-10: Starts with 10 and length is exactly 3
    - LMV-11: Starts with 11 and length is exactly 3
    - LMV-1: Starts with 1, exactly 2 characters, numeric
    - LMV-2: Starts with 2
    - LMV-3: Starts with 3
    - LMV-4: Starts with 4
    - LMV-5: Starts with 5
    - LMV-6: Starts with 6
    - LMV-7: Starts with 7
    - LMV-8: Starts with 8
    - LMV-9: Starts with 9
    - HV-1: Starts with H1 (case-insensitive)
    - HV-2: Starts with H2 (case-insensitive)
    - HV-3: Starts with H3 (case-insensitive)
    - HV-4: Starts with H4 (case-insensitive)
    - Unknown: Otherwise
    """
    if supply_type_val is None or pd.isna(supply_type_val):
        return "Unknown"
    st = str(supply_type_val).strip().upper()
    if not st:
        return "Unknown"
    
    # Precedence rule: check 10 and 11 first
    if st.startswith("10") and len(st) == 3:
        return "LMV-10"
    if st.startswith("11") and len(st) == 3:
        return "LMV-11"
    if st.startswith("1") and len(st) == 2 and st.isdigit():
        return "LMV-1"
    
    if st.startswith("2"):
        return "LMV-2"
    if st.startswith("3"):
        return "LMV-3"
    if st.startswith("4"):
        return "LMV-4"
    if st.startswith("5"):
        return "LMV-5"
    if st.startswith("6"):
        return "LMV-6"
    if st.startswith("7"):
        return "LMV-7"
    if st.startswith("8"):
        return "LMV-8"
    if st.startswith("9"):
        return "LMV-9"
    
    if st.startswith("H1"):
        return "HV-1"
    if st.startswith("H2"):
        return "HV-2"
    if st.startswith("H3"):
        return "HV-3"
    if st.startswith("H4"):
        return "HV-4"
    
    return "Unknown"

def derive_oa_config(mtr_config_ty_cd: Any) -> str:
    """
    3.4 OA Configuration Flag:
    - YES: If MTR_CONFIG_TY_CD contains substring 'OA' (case-insensitive).
    - NO: Otherwise.
    """
    if mtr_config_ty_cd is None or pd.isna(mtr_config_ty_cd):
        return "NO"
    s = str(mtr_config_ty_cd).strip().upper()
    return "YES" if "OA" in s else "NO"

def derive_smart_meter_flag(cmpresys_cd: Any) -> str:
    """
    3.5 Smart Meter Flag:
    - YES: If CMPRESYS_CODE contains substring 'MD' (case-insensitive).
    - NO: Otherwise.
    """
    if cmpresys_cd is None or pd.isna(cmpresys_cd):
        return "NO"
    s = str(cmpresys_cd).strip().upper()
    return "YES" if "MD" in s else "NO"

def convert_load_kw_and_range(load_val: Any, unit_val: Any) -> Tuple[Optional[int], str]:
    """
    3.6 Load Conversion and Load Range:
    Rules:
    - Apply CEIL to original LOAD to obtain next whole number.
    - Identify original unit from LOAD_UNIT:
      - KW: LOAD * 1.0
      - KVA: LOAD * 0.9
      - BHP: LOAD * 0.746
    - Apply CEIL to converted load to obtain next whole number.
    - Do not round off using conventional rounding.
    - Example: 5.2 KVA -> CEIL(5.2)*0.9 = 6*0.9 = 5.4 -> CEIL(5.4) = 6 KW.
    - Example: 10.5 BHP -> CEIL(10.5)*0.746 = 11*0.746 = 8.206 -> CEIL(8.206) = 9 KW.
    - Unknown/Invalid for 0, negative, missing unit, or invalid loads.
    """
    if load_val is None or pd.isna(load_val) or unit_val is None or pd.isna(unit_val):
        return None, "Unknown/Invalid"
    
    try:
        raw_load = float(load_val)
    except (ValueError, TypeError):
        return None, "Unknown/Invalid"
    
    if raw_load <= 0:
        return None, "Unknown/Invalid"
    
    unit = str(unit_val).strip().upper()
    ceil_orig = math.ceil(raw_load)
    
    if unit == "KW":
        converted = float(ceil_orig)
    elif unit == "KVA":
        converted = ceil_orig * 0.9
    elif unit == "BHP":
        converted = ceil_orig * 0.746
    else:
        # Missing or unidentifiable unit
        return None, "Unknown/Invalid"
    
    load_kw = math.ceil(converted)
    
    # Classify LOAD_RANGE
    if 1 <= load_kw <= 4:
        rng = "1-4 KW"
    elif 5 <= load_kw <= 9:
        rng = "5-9 KW"
    elif 10 <= load_kw <= 24:
        rng = "10-24 KW"
    elif 25 <= load_kw <= 50:
        rng = "25-50 KW"
    elif load_kw > 50:
        rng = "Above 50 KW"
    else:
        rng = "Unknown/Invalid"
        
    return load_kw, rng
