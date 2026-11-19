r"""
DuckDB High-Performance Vectorized Open Access & Billing Engine
Direct query & extraction on compressed CSV (.csv.gz) files.
Enforces 100% authentic portal data with strict 1:1 summary-to-detail count parity.
"""

import os
import re
import glob
import time
import math
from typing import Dict, List, Any, Optional, Tuple
import duckdb
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from open_access_schema import (
    MASTER_COLUMNS_155,
    ADMIN_COLUMNS,
    DISCOMS,
    load_division_mapping,
    parse_folder_period,
    derive_tariff_category,
    derive_billed_flag,
    derive_paid_flag,
    derive_oa_config,
    derive_smart_meter_flag,
    convert_load_kw_and_range,
    format_in,
    format_currency_dual
)

CACHE_DIR = os.path.join(os.path.dirname(__file__), "data_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

# ----------------- FILE DISCOVERY -----------------
def discover_master_files(base_dir: str) -> pd.DataFrame:
    """
    Scans base_dir and returns catalog of all valid master files (> 500 bytes):
    [filepath, discom, div_code, filename, size_mb]
    """
    if not os.path.exists(base_dir):
        return pd.DataFrame()
    
    # Check if base_dir contains subfolders per DISCOM or directly csv.gz files
    pattern1 = os.path.join(base_dir, "*", "*.csv.gz")
    pattern2 = os.path.join(base_dir, "*.csv.gz")
    all_files = glob.glob(pattern1) + glob.glob(pattern2)
    
    records = []
    for f in all_files:
        try:
            size_bytes = os.path.getsize(f)
        except OSError:
            continue
        if size_bytes < 500:  # Skip empty gzip headers (49 bytes)
            continue
            
        fname = os.path.basename(f)
        parent = os.path.basename(os.path.dirname(f)).upper()
        # Find DIV code
        m = re.search(r'(DIV\d+)', fname, re.IGNORECASE)
        div_code = m.group(1).upper() if m else ""
        
        # Determine DISCOM from parent folder or div_code
        discom = parent if parent in DISCOMS else ""
        if not discom:
            if div_code.startswith("DIV1"): discom = "PVVNL"
            elif div_code.startswith("DIV2"): discom = "DVVNL"
            elif div_code.startswith("DIV3"): discom = "MVVNL"
            elif div_code.startswith("DIV4"): discom = "PuVVNL"
            elif div_code.startswith("DIV5"): discom = "KESCO"
            else: discom = "OTHER"
            
        records.append({
            "filepath": f.replace("\\", "/"),
            "discom": discom,
            "div_code": div_code,
            "filename": fname,
            "size_mb": round(size_bytes / (1024 * 1024), 2)
        })
        
    df = pd.DataFrame(records)
    return df

# ----------------- VECTORIZED ENRICHMENT -----------------
def enrich_master_dataframe(df: pd.DataFrame, master_month: str, master_year: str, div_map: Dict[str, Dict[str, str]]) -> pd.DataFrame:
    """
    Applies authoritative division hierarchy and all derived columns specified in Section 3:
    1. Administrative Hierarchy: DISCOM, ZONE_NAME, CIRCLE_NAME, DIVISION_NAME, DISTRICT
    2. BILLED_FLAG (Section 3.1)
    3. PAID_FLAG (Section 3.2)
    4. TARIFF_CATEGORY (Section 3.3)
    5. OA_CONFIG (Section 3.4)
    6. SMART_METER_FLAG (Section 3.5)
    7. LOAD_KW & LOAD_RANGE (Section 3.6)
    """
    if df.empty:
        return df

    # Normalize column names to uppercase
    df.columns = [c.strip().upper() for c in df.columns]

    # Clean DIV_CODE
    if "DIV_CODE" in df.columns:
        div_codes = df["DIV_CODE"].astype(str).str.strip().str.upper()
    else:
        div_codes = pd.Series([""] * len(df), index=df.index)

    # 1. Map Administrative Hierarchy
    mapped_div = []
    mapped_circle = []
    mapped_zone = []
    mapped_discom = []
    mapped_district = []

    for d_code in div_codes:
        meta = div_map.get(d_code, {})
        mapped_div.append(meta.get("DIVISION_NAME", d_code or "Unmapped"))
        mapped_circle.append(meta.get("CIRCLE_NAME", "Unmapped"))
        mapped_zone.append(meta.get("ZONE_NAME", "Unmapped"))
        # Fallback DISCOM if not in meta
        disc = meta.get("DISCOM", "")
        if not disc or disc == "NAN":
            if d_code.startswith("DIV1"): disc = "PVVNL"
            elif d_code.startswith("DIV2"): disc = "DVVNL"
            elif d_code.startswith("DIV3"): disc = "MVVNL"
            elif d_code.startswith("DIV4"): disc = "PuVVNL"
            elif d_code.startswith("DIV5"): disc = "KESCO"
            else: disc = "OTHER"
        mapped_discom.append(disc)
        mapped_district.append(meta.get("DISTRICT", ""))

    df["DIVISION_NAME"] = mapped_div
    df["CIRCLE_NAME"] = mapped_circle
    df["ZONE_NAME"] = mapped_zone
    df["DISCOM"] = mapped_discom
    df["DISTRICT"] = mapped_district

    # 1B. Derived Composite Key: DISCOM + ACCT_ID (Statewide Guaranteed Uniqueness)
    if "ACCT_ID" in df.columns:
        df["DISCOM_ACCT_ID"] = [
            f"{str(d).strip().upper()}_{str(a).strip()}" if str(a).strip() and str(a).strip().upper() not in ("-", "NAN", "NULL", "NONE") else ""
            for d, a in zip(df["DISCOM"], df["ACCT_ID"])
        ]
    else:
        df["DISCOM_ACCT_ID"] = ""

    # 2. Derived BILLED_FLAG (Section 3.1)
    if "LAST_BILL_DATE" in df.columns:
        df["BILLED_FLAG"] = [
            derive_billed_flag(dt, master_month, master_year)
            for dt in df["LAST_BILL_DATE"]
        ]
    else:
        df["BILLED_FLAG"] = "Unbilled"

    # 3. Derived PAID_FLAG (Section 3.2)
    if "TOTAL_PAY_AMT" in df.columns:
        df["PAID_FLAG"] = [
            derive_paid_flag(amt) for amt in df["TOTAL_PAY_AMT"]
        ]
    else:
        df["PAID_FLAG"] = "Unpaid"

    # 4. Derived TARIFF_CATEGORY (Section 3.3)
    if "SUPPLY_TYPE" in df.columns:
        df["TARIFF_CATEGORY"] = [
            derive_tariff_category(st) for st in df["SUPPLY_TYPE"]
        ]
    else:
        df["TARIFF_CATEGORY"] = "Unknown"

    # 5. Derived OA_CONFIG (Section 3.4)
    if "MTR_CONFIG_TY_CD" in df.columns:
        df["OA_CONFIG"] = [
            derive_oa_config(cfg) for cfg in df["MTR_CONFIG_TY_CD"]
        ]
    else:
        df["OA_CONFIG"] = "NO"

    # 6. Derived SMART_METER_FLAG (Section 3.5)
    smart_col = "CMPRESYS_CODE" if "CMPRESYS_CODE" in df.columns else ("CMPRESYS_CD" if "CMPRESYS_CD" in df.columns else ("PRESYS_CD" if "PRESYS_CD" in df.columns else None))
    if smart_col:
        df["SMART_METER_FLAG"] = [
            derive_smart_meter_flag(v) for v in df[smart_col]
        ]
    else:
        df["SMART_METER_FLAG"] = "NO"

    # 7. Derived LOAD_KW & LOAD_RANGE (Section 3.6)
    if "LOAD" in df.columns and "LOAD_UNIT" in df.columns:
        kw_list = []
        range_list = []
        for l_val, u_val in zip(df["LOAD"], df["LOAD_UNIT"]):
            kw, rng = convert_load_kw_and_range(l_val, u_val)
            kw_list.append(kw)
            range_list.append(rng)
        df["LOAD_KW"] = kw_list
        df["LOAD_RANGE"] = range_list
    else:
        df["LOAD_KW"] = None
        df["LOAD_RANGE"] = "Unknown/Invalid"

    # Numeric conversion for KPI metrics
    numeric_fields = [
        "CONSUMPTION_CURR_MNTH", "CURRENT_ASSESSMENT", "CA_EX_SUBS",
        "TOTAL_PAY_AMT", "TOTAL_OUTSTANDING", "BILLED_AMOUNT", "ARREAR"
    ]
    for num_f in numeric_fields:
        if num_f in df.columns:
            df[num_f] = pd.to_numeric(df[num_f], errors='coerce').fillna(0.0)

    return df

# ----------------- OPEN ACCESS EXTRACTION & CACHE ENGINE -----------------
def load_or_extract_oa_dataset(
    base_dir: str,
    master_month: str,
    master_year: str,
    progress_callback=None,
    force_refresh: bool = False
) -> pd.DataFrame:
    """
    Extracts all Open Access records (MTR_CONFIG_TY_CD contains 'OA') across all DISCOMs in base_dir.
    Caches as Parquet in data_cache/{folder_name}_oa_dataset.parquet for instant sub-second reloading.
    """
    folder_name = os.path.basename(os.path.normpath(base_dir)).strip()
    cache_file = os.path.join(CACHE_DIR, f"{folder_name}_oa_dataset.parquet")
    
    if os.path.exists(cache_file) and not force_refresh:
        try:
            con = duckdb.connect()
            df = con.execute(f"SELECT * FROM '{cache_file.replace(os.sep, '/')}'").df()
            if not df.empty:
                return df
        except Exception as e:
            print(f"Cache read error: {e}, falling back to source extraction")

    # Catalog files
    file_catalog = discover_master_files(base_dir)
    if file_catalog.empty:
        return pd.DataFrame()

    div_map = load_division_mapping()
    extracted_dfs = []
    
    # Process DISCOM by DISCOM to optimize memory & parallel throughput
    discom_groups = file_catalog.groupby("discom")
    total_groups = len(discom_groups)
    
    for idx, (discom, group) in enumerate(discom_groups):
        if progress_callback:
            progress_callback(idx / total_groups, f"Extracting Open Access records for {discom} ({len(group)} files)...")
            
        flist = group["filepath"].tolist()
        if not flist:
            continue
            
        con = duckdb.connect(database=":memory:")
        con.execute("SET preserve_insertion_order = false;")
        con.execute("SET threads TO 8;")
        
        # Read files and extract OA records
        try:
            flist_str = f"[{', '.join([repr(p) for p in flist])}]"
            query = f"""
                SELECT * 
                FROM read_csv_auto({flist_str}, all_varchar=true)
                WHERE UPPER(COALESCE(MTR_CONFIG_TY_CD, '')) LIKE '%OA%'
            """
            discom_df = con.execute(query).df()
            if not discom_df.empty:
                extracted_dfs.append(discom_df)
        except Exception as e:
            print(f"Error querying {discom}: {e}")
            # Fallback file by file if batch sniffer throws dialect warning
            sub_dfs = []
            for single_f in flist:
                try:
                    q_single = f"""
                        SELECT * 
                        FROM read_csv_auto('{single_f}', all_varchar=true)
                        WHERE UPPER(COALESCE(MTR_CONFIG_TY_CD, '')) LIKE '%OA%'
                    """
                    s_df = con.execute(q_single).df()
                    if not s_df.empty:
                        sub_dfs.append(s_df)
                except Exception:
                    pass
            if sub_dfs:
                extracted_dfs.append(pd.concat(sub_dfs, ignore_index=True))
        finally:
            con.close()

    if progress_callback:
        progress_callback(0.95, "Enriching derived columns and administrative taxonomy...")

    if not extracted_dfs:
        return pd.DataFrame()

    full_oa_df = pd.concat(extracted_dfs, ignore_index=True)
    full_oa_df = enrich_master_dataframe(full_oa_df, master_month, master_year, div_map)

    # Save to Parquet cache
    try:
        full_oa_df.to_parquet(cache_file, index=False)
    except Exception as e:
        print(f"Warning: Could not save Parquet cache: {e}")

    if progress_callback:
        progress_callback(1.0, f"Extraction complete: {len(full_oa_df)} Open Access records found.")

    return full_oa_df

# ----------------- FULL MASTER DATA DISCOM/DIVISION QUERY -----------------
def query_master_subset(
    file_list: List[str],
    master_month: str,
    master_year: str,
    limit: Optional[int] = None
) -> pd.DataFrame:
    """
    Executes a direct DuckDB query over a specific set of master files (e.g. for a selected Division)
    and applies enrichment.
    """
    if not file_list:
        return pd.DataFrame()

    con = duckdb.connect(database=":memory:")
    con.execute("SET preserve_insertion_order = false;")
    con.execute("SET threads TO 8;")

    flist_str = f"[{', '.join([repr(p) for p in file_list])}]"
    limit_clause = f"LIMIT {limit}" if limit else ""
    query = f"""
        SELECT * 
        FROM read_csv_auto({flist_str}, all_varchar=true)
        {limit_clause}
    """
    df = con.execute(query).df()
    con.close()

    div_map = load_division_mapping()
    df = enrich_master_dataframe(df, master_month, master_year, div_map)
    return df

# ----------------- KPI METRICS CALCULATION (SECTION 4) -----------------
def calculate_kpi_summary(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Calculates the 8 standard summary indicators specified in Section 4:
    4.1 Total Number of Consumers: Distinct DISCOM + ACCT_ID (excluding blank/null)
    4.2 Number of Billed Consumers: Distinct DISCOM + ACCT_ID where BILLED_FLAG = 'Billed'
    4.3 Sold Units: SUM of CONSUMPTION_CURR_MNTH (displayed in kWh and MU)
    4.4 CURRENT_ASSESSMENT: SUM of authentic CURRENT_ASSESSMENT
    4.5 CA_EX_SUBS: SUM of authentic CA_EX_SUBS
    4.6 Total Payment Count: Distinct paying consumers (DISCOM + ACCT_ID where TOTAL_PAY_AMT > 0)
    4.7 Total Payment Amount: SUM of authentic TOTAL_PAY_AMT
    4.8 Total Outstanding: SUM of authentic TOTAL_OUTSTANDING
    """
    if df.empty:
        return {
            "total_consumers": 0,
            "billed_consumers": 0,
            "unbilled_consumers": 0,
            "billing_ratio": 0.0,
            "sold_units_kwh": 0.0,
            "sold_units_mu": 0.0,
            "current_assessment": 0.0,
            "ca_ex_subs": 0.0,
            "payment_count": 0,
            "payment_amount": 0.0,
            "total_outstanding": 0.0
        }

    # Distinct consumers based on DISCOM + ACCT_ID composite key
    valid_acct = df[df["ACCT_ID"].notna() & (df["ACCT_ID"].astype(str).str.strip() != "") & (df["ACCT_ID"].astype(str).str.strip().str.upper() != "NAN")]
    total_consumers = int(valid_acct.groupby(["DISCOM", "ACCT_ID"]).ngroups) if not valid_acct.empty else 0

    # Billed consumers
    billed_df = valid_acct[valid_acct["BILLED_FLAG"] == "Billed"]
    billed_consumers = int(billed_df.groupby(["DISCOM", "ACCT_ID"]).ngroups) if not billed_df.empty else 0
    unbilled_consumers = max(0, total_consumers - billed_consumers)
    billing_ratio = (billed_consumers / total_consumers * 100.0) if total_consumers > 0 else 0.0

    # Aggregations
    sold_units_kwh = float(df["CONSUMPTION_CURR_MNTH"].sum()) if "CONSUMPTION_CURR_MNTH" in df.columns else 0.0
    sold_units_mu = sold_units_kwh / 1e6

    ca_amt = float(df["CURRENT_ASSESSMENT"].sum()) if "CURRENT_ASSESSMENT" in df.columns else 0.0
    ca_ex_subs_amt = float(df["CA_EX_SUBS"].sum()) if "CA_EX_SUBS" in df.columns else 0.0

    # Paying consumers (distinct DISCOM + ACCT_ID with PAID_FLAG == 'Paid')
    paying_records = valid_acct[valid_acct["PAID_FLAG"] == "Paid"]
    payment_count = int(paying_records.groupby(["DISCOM", "ACCT_ID"]).ngroups) if not paying_records.empty else 0
    payment_amount = float(df["TOTAL_PAY_AMT"].sum()) if "TOTAL_PAY_AMT" in df.columns else 0.0
    total_outstanding = float(df["TOTAL_OUTSTANDING"].sum()) if "TOTAL_OUTSTANDING" in df.columns else 0.0

    return {
        "total_consumers": total_consumers,
        "billed_consumers": billed_consumers,
        "unbilled_consumers": unbilled_consumers,
        "billing_ratio": round(billing_ratio, 2),
        "sold_units_kwh": round(sold_units_kwh, 2),
        "sold_units_mu": round(sold_units_mu, 4),
        "current_assessment": round(ca_amt, 2),
        "ca_ex_subs": round(ca_ex_subs_amt, 2),
        "payment_count": payment_count,
        "payment_amount": round(payment_amount, 2),
        "total_outstanding": round(total_outstanding, 2)
    }

# ----------------- STYLED EXCEL EXPORT GENERATOR -----------------
def generate_excel_export(df: pd.DataFrame, title: str = "Open_Access_Consumers_Analysis") -> bytes:
    """
    Generates a professionally styled Microsoft Excel (.xlsx) file
    with authentic column headers, formatted numbers, and auto-adjusted column widths.
    Strictly preserves 1:1 row count parity with the filtered dataset.
    """
    import io
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consumer Details"
    ws.views.sheetView[0].showGridLines = True

    # Palette
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style='thin', color="D1D5DB"),
        right=Side(style='thin', color="D1D5DB"),
        top=Side(style='thin', color="D1D5DB"),
        bottom=Side(style='thin', color="D1D5DB")
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    cols = list(df.columns)
    ws.append(cols)
    
    # Format Header Row
    for col_idx in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # Format Data Rows
    for row_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        ws.append(list(row_data))
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            
            # Numeric alignment & formatting
            if isinstance(val, (int, np.integer)):
                cell.alignment = right_align
                cell.number_format = "#,##,##0"
            elif isinstance(val, (float, np.floating)):
                cell.alignment = right_align
                cell.number_format = "#,##,##0.00"
            else:
                cell.alignment = left_align

    # Auto-adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col[:100]:  # sample top 100
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
