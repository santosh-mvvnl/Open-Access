# Open Access Consumers & Billing Revenue Analysis Dashboard

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![DuckDB](https://img.shields.io/badge/DuckDB-Vectorized%20SQL-yellow.svg)](https://duckdb.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-Dashboard-red.svg)](https://streamlit.io/)
[![Policy](https://img.shields.io/badge/Data%20Policy-Pure%20Authentic%20Portal%20Data-brightgreen.svg)]()

Interactive high-performance dashboard for ingesting and analyzing monthly master billing registers across Uttar Pradesh electricity distribution companies (**DVVNL**, **KESCO**, **MVVNL**, **PuVVNL**, and **PVVNL**).

---

## ⚡ Key Features

1. **Dynamic Master Data Ingestion:**
   - Folder selector with Windows native Folder Picker and direct path text input.
   - Automatic extraction of Master Month (first two digits) and Master Year (last four digits) with manual override controls.
2. **Derived Data Columns:**
   - `BILLED_FLAG`: Billed / Unbilled based on whether `LAST_BILL_DATE` falls in the Master Month & Year.
   - `PAID_FLAG`: Paid / Unpaid based on `TOTAL_PAY_AMT > 0`.
   - `TARIFF_CATEGORY`: Evaluates LMV-10 & LMV-11 first, then LMV-1 through LMV-9, HV-1 through HV-4, and Unknown.
   - `OA_CONFIG`: YES / NO if `MTR_CONFIG_TY_CD` contains substring `OA`.
   - `SMART_METER_FLAG`: YES / NO if `CMPRESYS_CODE` contains substring `MD`.
   - `LOAD_KW` & `LOAD_RANGE`: Strict conversion with CEIL rule ($1\text{ KW}=1.0$, $1\text{ KVA}=0.9$, $1\text{ BHP}=0.746$) binned into standard ranges.
3. **Summary KPI Indicators:**
   - Total Consumers, Billed Consumers (with Billing %), Sold Units (in MU and kWh), Current Assessment (₹), CA Excluding Subsidy (₹), Payment Count, Realized Payment Amount (₹), and Total Outstanding Dues (₹).
   - Strict 1:1 count parity between summary counts and drilldown records.
4. **Cascading Filters (Excel AutoFilter Semantics):**
   - Cascading administrative hierarchy (`DISCOM` $\rightarrow$ `ZONE` $\rightarrow$ `CIRCLE` $\rightarrow$ `DIVISION`) joined with authoritative division mapping (`D:\Division Name Mapping\UPPCL Division Code Mapping.xlsx`).
   - 18 Operational and technical filters.
5. **Interactive Visualizations:**
   - Billing analysis, energy consumption breakdowns, revenue/realization comparisons, and Smart Meter / Open Access distribution charts.
6. **Consumer-Level Records Table & Styled Export:**
   - Search across consumers, pagination, column visibility presets, and styled Microsoft Excel (`.xlsx`) export.

---

## 🚀 Quick Start

### 1. Launch with One-Click Batch File
Double-click:
```powershell
run_dashboard.bat
```

### 2. Launch via Terminal
```powershell
python -m streamlit run app.py --server.port 8503
```
Open `http://localhost:8503` in your web browser.

---

## 📁 Repository Structure

```
d:\AG Project\Open Access\
├── app.py                             # Streamlit interactive application
├── duckdb_oa_engine.py                # Vectorized DuckDB extraction & caching engine
├── open_access_schema.py              # Authentic schema, period parser, and derived columns
├── requirements.txt                   # Project dependencies
├── run_dashboard.bat                  # One-click startup launcher
├── README.md                          # Project documentation
├── .gitignore                         # Git ignore configuration
└── docs/
    ├── SRS.md                         # Software Requirements Specification
    └── TECHNICAL_DESIGN_AND_MANUAL.md # Technical Design & Operating Manual
```
