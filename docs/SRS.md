# Software Requirements Specification (SRS)
## Open Access Consumers & Billing Revenue Analysis Dashboard

**Document Version:** 1.0.0  
**Effective Date:** November 2026  
**System Name:** Open Access & Billing Revenue Analysis System  
**Deployment Target:** UPPCL Electricity Distribution Companies (DVVNL, KESCO, MVVNL, PuVVNL, PVVNL)  
**Data Policy:** Pure Authentic Portal Data (Strict Zero Synthetic Data Policy)

---

### 1. Executive Summary & Objective

The **Open Access Consumers & Billing Revenue Analysis Dashboard** is an enterprise-grade analytical system developed to ingest, process, and evaluate monthly master billing registers across Uttar Pradesh electricity distribution companies. The system extracts and analyzes consumer-level billing status, energy consumption, revenue assessment, realization collections, and outstanding dues using authoritative master data maintained in monthly archives.

Key goals:
- Dynamic, generic, and configuration-driven processing without hardcoded file paths, DISCOM names, or fixed column offsets.
- High-performance vectorized DuckDB querying directly over compressed `.csv.gz` archives.
- Automated extraction of Master Month and Master Year from directory names with editable overrides.
- 100% adherence to authentic portal data with strict 1:1 count parity between summary indicators and downloadable consumer dossiers.

---

### 2. Functional Requirements Specification

#### 2.1 Input Controls & Master Period Resolution
- **FR-1.1 Folder Selection:** Provide a directory path selector with both text input and a Windows native folder browser (`tkinter.filedialog.askdirectory`) to select the folder containing monthly master files (e.g., `D:\Master Data\102026`).
- **FR-1.2 Master Month Extraction:** Automatically parse the first two digits of the selected folder name, map to calendar months (01–12), and provide manual correction controls.
- **FR-1.3 Master Year Extraction:** Automatically extract the last four digits of the selected folder name, set as the reference year, and provide manual correction controls.
- **FR-1.4 Reference Period Authority:** The Master Month and Master Year serve as the reference period for all billing-related evaluations.

#### 2.2 Data Transformation & Derived Attributes
- **FR-2.1 Billed Flag (`BILLED_FLAG`):**
  - `Billed`: If `LAST_BILL_DATE` falls within the selected Master Month and Master Year.
  - `Unbilled`: If `LAST_BILL_DATE` is blank, invalid, or outside the selected Master Month and Year.
- **FR-2.2 Paid Flag (`PAID_FLAG`):**
  - `Paid`: If `TOTAL_PAY_AMT > 0`.
  - `Unpaid`: If `TOTAL_PAY_AMT` is zero, blank, or null.
- **FR-2.3 Tariff Category (`TARIFF_CATEGORY`):** Derived from `SUPPLY_TYPE` with strict evaluation precedence:
  - `LMV-10`: Starts with `10` and string length is exactly 3.
  - `LMV-11`: Starts with `11` and string length is exactly 3.
  - `LMV-1`: Starts with `1`, length is exactly 2, and numeric.
  - `LMV-2` to `LMV-9`: Starts with digits `2` through `9`.
  - `HV-1` to `HV-4`: Starts with `H1` through `H4` (case-insensitive).
  - `Unknown`: Any unmatched or unidentifiable supply code.
- **FR-2.4 Open Access Configuration Flag (`OA_CONFIG`):**
  - `YES`: If `MTR_CONFIG_TY_CD` contains substring `OA` (case-insensitive).
  - `NO`: Otherwise.
- **FR-2.5 Smart Meter Flag (`SMART_METER_FLAG`):**
  - `YES`: If `CMPRESYS_CODE` (or `CMPRESYS_CD` / `PRESYS_CD`) contains substring `MD` (case-insensitive).
  - `NO`: Otherwise.
- **FR-2.6 Load Conversion (`LOAD_KW`) & Load Range (`LOAD_RANGE`):**
  - Ceiling rule applied to original `LOAD`: `CEIL(LOAD)`.
  - Unit conversion: $\text{KW} = \lceil\text{LOAD}\rceil$, $\text{KVA} = \lceil\lceil\text{LOAD}\rceil \times 0.9\rceil$, $\text{BHP} = \lceil\lceil\text{LOAD}\rceil \times 0.746\rceil$.
  - Range classifications: `1-4 KW`, `5-9 KW`, `10-24 KW`, `25-50 KW`, `Above 50 KW`, and `Unknown/Invalid`.

#### 2.3 Dashboard Summary Indicators
- **FR-3.1 Total Consumers:** Distinct count of valid `ACCT_ID` in the active scope.
- **FR-3.2 Billed Consumers:** Distinct count of `ACCT_ID` with `BILLED_FLAG = 'Billed'`.
- **FR-3.3 Sold Units:** Authentic sum of `CONSUMPTION_CURR_MNTH` displayed in kWh and Million Units (MU).
- **FR-3.4 Current Assessment:** Authentic sum of `CURRENT_ASSESSMENT` in Indian currency (`₹`).
- **FR-3.5 CA Excluding Subsidy:** Authentic sum of `CA_EX_SUBS` in Indian currency (`₹`).
- **FR-3.6 Total Payment Count:** Count of consumers with realized payments (`TOTAL_PAY_AMT > 0`).
- **FR-3.7 Total Payment Amount:** Authentic sum of `TOTAL_PAY_AMT` in Indian currency (`₹`).
- **FR-3.8 Total Outstanding:** Authentic sum of `TOTAL_OUTSTANDING` in Indian currency (`₹`).

#### 2.4 Multi-Dimensional Cascading Filters (Excel AutoFilter Semantics)
- Cascading administrative hierarchy: `DISCOM` $\rightarrow$ `ZONE` $\rightarrow$ `CIRCLE` $\rightarrow$ `DIVISION` mapped using authoritative `UPPCL Division Code Mapping.xlsx`.
- 18 Operational and technical filters: `OPR_FLG`, `TARIFF_CATEGORY`, `SUPPLY_TYPE`, `CONNECTION_TYPE`, `BILL_CYC_CD`, `MTR_TYPE_CD`, `MTR_CONFIG_TY_CD`, `MTR_MAKE`, `MR_SOURCE_CD`, `METER_READ_REMARK`, `BILL_BASIS`, `INF_BILL`, `CMPRESYS_CODE`, `OA_CONFIG`, `LOAD_RANGE`, `BILLED_FLAG`, `PAID_FLAG`, and `SMART_METER_FLAG`.
- Omni-directional cascading: Narrowing any filter updates available options in other filters.

#### 2.5 Visualizations & Analytics
- **Billing Analysis:** Donut chart of Billed vs Unbilled proportions, DISCOM-wise billing bars, and tariff-wise breakdowns.
- **Energy Consumption:** Sold units by DISCOM, by load range, and by tariff category.
- **Revenue & Realization:** Current Assessment vs Realized Payments, and total outstanding dues by DISCOM.
- **Smart Meter & OA Analysis:** Proportions of smart-metered vs non-smart consumers and OA configuration distributions.

#### 2.6 Consumer-Level Dossier Table & Parity Export
- Real-time search across `ACCT_ID`, `NAME`, `DIVISION_NAME`, and meter serial numbers.
- Column view presets (Standard Executive View, Billing & Financial, Technical & Metering, All Columns).
- Excel (`.xlsx`) export with navy headers, thin gridlines, and Indian number formatting.
- CSV export for downstream ETL pipelines.
- **Strict 1:1 Count Parity:** Every exported file matches the filtered record count exactly.

---

### 3. Non-Functional Requirements

- **Performance:** Sub-second filter latency achieved through DuckDB vectorized query execution and in-memory Parquet caching.
- **Reliability:** Graceful handling of empty `.csv.gz` files, null fields, malformed dates, and missing columns without throwing uncaught exceptions.
- **Formatting Standards:** Indian numbering system (`en-IN` format: `1,00,000`, `10,00,000`, `1,00,00,000`) applied to all numerical figures.
- **Authenticity:** Zero synthetic data policy strictly observed.

---

### 4. Traceability & Acceptance Matrix

| Requirement ID | Description | Implementation File | Status |
|---|---|---|---|
| REQ-01 | Folder Browser & Period Extraction | `app.py`, `open_access_schema.py` | Verified |
| REQ-02 | Derived Columns (Billed, Paid, Tariff, OA, Smart Meter, Load) | `open_access_schema.py`, `duckdb_oa_engine.py` | Verified |
| REQ-03 | 8 KPI Indicators with Clickable Drilldown | `app.py`, `duckdb_oa_engine.py` | Verified |
| REQ-04 | Cascading Administrative & Operational Filters | `app.py` | Verified |
| REQ-05 | 4 Graphical Analytical Suites | `app.py` | Verified |
| REQ-06 | Consumer Dossier Table with Search & Pagination | `app.py` | Verified |
| REQ-07 | Styled Excel & CSV Export with 1:1 Count Parity | `duckdb_oa_engine.py`, `app.py` | Verified |
| REQ-08 | Indian Number Formatting (Rule #8) | `open_access_schema.py` | Verified |
| REQ-09 | Pure Authentic Portal Data (Rule #4) | Entire System | Verified |
