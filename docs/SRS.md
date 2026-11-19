# Software Requirements Specification (SRS)
## Open Access Consumers & Billing Revenue Analysis Dashboard

**Document Version:** 3.0.0  
**Effective Date:** November 2026  
**System Name:** Open Access & Billing Revenue Analysis System  
**Deployment Target:** UPPCL Electricity Distribution Companies (DVVNL, KESCO, MVVNL, PuVVNL, PVVNL)  
**Compliance Standards:** Universal Agent Rules (`global.md`), Strict Zero Synthetic Data Policy, Rule #6 (1:1 Parity), Rule #7 (Excel AutoFilter Semantics Above Summaries across all 22 Dimensions), Rule #8 (Indian Number Formatting)

---

### 1. Executive Summary & Objective

The **Open Access Consumers & Billing Revenue Analysis Dashboard** is an enterprise-grade analytical system developed to ingest, process, and evaluate monthly master billing registers across Uttar Pradesh electricity distribution companies. The system extracts and analyzes consumer-level billing status, energy consumption, revenue assessment, realization collections, and outstanding dues using authoritative master data maintained in monthly archives.

Key architectural tenets:
- **Rule #7 Compliance (22 Contextual AutoFilters Above Summaries):** Clean horizontal filter bar positioned directly above the KPI summary cards with popovers, top search input, master selection checkboxes, and persistent state across all 22 authentic dimensions.
- **Rule #6 Compliance (Summary Hyperlinks & Strict 1:1 Parity):** Every displayed summary metric links directly to the detailed consumer dossier, preserving exact 1:1 row count parity.
- **Rule #8 Compliance (Indian Number Formatting):** Indian comma grouping (`1,00,000`, `10,00,000`, `1,00,00,000`) applied universally across cards, tables, tooltips, and exports.
- **High-Performance DuckDB Vectorized Querying:** Direct querying over compressed `.csv.gz` archives with Parquet caching for sub-second responsiveness.
- **Automated Period Extraction:** Folder name extraction of Master Month and Master Year with manual override controls.

---

### 2. Complete Inventory of All 22 Filters (Sections 5.1 & 5.2)

All 22 filters are implemented as Excel AutoFilter popovers directly above the KPI summaries:

#### 2.1 Band 1: Administrative Hierarchy (Cascading Filters 1 to 4)
1. **DISCOM:** `DISCOM` (DVVNL, KESCO, MVVNL, PuVVNL, PVVNL)
2. **ZONE:** `ZONE_NAME` (Dynamically cascaded from selected DISCOMs)
3. **CIRCLE:** `CIRCLE_NAME` (Dynamically cascaded from selected Zones)
4. **DIVISION:** `DIVISION_NAME` (Dynamically cascaded from selected Circles)

#### 2.2 Band 2: Core Operational, Billing & Technical Flags (Filters 5 to 11)
5. **OPR_FLAG:** `OPR_FLG` (`Y`, `N`, etc.)
6. **Tariff Category:** `TARIFF_CATEGORY` (`LMV-1` through `LMV-11`, `HV-1` through `HV-4`, `Unknown`)
7. **Billed Flag:** `BILLED_FLAG` (`Billed`, `Unbilled`)
8. **Paid Flag:** `PAID_FLAG` (`Paid`, `Unpaid`)
9. **OA Config:** `OA_CONFIG` (`YES`, `NO`)
10. **Smart Meter Flag:** `SMART_METER_FLAG` (`YES`, `NO`)
11. **Load Range:** `LOAD_RANGE` (`1-4 KW`, `5-9 KW`, `10-24 KW`, `25-50 KW`, `Above 50 KW`, `Unknown/Invalid`)

#### 2.3 Band 3: Metering, Reading & Connection Attributes (Filters 12 to 22)
12. **SUPPLY_TYPE:** Authentic connection supply tariff code (e.g. `HV-1`, `HV-2`, `10`, `11`, etc.)
13. **TYPE_of_Connection:** `CONNECTION_TYPE` (e.g. `POSTPAID`, `PREPAID`)
14. **BILL_CYC_CD:** Billing cycle code (e.g. `HV`, `MRI`, `CYCLE`)
15. **MTR_TYPE_CD:** Meter type code (e.g. `HTCT/PTS`, `LTCT`)
16. **MTR_CONFIG_TY_CD:** Meter configuration code (e.g. `OA-HTNEW-TOD`)
17. **MTR_MAKE:** Meter manufacturer
18. **MR_SOURCE_CD:** Meter reading source code
19. **METER_READ_REMARK:** Meter reader remark code
20. **BILL_BASIS:** Billing basis (e.g. `OK`, `MU`, `NA`)
21. **INF_BILL:** Informative bill flag (`Y`, `N`)
22. **PRESYS_CD:** Pre-system / communication protocol (`CMPRESYS_CODE`)

---

### 3. Summary Indicators & Drilldown (Rule #6 & Rule #8)

- **Total Consumers:** Distinct count of valid `ACCT_ID` in the active scope.
- **Billed Consumers:** Distinct count of `ACCT_ID` with `BILLED_FLAG = 'Billed'`.
- **Unbilled Consumers:** Distinct count of `ACCT_ID` with `BILLED_FLAG = 'Unbilled'`.
- **Sold Units:** Authentic sum of `CONSUMPTION_CURR_MNTH` in kWh and Million Units (MU).
- **Current Assessment:** Authentic sum of `CURRENT_ASSESSMENT` in Indian currency (`₹`).
- **CA Excluding Subsidy:** Authentic sum of `CA_EX_SUBS` in Indian currency (`₹`).
- **Total Payment Realization:** Authentic sum of `TOTAL_PAY_AMT` in Indian currency (`₹`) and count of paying consumers.
- **Total Outstanding:** Authentic sum of `TOTAL_OUTSTANDING` in Indian currency (`₹`).
- **Clickable Drilldown:** Clicking any summary card highlights the card and isolates the consumer dossier table to that exact cohort with strict 1:1 count parity.

---

### 4. Traceability & Acceptance Matrix

| Requirement ID | Description | Implementation File | Status |
|---|---|---|---|
| REQ-01 | Folder Browser & Period Extraction | `app.py`, `open_access_schema.py` | Verified |
| REQ-02 | Derived Columns (Billed, Paid, Tariff, OA, Smart Meter, Load) | `open_access_schema.py`, `duckdb_oa_engine.py` | Verified |
| REQ-03 | All 22 AutoFilter Popovers Above Summaries (Rule #7) | `app.py` | Verified |
| REQ-04 | 8 Clickable KPI Indicators (Rule #6 Drilldown) | `app.py`, `duckdb_oa_engine.py` | Verified |
| REQ-05 | 4 Graphical Analytical Suites | `app.py` | Verified |
| REQ-06 | Consumer Dossier Table with Search & Pagination | `app.py` | Verified |
| REQ-07 | Styled Excel & CSV Export with Strict 1:1 Count Parity | `duckdb_oa_engine.py`, `app.py` | Verified |
| REQ-08 | Indian Number Formatting (Rule #8) | `open_access_schema.py` | Verified |
| REQ-09 | Pure Authentic Portal Data (Rule #4) | Entire System | Verified |
