# Software Requirements Specification (SRS)
## Open Access Consumers & Billing Revenue Analysis Dashboard

**Document Version:** 4.0.0  
**Effective Date:** November 2026  
**System Name:** Open Access & Billing Revenue Analysis System  
**Deployment Target:** UPPCL Electricity Distribution Companies (DVVNL, KESCO, MVVNL, PuVVNL, PVVNL)  
**Compliance Standards:** Universal Agent Rules (`global.md`), Strict Zero Synthetic Data Policy, Rule #6 (1:1 Parity), Rule #7 (Exact Excel AutoFilter Semantics Above Summaries), Rule #8 (Indian Number Formatting)

---

### 1. Executive Summary & Objective

The **Open Access Consumers & Billing Revenue Analysis Dashboard** is an enterprise-grade analytical system developed to ingest, process, and evaluate monthly master billing registers across Uttar Pradesh electricity distribution companies. The system extracts and analyzes consumer-level billing status, energy consumption, revenue assessment, realization collections, and outstanding dues using authoritative master data maintained in monthly archives.

Key architectural tenets:
- **Administrative Hierarchy Summary Presentation:** Real-time administrative summary breakdown table presenting data down to the user's selected hierarchy level (DISCOM, Zone, Circle, or Division) with all upper-level parent hierarchy tiers displayed and aggregated, 100% mathematical reconciliation against KPI indicators, Indian number formatting, and dedicated executive export.
- **Full 167-Column Report Export by Default:** Excel (.xlsx) and CSV (.csv) reports export all 167 columns (155 authentic portal master register fields + administrative mappings + derived fields) by default, preserving strict 1:1 row count parity.
- **Composite Consumer Uniqueness (`DISCOM + ACCT_ID`):** Standalone `ACCT_ID` is not considered unique statewide. Instead, universal consumer uniqueness is strictly established via the composite combination of `DISCOM + ACCT_ID` (`DISCOM_ACCT_ID`), applied consistently across all KPI counters, aggregations, charts, and dossiers.
- **OA Config & OPR_FLAG Pre-selection:** Initial portal load targets active Open Access consumers, displaying `5. OPR_FLAG` pre-selected to `['Y']` and `9. OA Config` pre-selected to `['YES']`, with instant zero-latency widget synchronization across all 22 filters.
- **Rule #7 Compliance (22 Contextual AutoFilters Above Summaries):** Clean horizontal filter bar positioned directly above the KPI summary cards with popovers, top search input, persistent committed base snapshot, manual search selection with explicit `Add current selection to filter` branching, master selection checkboxes, and persistent state across all 22 authentic dimensions.
- **Rule #6 Compliance (Summary Hyperlinks & Strict 1:1 Parity):** Every displayed summary metric links directly to the detailed consumer dossier, preserving exact 1:1 row count parity.
- **Rule #8 Compliance (Indian Number Formatting):** Indian comma grouping (`1,00,000`, `10,00,000`, `1,00,00,000`) applied universally across cards, tables, tooltips, and exports.
- **Rule #5 Compliance (Global Rules Synchronization):** `.gemini/rules/global.md` maintained inside the workspace repository.
- **Rule #2 Compliance (Automated Git Tracking):** Clean Git version control on branch `main` configured for GitHub remote push.
- **High-Performance DuckDB Vectorized Querying:** Direct querying over compressed `.csv.gz` archives with Parquet caching for sub-second responsiveness.

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

### 3. Traceability & Acceptance Matrix

| Rule / Requirement | Description | Implementation File | Status |
|---|---|---|---|
| Rule #1 | Mandatory Confirmation Before Codebase Changes | Agent Workflow | Verified |
| Rule #2 | Automatic Git Add, Commit, and Push | `git` on branch `main` | Configured |
| Rule #3 | Mandatory 6 Structured Response Sections | System Responses | Verified |
| Rule #4 | Strict Zero Synthetic Data Policy | Entire System | Verified |
| Rule #5 | Automatic Synchronization of Global Rules | `.gemini/rules/global.md` | Synchronized |
| Rule #6 | Summary Hyperlinks and Strict 1:1 Data Parity | `app.py`, `duckdb_oa_engine.py` | Verified |
| Rule #7 | Contextual Dropdown Filters with Checkboxes Above Summaries | `app.py` | Verified |
| Rule #8 | Indian Number Formatting (`en-IN`) | `open_access_schema.py` | Verified |
| Rule #9 | Avoid Hardcoding & Configurable Architecture | Entire Architecture | Verified |
| Rule #10 | Validation, Testing, and Regression Prevention | Unit & Compilation Tests | Verified |
| Rule #11 | Mandatory Working Procedure & Scope/Priority | Systematic Process | Verified |
| Rule #12 | High-Performance Omnidirectional Cascading Filters | `duckdb_oa_engine.py` | Verified |
