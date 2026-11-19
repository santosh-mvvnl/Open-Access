r"""
Open Access Consumers & Billing Revenue Analysis Dashboard
A high-performance, interactive dashboard powered by DuckDB vectorized querying over UPPCL Master Data.
Adheres strictly to the Pure Authentic Portal Data Policy and Universal Agent Rules.
"""

import os
import re
import glob
import time
import math
from datetime import datetime
from typing import Dict, List, Any, Optional

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

from open_access_schema import (
    MASTER_COLUMNS_155,
    ADMIN_COLUMNS,
    DISCOMS,
    MONTH_NAMES,
    MONTH_SHORT_TO_NUM,
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

from duckdb_oa_engine import (
    discover_master_files,
    enrich_master_dataframe,
    load_or_extract_oa_dataset,
    calculate_kpi_summary,
    generate_excel_export,
    CACHE_DIR
)

# ----------------- PAGE CONFIGURATION -----------------
st.set_page_config(
    page_title="Open Access & Billing Revenue Analysis",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for executive aesthetic, KPI cards, and table styling
st.markdown("""
<style>
    /* Executive Card Styling */
    .kpi-container {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        margin-bottom: 20px;
    }
    .kpi-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border-radius: 10px;
        border: 1px solid #334155;
        padding: 16px 18px;
        color: #F8FAFC;
        flex: 1 1 calc(25% - 12px);
        min-width: 220px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
        transition: transform 0.15s ease, border-color 0.15s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        border-color: #38BDF8;
    }
    .kpi-title {
        font-size: 0.82rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #94A3B8;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 1.55rem;
        font-weight: 700;
        color: #38BDF8;
        line-height: 1.2;
    }
    .kpi-sub {
        font-size: 0.78rem;
        color: #CBD5E1;
        margin-top: 5px;
    }
    .highlight-card {
        border-left: 4px solid #38BDF8;
    }
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 9999px;
        font-size: 0.78rem;
        font-weight: 600;
    }
    .badge-success { background-color: #065F46; color: #D1FAE5; }
    .badge-warning { background-color: #854D0E; color: #FEF3C7; }
    .badge-info { background-color: #1E40AF; color: #DBEAFE; }
    
    /* Strict Table Styling */
    .dataframe th {
        background-color: #1E3A8A !important;
        color: white !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- SESSION STATE INITIALIZATION -----------------
if "folder_path" not in st.session_state:
    st.session_state.folder_path = r"D:\Master Data\102026"
if "drilldown_filter" not in st.session_state:
    st.session_state.drilldown_filter = "All"
if "oa_df" not in st.session_state:
    st.session_state.oa_df = None
if "last_loaded_folder" not in st.session_state:
    st.session_state.last_loaded_folder = ""

# ----------------- FOLDER BROWSER HELPER -----------------
def browse_folder_dialog():
    """Opens native Windows folder selection dialog."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        selected = filedialog.askdirectory(
            initialdir=r"D:\Master Data",
            title="Select Monthly Master Data Folder"
        )
        root.destroy()
        if selected:
            st.session_state.folder_path = os.path.normpath(selected)
            st.rerun()
    except Exception as e:
        st.sidebar.error(f"Could not launch native folder picker: {e}")

# ----------------- HEADER & INPUT CONTROLS -----------------
st.markdown("## ⚡ Open Access Consumers & Billing Revenue Analysis Dashboard")
st.markdown("<p style='color: #64748B; margin-top: -12px;'>Authoritative monthly analysis of consumer billing, energy consumption, revenue collection, and outstanding dues across Uttar Pradesh electricity distribution companies.</p>", unsafe_allow_html=True)

# Top Input Controls Box
with st.container():
    col_input1, col_input2, col_input3, col_input4 = st.columns([4, 1.8, 1.8, 2.4])
    
    with col_input1:
        # Detect available month folders under D:\Master Data
        detected_folders = []
        if os.path.exists(r"D:\Master Data"):
            subdirs = [f for f in glob.glob(r"D:\Master Data\*") if os.path.isdir(f)]
            subdirs.sort(reverse=True)
            detected_folders = subdirs
            
        folder_selection = st.text_input(
            "📁 Master Data Folder Path:",
            value=st.session_state.folder_path,
            placeholder=r"e.g. D:\Master Data\102026",
            help="Directory containing monthly master billing register files (.csv.gz)"
        )
        if folder_selection != st.session_state.folder_path:
            st.session_state.folder_path = folder_selection

    with col_input2:
        st.write("")
        st.write("")
        if st.button("📂 Browse Folder", use_container_width=True, help="Open native Windows Folder Picker"):
            browse_folder_dialog()

    # Automatically extract month & year from folder name
    period_info = parse_folder_period(st.session_state.folder_path)
    
    with col_input3:
        # Detected Month Display with manual override in sidebar
        st.metric(
            label="Master Month",
            value=f"{period_info['month_name']} ({period_info['month']})",
            help="Extracted from first 2 digits of folder name"
        )

    with col_input4:
        st.metric(
            label="Master Year",
            value=period_info['year'],
            help="Extracted from last 4 digits of folder name"
        )

# ----------------- SIDEBAR CONFIGURATION & FILTERS -----------------
st.sidebar.header("⚙️ Period & Scope Settings")

# Detected period manual override controls (Section 2.2 & 2.3)
with st.sidebar.expander("🛠️ Manual Period Override", expanded=False):
    override_month = st.selectbox(
        "Master Month (Override):",
        options=[f"{i:02d}" for i in range(1, 13)],
        index=int(period_info['month']) - 1,
        format_func=lambda m: f"{m} - {MONTH_NAMES.get(m, '')}"
    )
    override_year = st.text_input(
        "Master Year (Override):",
        value=period_info['year']
    )
    st.caption("The Master Month & Year serve as the reference period for BILLED_FLAG calculations.")

effective_month = override_month
effective_year = override_year.strip()

# Analysis Mode Selector
st.sidebar.markdown("---")
st.sidebar.subheader("🎯 Analysis Mode")
analysis_mode = st.sidebar.radio(
    "Target Scope:",
    [
        "⚡ Open Access Consumers (OPR=Y & OA Substring)",
        "🌐 Full Master Dataset (All Consumers)"
    ],
    index=0,
    help="Toggle between specialized Open Access analysis and full consumer master dataset."
)
is_oa_mode = "Open Access" in analysis_mode

# Load / Refresh Button
load_col1, load_col2 = st.sidebar.columns(2)
with load_col1:
    btn_load = st.button("🚀 Load Data", use_container_width=True, type="primary")
with load_col2:
    btn_refresh = st.button("🔄 Force Refresh", use_container_width=True, help="Re-scan and rebuild cache")

# Data Loading Logic
should_load = btn_load or btn_refresh or (st.session_state.oa_df is None) or (st.session_state.last_loaded_folder != st.session_state.folder_path)

if should_load:
    if not os.path.exists(st.session_state.folder_path):
        st.error(f"Selected folder does not exist: `{st.session_state.folder_path}`. Please verify the path.")
    else:
        with st.spinner(f"Processing Master Data for {MONTH_NAMES.get(effective_month, '')} {effective_year}..."):
            t_start = time.time()
            progress_bar = st.progress(0.0)
            status_text = st.empty()

            def p_cb(frac, text):
                progress_bar.progress(frac)
                status_text.text(text)

            df_loaded = load_or_extract_oa_dataset(
                base_dir=st.session_state.folder_path,
                master_month=effective_month,
                master_year=effective_year,
                progress_callback=p_cb,
                force_refresh=btn_refresh
            )
            progress_bar.empty()
            status_text.empty()
            t_elapsed = time.time() - t_start

            if df_loaded.empty:
                st.warning("No qualifying records found in the selected folder.")
            else:
                st.session_state.oa_df = df_loaded
                st.session_state.last_loaded_folder = st.session_state.folder_path
                st.sidebar.success(f"Loaded {len(df_loaded)} records in {t_elapsed:.2f}s!")

df_raw = st.session_state.oa_df

if df_raw is None or df_raw.empty:
    st.info("👆 Please click **'🚀 Load Data'** in the sidebar to load and analyze Master Data.")
    st.stop()

# ----------------- CASCADING FILTERS (EXCEL AUTOFILTER SEMANTICS) -----------------
st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Interactive Cascading Filters")

# Reset Filters Button
if st.sidebar.button("🔄 Reset All Filters", use_container_width=True):
    st.session_state.drilldown_filter = "All"
    for k in list(st.session_state.keys()):
        if k.startswith("flt_"):
            del st.session_state[k]
    st.rerun()

# Filter Storage Helper
def get_filter_selection(key: str, default: List[Any]) -> List[Any]:
    return st.session_state.get(f"flt_{key}", default)

# 1. Administrative Hierarchy Cascading (DISCOM -> ZONE -> CIRCLE -> DIVISION)
with st.sidebar.expander("🏛️ Administrative Hierarchy", expanded=True):
    # Available DISCOMs
    avail_discoms = sorted([d for d in df_raw["DISCOM"].dropna().unique() if str(d).strip() != ""])
    sel_discoms = st.multiselect(
        "DISCOM:",
        options=avail_discoms,
        default=get_filter_selection("discom", avail_discoms),
        key="flt_discom"
    )

    # Cascaded Zones based on selected DISCOMs
    df_step1 = df_raw[df_raw["DISCOM"].isin(sel_discoms)] if sel_discoms else df_raw
    avail_zones = sorted([z for z in df_step1["ZONE_NAME"].dropna().unique() if str(z).strip() != ""])
    sel_zones = st.multiselect(
        "Zone:",
        options=avail_zones,
        default=[z for z in get_filter_selection("zone", avail_zones) if z in avail_zones],
        key="flt_zone"
    )

    # Cascaded Circles based on selected Zones
    df_step2 = df_step1[df_step1["ZONE_NAME"].isin(sel_zones)] if sel_zones else df_step1
    avail_circles = sorted([c for c in df_step2["CIRCLE_NAME"].dropna().unique() if str(c).strip() != ""])
    sel_circles = st.multiselect(
        "Circle:",
        options=avail_circles,
        default=[c for c in get_filter_selection("circle", avail_circles) if c in avail_circles],
        key="flt_circle"
    )

    # Cascaded Divisions based on selected Circles
    df_step3 = df_step2[df_step2["CIRCLE_NAME"].isin(sel_circles)] if sel_circles else df_step2
    avail_divs = sorted([d for d in df_step3["DIVISION_NAME"].dropna().unique() if str(d).strip() != ""])
    sel_divs = st.multiselect(
        "Division:",
        options=avail_divs,
        default=[d for d in get_filter_selection("div", avail_divs) if d in avail_divs],
        key="flt_div"
    )

# 2. Operational & Consumer Filters
with st.sidebar.expander("⚡ Operational & Consumer Filters", expanded=True):
    # Base filter subset
    df_admin = df_step3[df_step3["DIVISION_NAME"].isin(sel_divs)] if sel_divs else df_step3

    # OPR_FLG
    avail_opr = sorted([str(v) for v in df_admin["OPR_FLG"].dropna().unique() if str(v).strip() != ""])
    default_opr = ["Y"] if "Y" in avail_opr and is_oa_mode else avail_opr
    sel_opr = st.multiselect("OPR_FLG:", options=avail_opr, default=[v for v in get_filter_selection("opr", default_opr) if v in avail_opr], key="flt_opr")

    # Tariff Category
    avail_tariff = sorted([str(v) for v in df_admin["TARIFF_CATEGORY"].dropna().unique() if str(v).strip() != ""])
    sel_tariff = st.multiselect("Tariff Category:", options=avail_tariff, default=[v for v in get_filter_selection("tariff", avail_tariff) if v in avail_tariff], key="flt_tariff")

    # Billed Flag
    avail_billed = ["All", "Billed", "Unbilled"]
    sel_billed = st.selectbox("Billed Flag:", options=avail_billed, index=0, key="flt_billed")

    # Paid Flag
    avail_paid = ["All", "Paid", "Unpaid"]
    sel_paid = st.selectbox("Paid Flag:", options=avail_paid, index=0, key="flt_paid")

    # OA Config Flag
    avail_oa = ["All", "YES", "NO"]
    default_oa_idx = 1 if is_oa_mode else 0
    sel_oa = st.selectbox("OA Config (MTR_CONFIG contains OA):", options=avail_oa, index=default_oa_idx, key="flt_oa")

    # Smart Meter Flag
    avail_sm = ["All", "YES", "NO"]
    sel_sm = st.selectbox("Smart Meter Flag (CMPRESYS contains MD):", options=avail_sm, index=0, key="flt_sm")

    # Load Range
    avail_load_rng = ["1-4 KW", "5-9 KW", "10-24 KW", "25-50 KW", "Above 50 KW", "Unknown/Invalid"]
    present_load_rng = [r for r in avail_load_rng if r in df_admin["LOAD_RANGE"].unique()]
    sel_load_rng = st.multiselect("Load Range:", options=present_load_rng, default=[v for v in get_filter_selection("load_rng", present_load_rng) if v in present_load_rng], key="flt_load_rng")

with st.sidebar.expander("📟 Metering & Billing Attributes", expanded=False):
    # Supply Type
    avail_supply = sorted([str(v) for v in df_admin["SUPPLY_TYPE"].dropna().unique() if str(v).strip() != ""])
    sel_supply = st.multiselect("SUPPLY_TYPE:", options=avail_supply, default=avail_supply, key="flt_supply")

    # BILL_CYC_CD
    avail_cyc = sorted([str(v) for v in df_admin["BILL_CYC_CD"].dropna().unique() if str(v).strip() != ""])
    sel_cyc = st.multiselect("BILL_CYC_CD:", options=avail_cyc, default=avail_cyc, key="flt_cyc")

    # MTR_TYPE_CD
    avail_mtr_type = sorted([str(v) for v in df_admin["MTR_TYPE_CD"].dropna().unique() if str(v).strip() != ""])
    sel_mtr_type = st.multiselect("MTR_TYPE_CD:", options=avail_mtr_type, default=avail_mtr_type, key="flt_mtr_type")

    # MTR_MAKE
    avail_make = sorted([str(v) for v in df_admin["MTR_MAKE"].dropna().unique() if str(v).strip() != ""])
    sel_make = st.multiselect("MTR_MAKE:", options=avail_make, default=avail_make, key="flt_make")

    # BILL_BASIS
    avail_basis = sorted([str(v) for v in df_admin["BILL_BASIS"].dropna().unique() if str(v).strip() != ""])
    sel_basis = st.multiselect("BILL_BASIS:", options=avail_basis, default=avail_basis, key="flt_basis")

    # INF_BILL
    avail_inf = sorted([str(v) for v in df_admin["INF_BILL"].dropna().unique() if str(v).strip() != ""])
    sel_inf = st.multiselect("INF_BILL:", options=avail_inf, default=avail_inf, key="flt_inf")

# ----------------- APPLY ACTIVE FILTERS TO DATASET -----------------
filtered_df = df_raw.copy()

# Administrative Hierarchy Filter
if sel_discoms:
    filtered_df = filtered_df[filtered_df["DISCOM"].isin(sel_discoms)]
if sel_zones:
    filtered_df = filtered_df[filtered_df["ZONE_NAME"].isin(sel_zones)]
if sel_circles:
    filtered_df = filtered_df[filtered_df["CIRCLE_NAME"].isin(sel_circles)]
if sel_divs:
    filtered_df = filtered_df[filtered_df["DIVISION_NAME"].isin(sel_divs)]

# Operational Filters
if sel_opr:
    filtered_df = filtered_df[filtered_df["OPR_FLG"].astype(str).isin(sel_opr)]
if sel_tariff:
    filtered_df = filtered_df[filtered_df["TARIFF_CATEGORY"].astype(str).isin(sel_tariff)]
if sel_billed != "All":
    filtered_df = filtered_df[filtered_df["BILLED_FLAG"] == sel_billed]
if sel_paid != "All":
    filtered_df = filtered_df[filtered_df["PAID_FLAG"] == sel_paid]
if sel_oa != "All":
    filtered_df = filtered_df[filtered_df["OA_CONFIG"] == sel_oa]
if sel_sm != "All":
    filtered_df = filtered_df[filtered_df["SMART_METER_FLAG"] == sel_sm]
if sel_load_rng:
    filtered_df = filtered_df[filtered_df["LOAD_RANGE"].isin(sel_load_rng)]

# Extended Metering Filters
if sel_supply:
    filtered_df = filtered_df[filtered_df["SUPPLY_TYPE"].astype(str).isin(sel_supply)]
if sel_cyc:
    filtered_df = filtered_df[filtered_df["BILL_CYC_CD"].astype(str).isin(sel_cyc)]
if sel_mtr_type:
    filtered_df = filtered_df[filtered_df["MTR_TYPE_CD"].astype(str).isin(sel_mtr_type)]
if sel_make:
    filtered_df = filtered_df[filtered_df["MTR_MAKE"].astype(str).isin(sel_make)]
if sel_basis:
    filtered_df = filtered_df[filtered_df["BILL_BASIS"].astype(str).isin(sel_basis)]
if sel_inf:
    filtered_df = filtered_df[filtered_df["INF_BILL"].astype(str).isin(sel_inf)]

# Calculate KPI Indicators
kpi = calculate_kpi_summary(filtered_df)

# ----------------- SECTION 4: CLICKABLE KPI SUMMARY INDICATORS -----------------
st.markdown("### 📊 Monthly Billing & Revenue Performance Indicators")
st.markdown(f"**Reference Period:** `{MONTH_NAMES.get(effective_month, '')} {effective_year}` &nbsp;|&nbsp; **Scope:** `{analysis_mode}` &nbsp;|&nbsp; **Filtered Consumers:** `{format_in(kpi['total_consumers'])}`")

kpi_c1, kpi_c2, kpi_c3, kpi_c4 = st.columns(4)
kpi_c5, kpi_c6, kpi_c7, kpi_c8 = st.columns(4)

with kpi_c1:
    st.markdown(f"""
    <div class="kpi-card highlight-card">
        <div class="kpi-title">Total Consumers (Distinct ACCT_ID)</div>
        <div class="kpi-value">{format_in(kpi['total_consumers'])}</div>
        <div class="kpi-sub">Total active in filtered scope</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔍 View All in Scope", key="btn_drill_all", use_container_width=True):
        st.session_state.drilldown_filter = "All"
        st.rerun()

with kpi_c2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Billed Consumers</div>
        <div class="kpi-value">{format_in(kpi['billed_consumers'])}</div>
        <div class="kpi-sub">Billing Ratio: <b>{kpi['billing_ratio']:.1f}%</b></div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔍 Filter Billed Only", key="btn_drill_billed", use_container_width=True):
        st.session_state.drilldown_filter = "Billed"
        st.rerun()

with kpi_c3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Unbilled Consumers</div>
        <div class="kpi-value">{format_in(kpi['unbilled_consumers'])}</div>
        <div class="kpi-sub">Unbilled Ratio: <b>{100.0 - kpi['billing_ratio']:.1f}%</b></div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔍 Filter Unbilled Only", key="btn_drill_unbilled", use_container_width=True):
        st.session_state.drilldown_filter = "Unbilled"
        st.rerun()

with kpi_c4:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Sold Units (Consumption)</div>
        <div class="kpi-value">{format_in(kpi['sold_units_mu'], 2)} <span style="font-size: 1rem; color: #94A3B8;">MU</span></div>
        <div class="kpi-sub">{format_in(kpi['sold_units_kwh'], 0)} kWh</div>
    </div>
    """, unsafe_allow_html=True)
    st.button("⚡ Energy Sold", key="btn_drill_sold", disabled=True, use_container_width=True)

with kpi_c5:
    st.markdown(f"""
    <div class="kpi-card highlight-card">
        <div class="kpi-title">Current Assessment (CA)</div>
        <div class="kpi-value">{format_currency_dual(kpi['current_assessment'])}</div>
        <div class="kpi-sub">Authentic Current Assessment</div>
    </div>
    """, unsafe_allow_html=True)
    st.button("💰 Current Assessment", key="btn_drill_ca", disabled=True, use_container_width=True)

with kpi_c6:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">CA Excl. Subsidy (CA_EX_SUBS)</div>
        <div class="kpi-value">{format_currency_dual(kpi['ca_ex_subs'])}</div>
        <div class="kpi-sub">Excluding government subsidy</div>
    </div>
    """, unsafe_allow_html=True)
    st.button("🏷️ CA Excl. Subsidy", key="btn_drill_ca_sub", disabled=True, use_container_width=True)

with kpi_c7:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Total Realization (Payments)</div>
        <div class="kpi-value">{format_currency_dual(kpi['payment_amount'])}</div>
        <div class="kpi-sub">Paying Consumers: <b>{format_in(kpi['payment_count'])}</b></div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔍 Filter Paying Only", key="btn_drill_paid", use_container_width=True):
        st.session_state.drilldown_filter = "Paid"
        st.rerun()

with kpi_c8:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Total Outstanding Dues</div>
        <div class="kpi-value">{format_currency_dual(kpi['total_outstanding'])}</div>
        <div class="kpi-sub">Cumulative arrears & dues</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔍 Filter Outstanding Only", key="btn_drill_out", use_container_width=True):
        st.session_state.drilldown_filter = "Outstanding"
        st.rerun()

# ----------------- SECTION 6: DASHBOARD VISUALIZATIONS -----------------
st.markdown("---")
st.markdown("### 📈 Analytical Breakdowns & Visualizations")

tab_bill, tab_energy, tab_rev, tab_tech = st.tabs([
    "📑 Billing Analysis",
    "⚡ Energy Consumption",
    "💰 Revenue & Realization",
    "📟 Smart Meter & OA Analysis"
])

# Tab 1: Billing Analysis
with tab_bill:
    col_b1, col_b2 = st.columns([1, 1.5])
    
    with col_b1:
        # Donut Chart: Billed vs Unbilled
        b_labels = ["Billed Consumers", "Unbilled Consumers"]
        b_values = [kpi["billed_consumers"], kpi["unbilled_consumers"]]
        fig_donut = go.Figure(data=[go.Pie(
            labels=b_labels,
            values=b_values,
            hole=0.55,
            marker=dict(colors=["#10B981", "#EF4444"]),
            textinfo="label+percent",
            hoverinfo="label+value+percent"
        )])
        fig_donut.update_layout(
            title_text="<b>Billed vs Unbilled Proportions</b>",
            margin=dict(l=20, r=20, t=40, b=20),
            height=320
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    with col_b2:
        # Billed Consumers by DISCOM
        if not filtered_df.empty:
            grp_disc = filtered_df.groupby(["DISCOM", "BILLED_FLAG"])["ACCT_ID"].nunique().reset_index()
            fig_disc_bar = px.bar(
                grp_disc,
                x="DISCOM",
                y="ACCT_ID",
                color="BILLED_FLAG",
                barmode="group",
                title="<b>Billed vs Unbilled Consumers by DISCOM</b>",
                labels={"ACCT_ID": "Consumer Count", "BILLED_FLAG": "Status"},
                color_discrete_map={"Billed": "#10B981", "Unbilled": "#EF4444"},
                text_auto=True
            )
            fig_disc_bar.update_layout(height=320, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_disc_bar, use_container_width=True)

    # Billing by Tariff Category
    if not filtered_df.empty:
        grp_tariff = filtered_df.groupby(["TARIFF_CATEGORY", "BILLED_FLAG"])["ACCT_ID"].nunique().reset_index()
        fig_tariff = px.bar(
            grp_tariff,
            x="TARIFF_CATEGORY",
            y="ACCT_ID",
            color="BILLED_FLAG",
            title="<b>Billed Consumers by Tariff Category</b>",
            labels={"ACCT_ID": "Consumer Count", "TARIFF_CATEGORY": "Tariff"},
            color_discrete_map={"Billed": "#10B981", "Unbilled": "#EF4444"},
            text_auto=True
        )
        fig_tariff.update_layout(height=320, margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_tariff, use_container_width=True)

# Tab 2: Energy Consumption Analysis
with tab_energy:
    col_e1, col_e2 = st.columns(2)
    
    with col_e1:
        if not filtered_df.empty:
            grp_energy_disc = filtered_df.groupby("DISCOM")["CONSUMPTION_CURR_MNTH"].sum().reset_index()
            grp_energy_disc["MU"] = grp_energy_disc["CONSUMPTION_CURR_MNTH"] / 1e6
            fig_e_disc = px.bar(
                grp_energy_disc,
                x="DISCOM",
                y="MU",
                color="DISCOM",
                title="<b>Sold Units (Million Units MU) by DISCOM</b>",
                labels={"MU": "Energy Sold (MU)"},
                text_auto=".2f"
            )
            fig_e_disc.update_layout(height=340, margin=dict(l=20, r=20, t=40, b=20), showlegend=False)
            st.plotly_chart(fig_e_disc, use_container_width=True)

    with col_e2:
        if not filtered_df.empty:
            grp_energy_rng = filtered_df.groupby("LOAD_RANGE")["CONSUMPTION_CURR_MNTH"].sum().reset_index()
            grp_energy_rng["MU"] = grp_energy_rng["CONSUMPTION_CURR_MNTH"] / 1e6
            fig_e_rng = px.bar(
                grp_energy_rng,
                x="LOAD_RANGE",
                y="MU",
                title="<b>Sold Units (MU) by Load Range</b>",
                labels={"MU": "Energy Sold (MU)", "LOAD_RANGE": "Contracted Load Range"},
                color="LOAD_RANGE",
                text_auto=".2f"
            )
            fig_e_rng.update_layout(height=340, margin=dict(l=20, r=20, t=40, b=20), showlegend=False)
            st.plotly_chart(fig_e_rng, use_container_width=True)

# Tab 3: Revenue & Realization Analysis
with tab_rev:
    col_r1, col_r2 = st.columns(2)
    
    with col_r1:
        if not filtered_df.empty:
            grp_rev_disc = filtered_df.groupby("DISCOM").agg({
                "CURRENT_ASSESSMENT": lambda x: x.sum() / 1e7,
                "TOTAL_PAY_AMT": lambda x: x.sum() / 1e7
            }).reset_index()
            grp_rev_disc = grp_rev_disc.rename(columns={
                "CURRENT_ASSESSMENT": "Assessment (₹ Cr)",
                "TOTAL_PAY_AMT": "Realization (₹ Cr)"
            })
            grp_melt = grp_rev_disc.melt(id_vars="DISCOM", value_vars=["Assessment (₹ Cr)", "Realization (₹ Cr)"], var_name="Metric", value_name="Amount_Cr")
            
            fig_rev = px.bar(
                grp_melt,
                x="DISCOM",
                y="Amount_Cr",
                color="Metric",
                barmode="group",
                title="<b>Current Assessment vs Payment Realization (₹ Crore)</b>",
                labels={"Amount_Cr": "₹ Crore"},
                color_discrete_map={"Assessment (₹ Cr)": "#3B82F6", "Realization (₹ Cr)": "#10B981"},
                text_auto=".2f"
            )
            fig_rev.update_layout(height=340, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_rev, use_container_width=True)

    with col_r2:
        if not filtered_df.empty:
            grp_out_disc = filtered_df.groupby("DISCOM")["TOTAL_OUTSTANDING"].sum().reset_index()
            grp_out_disc["Outstanding_Cr"] = grp_out_disc["TOTAL_OUTSTANDING"] / 1e7
            fig_out = px.bar(
                grp_out_disc,
                x="DISCOM",
                y="Outstanding_Cr",
                color="DISCOM",
                title="<b>Total Outstanding Balance by DISCOM (₹ Crore)</b>",
                labels={"Outstanding_Cr": "Outstanding (₹ Cr)"},
                text_auto=".2f"
            )
            fig_out.update_layout(height=340, margin=dict(l=20, r=20, t=40, b=20), showlegend=False)
            st.plotly_chart(fig_out, use_container_width=True)

# Tab 4: Smart Meter & OA Analysis
with tab_tech:
    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        if not filtered_df.empty:
            grp_sm = filtered_df["SMART_METER_FLAG"].value_counts().reset_index()
            grp_sm.columns = ["Smart_Meter", "Consumers"]
            fig_sm = px.pie(
                grp_sm,
                names="Smart_Meter",
                values="Consumers",
                hole=0.5,
                title="<b>Smart Meter vs Non-Smart Meter Consumers</b>",
                color="Smart_Meter",
                color_discrete_map={"YES": "#06B6D4", "NO": "#64748B"}
            )
            fig_sm.update_layout(height=320, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_sm, use_container_width=True)

    with col_t2:
        if not filtered_df.empty:
            grp_oa = filtered_df["OA_CONFIG"].value_counts().reset_index()
            grp_oa.columns = ["OA_Config", "Consumers"]
            fig_oa = px.pie(
                grp_oa,
                names="OA_Config",
                values="Consumers",
                hole=0.5,
                title="<b>Open Access Configuration Distribution</b>",
                color="OA_Config",
                color_discrete_map={"YES": "#8B5CF6", "NO": "#94A3B8"}
            )
            fig_oa.update_layout(height=320, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_oa, use_container_width=True)

# ----------------- SECTION 7: CONSUMER-LEVEL DATA TABLE & EXCEL EXPORT -----------------
st.markdown("---")

# Active Drilldown Application
drill_state = st.session_state.drilldown_filter
display_df = filtered_df.copy()

if drill_state == "Billed":
    display_df = display_df[display_df["BILLED_FLAG"] == "Billed"]
elif drill_state == "Unbilled":
    display_df = display_df[display_df["BILLED_FLAG"] == "Unbilled"]
elif drill_state == "Paid":
    display_df = display_df[display_df["PAID_FLAG"] == "Paid"]
elif drill_state == "Outstanding":
    display_df = display_df[display_df["TOTAL_OUTSTANDING"] > 0]

st.markdown(f"### 📋 Consumer-Level Dossier Table (Strict 1:1 Parity)")
st.markdown(f"**Active Table Isolation:** `{drill_state}` &nbsp;|&nbsp; **Qualifying Consumer Records:** `{format_in(len(display_df))}` of `{format_in(len(filtered_df))}` filtered")

# Table Search & Column Customization Bar
t_col1, t_col2, t_col3 = st.columns([3, 2, 2])

with t_col1:
    search_query = st.text_input(
        "🔎 Search Records (Account ID, Name, Division, Meter No):",
        placeholder="Type search terms...",
        help="Real-time multi-column search across records"
    )

with t_col2:
    view_preset = st.selectbox(
        "👁️ Column View Preset:",
        ["Standard Executive View", "Billing & Financial", "Technical & Metering", "All Available Columns"],
        index=0
    )

with t_col3:
    records_per_page = st.selectbox(
        "📄 Records Per Page:",
        options=[25, 50, 100, 250, 500, 1000],
        index=1
    )

# Apply Text Search
if search_query.strip():
    q = search_query.strip().upper()
    search_cols = [c for c in ["ACCT_ID", "NAME", "DIV_CODE", "DIVISION_NAME", "SERIAL_NBR", "METER_BADGE_NO", "SCNO", "KNO"] if c in display_df.columns]
    mask = pd.Series([False] * len(display_df), index=display_df.index)
    for c in search_cols:
        mask = mask | display_df[c].astype(str).str.upper().str.contains(q, na=False)
    display_df = display_df[mask]

# Define Preset Columns
if view_preset == "Standard Executive View":
    preset_cols = [
        "ACCT_ID", "NAME", "DISCOM", "ZONE_NAME", "CIRCLE_NAME", "DIVISION_NAME",
        "SUPPLY_TYPE", "TARIFF_CATEGORY", "LOAD", "LOAD_UNIT", "LOAD_KW", "LOAD_RANGE",
        "OPR_FLG", "BILLED_FLAG", "PAID_FLAG", "LAST_BILL_DATE",
        "CONSUMPTION_CURR_MNTH", "CURRENT_ASSESSMENT", "CA_EX_SUBS",
        "TOTAL_PAY_AMT", "TOTAL_OUTSTANDING", "MTR_CONFIG_TY_CD", "OA_CONFIG", "SMART_METER_FLAG"
    ]
elif view_preset == "Billing & Financial":
    preset_cols = [
        "ACCT_ID", "NAME", "DISCOM", "DIVISION_NAME", "SUPPLY_TYPE", "TARIFF_CATEGORY",
        "LAST_BILL_DATE", "BILLED_FLAG", "CONSUMPTION_CURR_MNTH", "BILLED_AMOUNT",
        "CURRENT_ASSESSMENT", "CA_EX_SUBS", "ENERGY_AMT", "NON_ENERGY_AMT",
        "PAY_AMT", "TOTAL_PAY_AMT", "PAID_FLAG", "PAY_DATE", "TOTAL_OUTSTANDING", "ARREAR"
    ]
elif view_preset == "Technical & Metering":
    preset_cols = [
        "ACCT_ID", "NAME", "DISCOM", "DIVISION_NAME", "LOAD", "LOAD_UNIT", "LOAD_KW", "LOAD_RANGE",
        "SERIAL_NBR", "MTR_MAKE", "MTR_TYPE_CD", "MTR_CONFIG_TY_CD", "OA_CONFIG",
        "CMPRESYS_CODE", "SMART_METER_FLAG", "OPENING_READING_KWH", "CLOSING_READING_KWH",
        "CONSUMPTION_CURR_MNTH", "POWER_FACTOR", "METER_READ_REMARK"
    ]
else:
    preset_cols = list(display_df.columns)

# Intersect with present columns
active_cols = [c for c in preset_cols if c in display_df.columns]

# Pagination Controls
total_rows = len(display_df)
total_pages = max(1, math.ceil(total_rows / records_per_page))

p_col1, p_col2, p_col3 = st.columns([2, 3, 2])
with p_col2:
    page_num = st.number_input(f"Page (1 to {total_pages}):", min_value=1, max_value=total_pages, value=1, step=1)

start_idx = (page_num - 1) * records_per_page
end_idx = min(start_idx + records_per_page, total_rows)

# Paginated DataFrame
page_df = display_df.iloc[start_idx:end_idx][active_cols]

# Render Data Table
st.dataframe(
    page_df,
    use_container_width=True,
    hide_index=True
)

st.caption(f"Showing records {format_in(start_idx + 1 if total_rows > 0 else 0)} to {format_in(end_idx)} of {format_in(total_rows)} qualifying consumer records.")

# ----------------- DOWNLOAD BUTTONS (1:1 PARITY) -----------------
st.markdown("#### 📥 Authoritative Data Exports (1:1 Count Parity)")

dl_col1, dl_col2 = st.columns(2)

with dl_col1:
    excel_data = generate_excel_export(display_df[active_cols], title="Open_Access_Consumers")
    st.download_button(
        label=f"📥 Download Filtered Records as Excel (.xlsx) [{format_in(len(display_df))} Rows]",
        data=excel_data,
        file_name=f"Open_Access_Consumers_{effective_month}_{effective_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with dl_col2:
    csv_bytes = display_df[active_cols].to_csv(index=False).encode("utf-8")
    st.download_button(
        label=f"📥 Download Filtered Records as CSV (.csv) [{format_in(len(display_df))} Rows]",
        data=csv_bytes,
        file_name=f"Open_Access_Consumers_{effective_month}_{effective_year}.csv",
        mime="text/csv",
        use_container_width=True
    )
