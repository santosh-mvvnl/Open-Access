r"""
Open Access Consumers & Billing Revenue Analysis Dashboard
A high-performance, executive dashboard conforming strictly to Universal Agent Rules (global.md):
- Rule #1: Mandatory Confirmation Before Codebase Changes
- Rule #2: Automatic Git Add, Commit, and Push After Approved Changes
- Rule #3: Mandatory 6 Structured Response Sections
- Rule #4: Pure Authentic Portal Data (Strict Zero Synthetic Data Policy)
- Rule #5: Automatic Synchronization of Global Rules (.gemini/rules/global.md)
- Rule #6: Summary Hyperlinks and Strict 1:1 Data Parity
- Rule #7: Contextual Dropdown Filters with Checkboxes ABOVE Summaries (Exact Excel AutoFilter Semantics)
- Rule #8: Indian Number Formatting (1,00,000, 10,00,000, 1,00,00,000)
- Rule #9: Avoid Hardcoding & Generalized Configurable Architecture
- Rule #10: Validation, Testing, and Regression Prevention
- Rule #11: Mandatory Working Procedure & Scope/Priority
- Rule #12: High-Performance Omnidirectional Cascading Filters on Large Datasets
"""

import os
import re
import glob
import time
import math
from datetime import datetime
from typing import Dict, List, Any, Optional, Set

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
    initial_sidebar_state="collapsed"
)

# ----------------- REFINED EXECUTIVE STYLING -----------------
st.markdown("""
<style>
    /* Global Layout Constraints */
    .main .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        max-width: 98%;
    }
    
    /* Top Header Bar */
    .header-banner {
        background: linear-gradient(90deg, #0F172A 0%, #1E293B 100%);
        padding: 14px 22px;
        border-radius: 10px;
        border: 1px solid #334155;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .header-title {
        font-size: 1.4rem;
        font-weight: 700;
        color: #F8FAFC;
        margin: 0;
        letter-spacing: -0.01em;
    }
    .header-sub {
        font-size: 0.8rem;
        color: #94A3B8;
        margin-top: 2px;
        margin-bottom: 0;
    }
    
    /* Filter Bar Category Headers (Rule #7) */
    .filter-band-header {
        font-size: 0.8rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #38BDF8;
        margin-top: 6px;
        margin-bottom: 4px;
    }

    /* Executive KPI Cards (Rule #6 & #8) */
    .kpi-item {
        background: #0F172A;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 12px 14px;
        transition: all 0.2s ease;
        position: relative;
    }
    .kpi-item:hover {
        border-color: #38BDF8;
        box-shadow: 0 4px 12px rgba(56, 189, 248, 0.12);
    }
    .kpi-item.active-drill {
        border-color: #38BDF8;
        background: #132238;
        box-shadow: 0 0 0 2px #38BDF8;
    }
    .kpi-header-label {
        font-size: 0.74rem;
        font-weight: 600;
        text-transform: uppercase;
        color: #94A3B8;
        letter-spacing: 0.04em;
        margin-bottom: 3px;
    }
    .kpi-primary-val {
        font-size: 1.45rem;
        font-weight: 700;
        color: #38BDF8;
        line-height: 1.15;
    }
    .kpi-secondary-info {
        font-size: 0.74rem;
        color: #CBD5E1;
        margin-top: 3px;
    }
    
    /* Popover button compact styling */
    button[data-testid="baseButton-secondary"] {
        border-radius: 6px;
        font-size: 0.78rem;
        padding: 2px 8px;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- SESSION STATE MANAGEMENT -----------------
if "folder_path" not in st.session_state:
    st.session_state.folder_path = r"D:\Master Data\102026"
if "drilldown_filter" not in st.session_state:
    st.session_state.drilldown_filter = "All"
if "oa_df" not in st.session_state:
    st.session_state.oa_df = None
if "last_loaded_folder" not in st.session_state:
    st.session_state.last_loaded_folder = ""
if "filter_states" not in st.session_state:
    st.session_state.filter_states = {}
if "search_snapshots" not in st.session_state:
    st.session_state.search_snapshots = {}

# ----------------- NATIVE FOLDER BROWSER -----------------
def open_native_folder_browser():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        selected = filedialog.askdirectory(
            initialdir=r"D:\Master Data",
            title="Select Master Data Month Folder"
        )
        root.destroy()
        if selected:
            st.session_state.folder_path = os.path.normpath(selected)
            st.rerun()
    except Exception as e:
        st.warning(f"Could not open native folder browser: {e}")

# ----------------- 1. HEADER & TOP CONTROL BAR -----------------
st.markdown("""
<div class="header-banner">
    <div>
        <div class="header-title">⚡ Open Access Consumers & Billing Revenue Analysis Dashboard</div>
        <div class="header-sub">Pure Authentic Portal Data • UPPCL Master Billing Register (DVVNL • KESCO • MVVNL • PuVVNL • PVVNL)</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Master Data Folder & Period Bar
p_col1, p_col2, p_col3, p_col4, p_col5 = st.columns([3.5, 1.2, 2.2, 1.5, 1.5])

with p_col1:
    folder_input = st.text_input(
        "Master Data Folder Path:",
        value=st.session_state.folder_path,
        placeholder=r"e.g. D:\Master Data\102026",
        label_visibility="collapsed"
    )
    if folder_input != st.session_state.folder_path:
        st.session_state.folder_path = folder_input

with p_col2:
    if st.button("📂 Browse", use_container_width=True, help="Select folder using Windows Folder Browser"):
        open_native_folder_browser()

# Automatic Period Extraction (Section 2.2 & 2.3)
period_info = parse_folder_period(st.session_state.folder_path)

with p_col3:
    st.markdown(f"""
    <div style="background: #0F172A; border: 1px solid #334155; border-radius: 8px; padding: 6px 12px; height: 38px; display: flex; align-items: center; justify-content: space-between;">
        <span style="font-size: 0.82rem; color: #94A3B8;">Master Period:</span>
        <span style="font-size: 0.90rem; font-weight: 700; color: #38BDF8;">{period_info['display']}</span>
    </div>
    """, unsafe_allow_html=True)

with p_col4:
    btn_load = st.button("🚀 Load Data", type="primary", use_container_width=True)

with p_col5:
    btn_refresh = st.button("🔄 Refresh Cache", use_container_width=True, help="Re-scan archives and rebuild cache")

# Manual Period Override Popover
with st.expander("🛠️ Master Period Manual Override (Optional)", expanded=False):
    ov_c1, ov_c2, ov_c3 = st.columns([2, 2, 4])
    with ov_c1:
        override_m = st.selectbox(
            "Master Month (Override):",
            options=[f"{i:02d}" for i in range(1, 13)],
            index=int(period_info['month']) - 1,
            format_func=lambda m: f"{m} - {MONTH_NAMES.get(m, '')}"
        )
    with ov_c2:
        override_y = st.text_input("Master Year (Override):", value=period_info['year'])
    with ov_c3:
        st.caption("The Master Month & Year serve as the reference period for all BILLED_FLAG calculations.")

effective_month = override_m
effective_year = override_y.strip()

# Data Ingestion
should_load = btn_load or btn_refresh or (st.session_state.oa_df is None) or (st.session_state.last_loaded_folder != st.session_state.folder_path)

if should_load:
    if not os.path.exists(st.session_state.folder_path):
        st.error(f"Directory not found: `{st.session_state.folder_path}`. Please verify the folder path.")
    else:
        with st.spinner(f"Ingesting authentic master data for {MONTH_NAMES.get(effective_month, '')} {effective_year}..."):
            t0 = time.time()
            pbar = st.progress(0.0)
            st_text = st.empty()

            def progress_cb(frac, msg):
                pbar.progress(frac)
                st_text.text(msg)

            df_raw = load_or_extract_oa_dataset(
                base_dir=st.session_state.folder_path,
                master_month=effective_month,
                master_year=effective_year,
                progress_callback=progress_cb,
                force_refresh=btn_refresh
            )
            pbar.empty()
            st_text.empty()
            elapsed = time.time() - t0

            if df_raw.empty:
                st.warning("No qualifying Open Access records found in the selected folder.")
            else:
                st.session_state.oa_df = df_raw
                st.session_state.last_loaded_folder = st.session_state.folder_path
                st.toast(f"Loaded {len(df_raw)} Open Access records in {elapsed:.2f}s!", icon="⚡")

df_full = st.session_state.oa_df

if df_full is None or df_full.empty:
    st.info("👆 Click **'🚀 Load Data'** above to ingest and analyze authentic monthly master records.")
    st.stop()

# ----------------- 2. EXCEL AUTOFILTER ENGINE (RULE #7: EXACT SPECIFICATION) -----------------
def render_excel_autofilter_popover(
    label: str,
    available_options: List[Any],
    filter_key: str,
    default_selected: Optional[List[Any]] = None
) -> List[Any]:
    """
    Renders an AutoFilter popover conforming strictly to Rule #7:
    - Default State — All Checked: Default to all options checked (value=True) initially (or default_selected if specified).
    - Top Search Box (Excel AutoFilter Style): Positioned permanently at the top.
    - Manual Search Selection with Explicit Add-to-Filter Branching:
      * When search is entered, newly matching options & (Select All Search Results) are unchecked by default.
      * If Add current selection to filter is unchecked: Isolates strictly to newly selected search items.
      * If Add current selection to filter is checked: Seamlessly merges committed base with selected search items.
    - (Select All Search Results) Checkbox: Dynamically displayed when text is in search box.
    - Add current selection to filter Checkbox: Controls isolated vs incremental search merge.
    - Persistent Committed Base Snapshot: Base selections preserved before a search query begins.
    - (Select All) Checkbox When Search is Empty: Master toggle for scope.
    - Pure Master Checkbox Bulk Selection: Driven exclusively via master checkboxes without internal push buttons.
    - Independent Filter State Isolation: Never resets or invalidates other filters.
    - Instant Zero-Latency Synchronization: Syncs widget state before badge and downstream calculation on the current turn.
    """
    state_store = st.session_state.filter_states
    snapshots = st.session_state.search_snapshots

    # 1. Initialize persistent committed base
    if filter_key not in state_store:
        if default_selected is not None:
            def_set = {str(x) for x in default_selected}
            state_store[filter_key] = {str(opt): (str(opt) in def_set) for opt in available_options}
        else:
            state_store[filter_key] = {str(opt): True for opt in available_options}
        
    choice_dict = state_store[filter_key]
    for opt in available_options:
        s_opt = str(opt)
        if s_opt not in choice_dict:
            if default_selected is not None:
                choice_dict[s_opt] = (s_opt in {str(x) for x in default_selected})
            else:
                choice_dict[s_opt] = True

    # 2. IMMEDIATE WIDGET-STATE SYNCHRONIZATION (BEFORE BADGE & RENDERING)
    # Check if (Select All) master was toggled
    master_key = f"{filter_key}_chk_master_all"
    prev_master_key = f"{filter_key}_prev_master"

    if master_key in st.session_state:
        curr_master = st.session_state[master_key]
        prev_master = st.session_state.get(prev_master_key, curr_master)
        if curr_master != prev_master:
            # User clicked master checkbox!
            for opt in available_options:
                s_opt = str(opt)
                choice_dict[s_opt] = curr_master
                st.session_state[f"{filter_key}_item_{s_opt}"] = curr_master
            st.session_state[prev_master_key] = curr_master
        else:
            # Sync individual item checkboxes into choice_dict
            for opt in available_options:
                s_opt = str(opt)
                item_key = f"{filter_key}_item_{s_opt}"
                if item_key in st.session_state:
                    choice_dict[s_opt] = bool(st.session_state[item_key])
                else:
                    st.session_state[item_key] = choice_dict.get(s_opt, True)
    else:
        # Initial run or after reset: pre-seed item widget keys
        for opt in available_options:
            s_opt = str(opt)
            item_key = f"{filter_key}_item_{s_opt}"
            if item_key in st.session_state:
                choice_dict[s_opt] = bool(st.session_state[item_key])
            else:
                st.session_state[item_key] = choice_dict.get(s_opt, True)

    # Search Mode widget sync
    search_query = st.session_state.get(f"{filter_key}_search_txt", "").strip()
    if search_query:
        q_lower = search_query.lower()
        matching_options = [opt for opt in available_options if q_lower in str(opt).lower()]
        search_dict_key = f"{filter_key}_search_items_{search_query}"
        if search_dict_key in st.session_state:
            search_choices = st.session_state[search_dict_key]
            search_master_key = f"{filter_key}_chk_search_master"
            prev_sm_key = f"{filter_key}_prev_search_master"
            if search_master_key in st.session_state:
                curr_sm = st.session_state[search_master_key]
                prev_sm = st.session_state.get(prev_sm_key, curr_sm)
                if curr_sm != prev_sm:
                    for opt in matching_options:
                        s_opt = str(opt)
                        search_choices[s_opt] = curr_sm
                        st.session_state[f"{filter_key}_sitem_{s_opt}"] = curr_sm
                    st.session_state[prev_sm_key] = curr_sm
                else:
                    for opt in matching_options:
                        s_opt = str(opt)
                        sitem_key = f"{filter_key}_sitem_{s_opt}"
                        if sitem_key in st.session_state:
                            search_choices[s_opt] = bool(st.session_state[sitem_key])
            add_to_filter = st.session_state.get(f"{filter_key}_add_merge_chk", False)
            base_snap = snapshots.get(filter_key, dict(choice_dict))
            if add_to_filter:
                for opt in available_options:
                    s_opt = str(opt)
                    if s_opt in search_choices:
                        choice_dict[s_opt] = base_snap.get(s_opt, False) or search_choices[s_opt]
                    else:
                        choice_dict[s_opt] = base_snap.get(s_opt, False)
            else:
                for opt in available_options:
                    s_opt = str(opt)
                    choice_dict[s_opt] = search_choices.get(s_opt, False)
            for opt in available_options:
                s_opt = str(opt)
                st.session_state[f"{filter_key}_item_{s_opt}"] = choice_dict[s_opt]

    # Current committed selections
    active_checked = [opt for opt in available_options if choice_dict.get(str(opt), True)]
    total_opts = len(available_options)
    n_sel = len(active_checked)

    # Dynamic status badge
    if n_sel == total_opts and total_opts > 0:
        badge = "(All)"
    elif n_sel == 0:
        badge = "(None)"
    else:
        badge = f"({n_sel}/{total_opts})"

    popover_label = f"{label}: {badge}"
    st.session_state[prev_master_key] = (n_sel == total_opts)

    with st.popover(popover_label, use_container_width=True):
        st.markdown(f"**Filter by {label}**")
        
        if total_opts > 0:
            # Top Search Box (Excel AutoFilter Style)
            search_query = st.text_input(
                "Search",
                key=f"{filter_key}_search_txt",
                placeholder="Search values...",
                label_visibility="collapsed"
            ).strip()

            # Persistent Committed Base Snapshot Management
            prev_search = st.session_state.get(f"{filter_key}_prev_search", "")
            if not prev_search and search_query:
                # Capture base snapshot right as search begins
                snapshots[filter_key] = dict(choice_dict)
            elif prev_search and not search_query:
                # Cleared search: restore or commit base snapshot
                if filter_key in snapshots:
                    del snapshots[filter_key]

            st.session_state[f"{filter_key}_prev_search"] = search_query

            if search_query:
                # Active Search Mode
                q_lower = search_query.lower()
                matching_options = [
                    opt for opt in available_options
                    if q_lower in str(opt).lower()
                ]

                # Rule #7: Add current selection to filter Checkbox (Incremental Search Merge)
                add_to_filter = st.checkbox(
                    "Add current selection to filter",
                    key=f"{filter_key}_add_merge_chk",
                    value=False,
                    help="If unchecked (default): Isolates strictly to searched items. If checked: Merges previous base with searched items."
                )

                # Initialize search items as UNCHECKED by default on first entry
                search_dict_key = f"{filter_key}_search_items_{search_query}"
                if search_dict_key not in st.session_state:
                    st.session_state[search_dict_key] = {str(opt): False for opt in matching_options}
                search_choices = st.session_state[search_dict_key]

                # (Select All Search Results) Checkbox
                n_search_sel = sum(1 for opt in matching_options if search_choices.get(str(opt), False))
                is_all_search_checked = (n_search_sel == len(matching_options)) and len(matching_options) > 0

                chk_all_search = st.checkbox(
                    f"(Select All Search Results) [{len(matching_options)}]",
                    key=f"{filter_key}_chk_search_master",
                    value=is_all_search_checked
                )
                st.session_state[f"{filter_key}_prev_search_master"] = chk_all_search

                st.markdown("<hr style='margin: 4px 0 6px 0; border: none; border-top: 1px solid #334155;'/>", unsafe_allow_html=True)

                # Scrollable Search Results Checkbox List
                with st.container(height=170):
                    for opt in matching_options:
                        s_opt = str(opt)
                        curr_v = search_choices.get(s_opt, False)
                        sc = st.checkbox(s_opt, value=curr_v, key=f"{filter_key}_sitem_{s_opt}")
                        search_choices[s_opt] = sc

                # Merge Logic as mandated by Rule #7
                base_snap = snapshots.get(filter_key, dict(choice_dict))
                if add_to_filter:
                    # Merges previous base with checked search results
                    for opt in available_options:
                        s_opt = str(opt)
                        if s_opt in search_choices:
                            choice_dict[s_opt] = base_snap.get(s_opt, False) or search_choices[s_opt]
                        else:
                            choice_dict[s_opt] = base_snap.get(s_opt, False)
                else:
                    # Isolates strictly to searched items
                    for opt in available_options:
                        s_opt = str(opt)
                        choice_dict[s_opt] = search_choices.get(s_opt, False)

            else:
                # Empty Search Mode: Pure Master (Select All) Checkbox
                is_all_checked = (n_sel == total_opts)
                chk_select_all = st.checkbox(
                    f"(Select All) [{total_opts}]",
                    key=master_key,
                    value=is_all_checked
                )
                st.session_state[prev_master_key] = chk_select_all

                st.markdown("<hr style='margin: 4px 0 6px 0; border: none; border-top: 1px solid #334155;'/>", unsafe_allow_html=True)

                # Scrollable Full Options Checkbox List
                with st.container(height=170):
                    for opt in available_options:
                        s_opt = str(opt)
                        curr_val = choice_dict.get(s_opt, True)
                        checked = st.checkbox(s_opt, value=curr_val, key=f"{filter_key}_item_{s_opt}")
                        choice_dict[s_opt] = checked

        else:
            st.caption("No options available in current scope.")

    return [opt for opt in available_options if choice_dict.get(str(opt), True)]

# ----------------- RENDER ALL 22 FILTERS DIRECTLY ABOVE SUMMARIES -----------------
st.markdown("<div class='filter-band-header'>🏛️ Band 1: Administrative Hierarchy (Cascading Filters 1 to 4)</div>", unsafe_allow_html=True)

# Band 1: Administrative Hierarchy (Cascading 4 Filters)
b1_c1, b1_c2, b1_c3, b1_c4 = st.columns(4)

with b1_c1:
    avail_discoms = sorted([d for d in df_full["DISCOM"].dropna().unique() if str(d).strip()])
    sel_discoms = render_excel_autofilter_popover("1. DISCOM", avail_discoms, "af_discom")

df_step1 = df_full[df_full["DISCOM"].isin(sel_discoms)] if len(sel_discoms) < len(avail_discoms) else df_full

with b1_c2:
    avail_zones = sorted([z for z in df_step1["ZONE_NAME"].dropna().unique() if str(z).strip()])
    sel_zones = render_excel_autofilter_popover("2. Zone", avail_zones, "af_zone")

df_step2 = df_step1[df_step1["ZONE_NAME"].isin(sel_zones)] if len(sel_zones) < len(avail_zones) else df_step1

with b1_c3:
    avail_circles = sorted([c for c in df_step2["CIRCLE_NAME"].dropna().unique() if str(c).strip()])
    sel_circles = render_excel_autofilter_popover("3. Circle", avail_circles, "af_circle")

df_step3 = df_step2[df_step2["CIRCLE_NAME"].isin(sel_circles)] if len(sel_circles) < len(avail_circles) else df_step2

with b1_c4:
    avail_divs = sorted([d for d in df_step3["DIVISION_NAME"].dropna().unique() if str(d).strip()])
    sel_divs = render_excel_autofilter_popover("4. Division", avail_divs, "af_division")

df_admin = df_step3[df_step3["DIVISION_NAME"].isin(sel_divs)] if len(sel_divs) < len(avail_divs) else df_step3

# Band 2: Operational, Billing & Technical Flags (Filters 5 to 11)
st.markdown("<div class='filter-band-header'>⚡ Band 2: Operational, Billing & Technical Status (Filters 5 to 11)</div>", unsafe_allow_html=True)
b2_c1, b2_c2, b2_c3, b2_c4, b2_c5, b2_c6, b2_c7 = st.columns(7)

with b2_c1:
    avail_opr = sorted(list(set(["Y", "N"]).union(set(str(o) for o in df_admin["OPR_FLG"].dropna().unique() if str(o).strip()))))
    sel_opr = render_excel_autofilter_popover("5. OPR_FLAG", avail_opr, "af_opr", default_selected=["Y"])

with b2_c2:
    avail_tariff = sorted([str(t) for t in df_admin["TARIFF_CATEGORY"].dropna().unique() if str(t).strip()])
    sel_tariff = render_excel_autofilter_popover("6. Tariff Category", avail_tariff, "af_tariff")

with b2_c3:
    avail_billed = ["Billed", "Unbilled"]
    sel_billed = render_excel_autofilter_popover("7. Billed Flag", avail_billed, "af_billed")

with b2_c4:
    avail_paid = ["Paid", "Unpaid"]
    sel_paid = render_excel_autofilter_popover("8. Paid Flag", avail_paid, "af_paid")

with b2_c5:
    avail_oa = ["YES", "NO"]
    sel_oa = render_excel_autofilter_popover("9. OA Config", avail_oa, "af_oa", default_selected=["YES"])

with b2_c6:
    avail_sm = ["YES", "NO"]
    sel_sm = render_excel_autofilter_popover("10. Smart Meter Flag", avail_sm, "af_sm")

with b2_c7:
    avail_load_rng = ["1-4 KW", "5-9 KW", "10-24 KW", "25-50 KW", "Above 50 KW", "Unknown/Invalid"]
    present_load_rng = [r for r in avail_load_rng if r in df_admin["LOAD_RANGE"].unique()]
    sel_load_rng = render_excel_autofilter_popover("11. Load Range", present_load_rng, "af_load_range")

# Band 3: Metering, Reading & Connection Profile (Filters 12 to 22)
st.markdown("<div class='filter-band-header'>📟 Band 3: Metering, Reading & Connection Profile (Filters 12 to 22)</div>", unsafe_allow_html=True)
b3_r1_c1, b3_r1_c2, b3_r1_c3, b3_r1_c4, b3_r1_c5, b3_r1_c6 = st.columns(6)

with b3_r1_c1:
    avail_st = sorted([str(v) for v in df_admin["SUPPLY_TYPE"].dropna().unique() if str(v).strip()])
    sel_st = render_excel_autofilter_popover("12. SUPPLY_TYPE", avail_st, "af_supply_type")

with b3_r1_c2:
    avail_conn = sorted([str(v) for v in df_admin["CONNECTION_TYPE"].dropna().unique() if str(v).strip()])
    sel_conn = render_excel_autofilter_popover("13. TYPE_of_Connection", avail_conn, "af_conn_type")

with b3_r1_c3:
    avail_cyc = sorted([str(v) for v in df_admin["BILL_CYC_CD"].dropna().unique() if str(v).strip()])
    sel_cyc = render_excel_autofilter_popover("14. BILL_CYC_CD", avail_cyc, "af_bill_cyc")

with b3_r1_c4:
    avail_mtype = sorted([str(v) for v in df_admin["MTR_TYPE_CD"].dropna().unique() if str(v).strip()])
    sel_mtype = render_excel_autofilter_popover("15. MTR_TYPE_CD", avail_mtype, "af_mtr_type")

with b3_r1_c5:
    avail_mcfg = sorted([str(v) for v in df_admin["MTR_CONFIG_TY_CD"].dropna().unique() if str(v).strip()])
    sel_mcfg = render_excel_autofilter_popover("16. MTR_CONFIG_TY_CD", avail_mcfg, "af_mtr_cfg")

with b3_r1_c6:
    avail_make = sorted([str(v) for v in df_admin["MTR_MAKE"].dropna().unique() if str(v).strip()])
    sel_make = render_excel_autofilter_popover("17. MTR_MAKE", avail_make, "af_mtr_make")

# Sub-row 3B for remaining 5 filters + Universal Reset
b3_r2_c1, b3_r2_c2, b3_r2_c3, b3_r2_c4, b3_r2_c5, b3_r2_c6 = st.columns(6)

with b3_r2_c1:
    avail_mr_src = sorted([str(v) for v in df_admin["MR_SOURCE_CD"].dropna().unique() if str(v).strip()])
    sel_mr_src = render_excel_autofilter_popover("18. MR_SOURCE_CD", avail_mr_src, "af_mr_src")

with b3_r2_c2:
    avail_remark = sorted([str(v) for v in df_admin["METER_READ_REMARK"].dropna().unique() if str(v).strip()])
    sel_remark = render_excel_autofilter_popover("19. METER_READ_REMARK", avail_remark, "af_mr_remark")

with b3_r2_c3:
    avail_bb = sorted([str(v) for v in df_admin["BILL_BASIS"].dropna().unique() if str(v).strip()])
    sel_bb = render_excel_autofilter_popover("20. BILL_BASIS", avail_bb, "af_bill_basis")

with b3_r2_c4:
    avail_inf = sorted([str(v) for v in df_admin["INF_BILL"].dropna().unique() if str(v).strip()])
    sel_inf = render_excel_autofilter_popover("21. INF_BILL", avail_inf, "af_inf_bill")

with b3_r2_c5:
    ps_col = "CMPRESYS_CODE" if "CMPRESYS_CODE" in df_admin.columns else "PRESYS_CD"
    avail_presys = sorted([str(v) for v in df_admin[ps_col].dropna().unique() if str(v).strip()])
    sel_presys = render_excel_autofilter_popover("22. PRESYS_CD", avail_presys, "af_presys")

with b3_r2_c6:
    st.write("")
    if st.button("🔄 Reset All 22 Filters", use_container_width=True, help="Reset all 22 AutoFilters to initial default state"):
        st.session_state.filter_states = {}
        st.session_state.search_snapshots = {}
        st.session_state.drilldown_filter = "All"
        for k in list(st.session_state.keys()):
            if k.startswith("af_"):
                del st.session_state[k]
        st.rerun()

# ----------------- APPLY ACTIVE FILTERS ACROSS ALL 22 DIMENSIONS -----------------
active_df = df_admin.copy()

if len(sel_opr) < len(avail_opr):
    active_df = active_df[active_df["OPR_FLG"].astype(str).isin(sel_opr)]
if len(sel_tariff) < len(avail_tariff):
    active_df = active_df[active_df["TARIFF_CATEGORY"].astype(str).isin(sel_tariff)]
if len(sel_billed) < len(avail_billed):
    active_df = active_df[active_df["BILLED_FLAG"].isin(sel_billed)]
if len(sel_paid) < len(avail_paid):
    active_df = active_df[active_df["PAID_FLAG"].isin(sel_paid)]
if len(sel_oa) < len(avail_oa):
    active_df = active_df[active_df["OA_CONFIG"].isin(sel_oa)]
if len(sel_sm) < len(avail_sm):
    active_df = active_df[active_df["SMART_METER_FLAG"].isin(sel_sm)]
if len(sel_load_rng) < len(present_load_rng):
    active_df = active_df[active_df["LOAD_RANGE"].isin(sel_load_rng)]
if len(sel_st) < len(avail_st):
    active_df = active_df[active_df["SUPPLY_TYPE"].astype(str).isin(sel_st)]
if len(sel_conn) < len(avail_conn):
    active_df = active_df[active_df["CONNECTION_TYPE"].astype(str).isin(sel_conn)]
if len(sel_cyc) < len(avail_cyc):
    active_df = active_df[active_df["BILL_CYC_CD"].astype(str).isin(sel_cyc)]
if len(sel_mtype) < len(avail_mtype):
    active_df = active_df[active_df["MTR_TYPE_CD"].astype(str).isin(sel_mtype)]
if len(sel_mcfg) < len(avail_mcfg):
    active_df = active_df[active_df["MTR_CONFIG_TY_CD"].astype(str).isin(sel_mcfg)]
if len(sel_make) < len(avail_make):
    active_df = active_df[active_df["MTR_MAKE"].astype(str).isin(sel_make)]
if len(sel_mr_src) < len(avail_mr_src):
    active_df = active_df[active_df["MR_SOURCE_CD"].astype(str).isin(sel_mr_src)]
if len(sel_remark) < len(avail_remark):
    active_df = active_df[active_df["METER_READ_REMARK"].astype(str).isin(sel_remark)]
if len(sel_bb) < len(avail_bb):
    active_df = active_df[active_df["BILL_BASIS"].astype(str).isin(sel_bb)]
if len(sel_inf) < len(avail_inf):
    active_df = active_df[active_df["INF_BILL"].astype(str).isin(sel_inf)]
if len(sel_presys) < len(avail_presys):
    active_df = active_df[active_df[ps_col].astype(str).isin(sel_presys)]

# Calculate KPI Indicators
kpi = calculate_kpi_summary(active_df)

# ----------------- 3. SUMMARY INDICATORS & CLICKABLE DRILLDOWNS (RULE #6 & #8) -----------------
st.markdown("<hr style='margin: 10px 0 14px 0; border: none; border-top: 1px solid #334155;'/>", unsafe_allow_html=True)

kpi_r1_c1, kpi_r1_c2, kpi_r1_c3, kpi_r1_c4 = st.columns(4)
kpi_r2_c1, kpi_r2_c2, kpi_r2_c3, kpi_r2_c4 = st.columns(4)

curr_drill = st.session_state.drilldown_filter

with kpi_r1_c1:
    is_active = (curr_drill == "All")
    st.markdown(f"""
    <div class="kpi-item {'active-drill' if is_active else ''}">
        <div class="kpi-header-label">1. Total Consumers (Distinct DISCOM + ACCT_ID)</div>
        <div class="kpi-primary-val">{format_in(kpi['total_consumers'])}</div>
        <div class="kpi-secondary-info">All qualifying consumers in active scope</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔗 Open & Link Total Consumers", key="kpi_btn_all", use_container_width=True):
        st.session_state.drilldown_filter = "All"
        st.rerun()

with kpi_r1_c2:
    is_active = (curr_drill == "Billed")
    st.markdown(f"""
    <div class="kpi-item {'active-drill' if is_active else ''}">
        <div class="kpi-header-label">2. Billed Consumers</div>
        <div class="kpi-primary-val">{format_in(kpi['billed_consumers'])}</div>
        <div class="kpi-secondary-info">Billing Ratio: <b>{kpi['billing_ratio']:.1f}%</b></div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔗 Open & Link Billed Cohort", key="kpi_btn_billed", use_container_width=True):
        st.session_state.drilldown_filter = "Billed"
        st.rerun()

with kpi_r1_c3:
    is_active = (curr_drill == "Unbilled")
    st.markdown(f"""
    <div class="kpi-item {'active-drill' if is_active else ''}">
        <div class="kpi-header-label">3. Unbilled Consumers</div>
        <div class="kpi-primary-val">{format_in(kpi['unbilled_consumers'])}</div>
        <div class="kpi-secondary-info">Unbilled Ratio: <b>{100.0 - kpi['billing_ratio']:.1f}%</b></div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔗 Open & Link Unbilled Cohort", key="kpi_btn_unbilled", use_container_width=True):
        st.session_state.drilldown_filter = "Unbilled"
        st.rerun()

with kpi_r1_c4:
    st.markdown(f"""
    <div class="kpi-item">
        <div class="kpi-header-label">4. Sold Units (Consumption)</div>
        <div class="kpi-primary-val">{format_in(kpi['sold_units_mu'], 2)} <span style="font-size: 1rem; color: #94A3B8;">MU</span></div>
        <div class="kpi-secondary-info">{format_in(kpi['sold_units_kwh'], 0)} kWh</div>
    </div>
    """, unsafe_allow_html=True)
    st.caption("⚡ Aggregated Current Month Consumption")

with kpi_r2_c1:
    st.markdown(f"""
    <div class="kpi-item">
        <div class="kpi-header-label">5. Current Assessment (CA)</div>
        <div class="kpi-primary-val">{format_currency_dual(kpi['current_assessment'])}</div>
        <div class="kpi-secondary-info">Authentic Monthly Assessment</div>
    </div>
    """, unsafe_allow_html=True)
    st.caption("💰 Authentic Portal CA")

with kpi_r2_c2:
    st.markdown(f"""
    <div class="kpi-item">
        <div class="kpi-header-label">6. CA Excl. Subsidy (CA_EX_SUBS)</div>
        <div class="kpi-primary-val">{format_currency_dual(kpi['ca_ex_subs'])}</div>
        <div class="kpi-secondary-info">Excluding government subsidy</div>
    </div>
    """, unsafe_allow_html=True)
    st.caption("🏷️ Net CA Excl. Subsidy")

with kpi_r2_c3:
    is_active = (curr_drill == "Paid")
    st.markdown(f"""
    <div class="kpi-item {'active-drill' if is_active else ''}">
        <div class="kpi-header-label">7. Total Payment Realization</div>
        <div class="kpi-primary-val">{format_currency_dual(kpi['payment_amount'])}</div>
        <div class="kpi-secondary-info">Paying Consumers: <b>{format_in(kpi['payment_count'])}</b></div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔗 Open & Link Paying Consumers", key="kpi_btn_paid", use_container_width=True):
        st.session_state.drilldown_filter = "Paid"
        st.rerun()

with kpi_r2_c4:
    is_active = (curr_drill == "Outstanding")
    st.markdown(f"""
    <div class="kpi-item {'active-drill' if is_active else ''}">
        <div class="kpi-header-label">8. Total Outstanding Balance</div>
        <div class="kpi-primary-val">{format_currency_dual(kpi['total_outstanding'])}</div>
        <div class="kpi-secondary-info">Cumulative arrears & current dues</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🔗 Open & Link Dues Cohort", key="kpi_btn_outstanding", use_container_width=True):
        st.session_state.drilldown_filter = "Outstanding"
        st.rerun()

# ----------------- 4. VISUALIZATIONS SECTION (UNCLUTTERED TABS) -----------------
st.markdown("<hr style='margin: 12px 0 16px 0; border: none; border-top: 1px solid #334155;'/>", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "📑 Billing Performance",
    "⚡ Energy Consumption Analysis",
    "💰 Revenue & Realization Assessment",
    "📟 Smart Metering & Open Access Analysis"
])

with tab1:
    v1_c1, v1_c2 = st.columns([1, 1.6])
    with v1_c1:
        fig_donut = go.Figure(data=[go.Pie(
            labels=["Billed Consumers", "Unbilled Consumers"],
            values=[kpi["billed_consumers"], kpi["unbilled_consumers"]],
            hole=0.55,
            marker=dict(colors=["#10B981", "#EF4444"]),
            textinfo="label+percent"
        )])
        fig_donut.update_layout(
            title="<b>Billed vs Unbilled Proportions</b>",
            margin=dict(l=10, r=10, t=35, b=10),
            height=300
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    with v1_c2:
        if not active_df.empty:
            grp_disc = active_df.groupby(["DISCOM", "BILLED_FLAG"]).apply(
                lambda g: int(g.groupby(["DISCOM", "ACCT_ID"]).ngroups)
            ).reset_index(name="Consumers")
            fig_disc = px.bar(
                grp_disc,
                x="DISCOM",
                y="Consumers",
                color="BILLED_FLAG",
                barmode="group",
                title="<b>Billed Consumers by DISCOM (Distinct DISCOM + ACCT_ID)</b>",
                color_discrete_map={"Billed": "#10B981", "Unbilled": "#EF4444"},
                text_auto=True
            )
            fig_disc.update_layout(height=300, margin=dict(l=10, r=10, t=35, b=10))
            st.plotly_chart(fig_disc, use_container_width=True)

with tab2:
    v2_c1, v2_c2 = st.columns(2)
    with v2_c1:
        if not active_df.empty:
            grp_e_disc = active_df.groupby("DISCOM")["CONSUMPTION_CURR_MNTH"].sum().reset_index()
            grp_e_disc["MU"] = grp_e_disc["CONSUMPTION_CURR_MNTH"] / 1e6
            fig_e1 = px.bar(
                grp_e_disc,
                x="DISCOM",
                y="MU",
                color="DISCOM",
                title="<b>Sold Units (Million Units MU) by DISCOM</b>",
                text_auto=".2f"
            )
            fig_e1.update_layout(height=320, margin=dict(l=10, r=10, t=35, b=10), showlegend=False)
            st.plotly_chart(fig_e1, use_container_width=True)

    with v2_c2:
        if not active_df.empty:
            grp_e_rng = active_df.groupby("LOAD_RANGE")["CONSUMPTION_CURR_MNTH"].sum().reset_index()
            grp_e_rng["MU"] = grp_e_rng["CONSUMPTION_CURR_MNTH"] / 1e6
            fig_e2 = px.bar(
                grp_e_rng,
                x="LOAD_RANGE",
                y="MU",
                color="LOAD_RANGE",
                title="<b>Sold Units (MU) by Contracted Load Range</b>",
                text_auto=".2f"
            )
            fig_e2.update_layout(height=320, margin=dict(l=10, r=10, t=35, b=10), showlegend=False)
            st.plotly_chart(fig_e2, use_container_width=True)

with tab3:
    v3_c1, v3_c2 = st.columns(2)
    with v3_c1:
        if not active_df.empty:
            grp_rev = active_df.groupby("DISCOM").agg({
                "CURRENT_ASSESSMENT": lambda x: x.sum() / 1e7,
                "TOTAL_PAY_AMT": lambda x: x.sum() / 1e7
            }).reset_index()
            grp_rev = grp_rev.rename(columns={
                "CURRENT_ASSESSMENT": "Assessment (₹ Cr)",
                "TOTAL_PAY_AMT": "Realization (₹ Cr)"
            })
            grp_melt = grp_rev.melt(id_vars="DISCOM", value_vars=["Assessment (₹ Cr)", "Realization (₹ Cr)"], var_name="Metric", value_name="Amount_Cr")
            fig_rev = px.bar(
                grp_melt,
                x="DISCOM",
                y="Amount_Cr",
                color="Metric",
                barmode="group",
                title="<b>Current Assessment vs Payment Realization (₹ Crore)</b>",
                color_discrete_map={"Assessment (₹ Cr)": "#38BDF8", "Realization (₹ Cr)": "#10B981"},
                text_auto=".2f"
            )
            fig_rev.update_layout(height=320, margin=dict(l=10, r=10, t=35, b=10))
            st.plotly_chart(fig_rev, use_container_width=True)

    with v3_c2:
        if not active_df.empty:
            grp_out = active_df.groupby("DISCOM")["TOTAL_OUTSTANDING"].sum().reset_index()
            grp_out["Outstanding_Cr"] = grp_out["TOTAL_OUTSTANDING"] / 1e7
            fig_out = px.bar(
                grp_out,
                x="DISCOM",
                y="Outstanding_Cr",
                color="DISCOM",
                title="<b>Total Outstanding Balance by DISCOM (₹ Crore)</b>",
                text_auto=".2f"
            )
            fig_out.update_layout(height=320, margin=dict(l=10, r=10, t=35, b=10), showlegend=False)
            st.plotly_chart(fig_out, use_container_width=True)

with tab4:
    v4_c1, v4_c2 = st.columns(2)
    with v4_c1:
        if not active_df.empty:
            grp_sm = active_df["SMART_METER_FLAG"].value_counts().reset_index()
            grp_sm.columns = ["Smart_Meter", "Consumers"]
            fig_sm = px.pie(
                grp_sm,
                names="Smart_Meter",
                values="Consumers",
                hole=0.5,
                title="<b>Smart Metered Consumers Distribution</b>",
                color="Smart_Meter",
                color_discrete_map={"YES": "#06B6D4", "NO": "#64748B"}
            )
            fig_sm.update_layout(height=300, margin=dict(l=10, r=10, t=35, b=10))
            st.plotly_chart(fig_sm, use_container_width=True)

    with v4_c2:
        if not active_df.empty:
            grp_oa = active_df["OA_CONFIG"].value_counts().reset_index()
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
            fig_oa.update_layout(height=300, margin=dict(l=10, r=10, t=35, b=10))
            st.plotly_chart(fig_oa, use_container_width=True)

# ----------------- 5. CONSUMER-LEVEL DOSSIER TABLE & STRICT 1:1 PARITY (RULE #6) -----------------
st.markdown("<hr style='margin: 14px 0 16px 0; border: none; border-top: 1px solid #334155;'/>", unsafe_allow_html=True)

table_df = active_df.copy()
drill_type = st.session_state.drilldown_filter

if drill_type == "Billed":
    table_df = table_df[table_df["BILLED_FLAG"] == "Billed"]
elif drill_type == "Unbilled":
    table_df = table_df[table_df["BILLED_FLAG"] == "Unbilled"]
elif drill_type == "Paid":
    table_df = table_df[table_df["PAID_FLAG"] == "Paid"]
elif drill_type == "Outstanding":
    table_df = table_df[table_df["TOTAL_OUTSTANDING"] > 0]

st.markdown(f"""
<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
    <div>
        <span style="font-size: 1.15rem; font-weight: 700; color: #F8FAFC;">📋 Detailed Consumer Records (Strict 1:1 Parity)</span>
        <span style="font-size: 0.82rem; color: #94A3B8; margin-left: 10px;">Active Cohort: <b style="color: #38BDF8;">{drill_type}</b></span>
    </div>
    <div style="font-size: 0.85rem; color: #10B981; font-weight: 600;">
        ✓ Exactly {format_in(len(table_df))} Qualifying Records
    </div>
</div>
""", unsafe_allow_html=True)

t_col1, t_col2, t_col3 = st.columns([3, 2, 2])

with t_col1:
    search_term = st.text_input(
        "Search Table:",
        placeholder="Type Account ID, Consumer Name, Division, or Meter No...",
        label_visibility="collapsed"
    ).strip().upper()

with t_col2:
    col_preset = st.selectbox(
        "Column Preset:",
        ["Standard Executive View", "Billing & Financial", "Technical & Metering", "All Available Columns"],
        index=0,
        label_visibility="collapsed"
    )

with t_col3:
    records_per_page = st.selectbox(
        "Page Size:",
        options=[25, 50, 100, 250, 500, 1000],
        index=1,
        label_visibility="collapsed"
    )

if search_term:
    target_search_fields = [c for c in ["DISCOM_ACCT_ID", "ACCT_ID", "NAME", "DIV_CODE", "DIVISION_NAME", "SERIAL_NBR", "METER_BADGE_NO", "SCNO", "KNO"] if c in table_df.columns]
    mask = pd.Series([False] * len(table_df), index=table_df.index)
    for c in target_search_fields:
        mask = mask | table_df[c].astype(str).str.upper().str.contains(search_term, na=False)
    table_df = table_df[mask]

if col_preset == "Standard Executive View":
    chosen_cols = [
        "DISCOM_ACCT_ID", "ACCT_ID", "NAME", "DISCOM", "ZONE_NAME", "CIRCLE_NAME", "DIVISION_NAME",
        "SUPPLY_TYPE", "TARIFF_CATEGORY", "LOAD", "LOAD_UNIT", "LOAD_KW", "LOAD_RANGE",
        "OPR_FLG", "BILLED_FLAG", "PAID_FLAG", "LAST_BILL_DATE",
        "CONSUMPTION_CURR_MNTH", "CURRENT_ASSESSMENT", "CA_EX_SUBS",
        "TOTAL_PAY_AMT", "TOTAL_OUTSTANDING", "MTR_CONFIG_TY_CD", "OA_CONFIG", "SMART_METER_FLAG"
    ]
elif col_preset == "Billing & Financial":
    chosen_cols = [
        "DISCOM_ACCT_ID", "ACCT_ID", "NAME", "DISCOM", "DIVISION_NAME", "SUPPLY_TYPE", "TARIFF_CATEGORY",
        "LAST_BILL_DATE", "BILLED_FLAG", "CONSUMPTION_CURR_MNTH", "BILLED_AMOUNT",
        "CURRENT_ASSESSMENT", "CA_EX_SUBS", "ENERGY_AMT", "NON_ENERGY_AMT",
        "PAY_AMT", "TOTAL_PAY_AMT", "PAID_FLAG", "PAY_DATE", "TOTAL_OUTSTANDING", "ARREAR"
    ]
elif col_preset == "Technical & Metering":
    chosen_cols = [
        "DISCOM_ACCT_ID", "ACCT_ID", "NAME", "DISCOM", "DIVISION_NAME", "LOAD", "LOAD_UNIT", "LOAD_KW", "LOAD_RANGE",
        "SERIAL_NBR", "MTR_MAKE", "MTR_TYPE_CD", "MTR_CONFIG_TY_CD", "OA_CONFIG",
        "CMPRESYS_CODE", "SMART_METER_FLAG", "OPENING_READING_KWH", "CLOSING_READING_KWH",
        "CONSUMPTION_CURR_MNTH", "POWER_FACTOR", "METER_READ_REMARK"
    ]
else:
    chosen_cols = list(table_df.columns)

final_cols = [c for c in chosen_cols if c in table_df.columns]

total_table_rows = len(table_df)
total_pages = max(1, math.ceil(total_table_rows / records_per_page))

pg_c1, pg_c2, pg_c3 = st.columns([2, 3, 2])
with pg_c2:
    selected_page = st.number_input(f"Page (1 to {total_pages}):", min_value=1, max_value=total_pages, value=1, step=1)

idx_start = (selected_page - 1) * records_per_page
idx_end = min(idx_start + records_per_page, total_table_rows)

st.dataframe(
    table_df.iloc[idx_start:idx_end][final_cols],
    use_container_width=True,
    hide_index=True
)

st.caption(f"Displaying rows {format_in(idx_start + 1 if total_table_rows > 0 else 0)} to {format_in(idx_end)} of {format_in(total_table_rows)} qualifying consumer records.")

d_col1, d_col2 = st.columns(2)

with d_col1:
    excel_bytes = generate_excel_export(table_df[final_cols], title="Open_Access_Consumers")
    st.download_button(
        label=f"📥 Download Exact Records as Excel (.xlsx) [{format_in(len(table_df))} Rows]",
        data=excel_bytes,
        file_name=f"Open_Access_Consumers_{effective_month}_{effective_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with d_col2:
    csv_data = table_df[final_cols].to_csv(index=False).encode("utf-8")
    st.download_button(
        label=f"📥 Download Exact Records as CSV (.csv) [{format_in(len(table_df))} Rows]",
        data=csv_data,
        file_name=f"Open_Access_Consumers_{effective_month}_{effective_year}.csv",
        mime="text/csv",
        use_container_width=True
    )
