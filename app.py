import streamlit as st
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import altair as alt

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# --- File upload section ---
st.sidebar.header("Upload Data Files")
scan_file = st.sidebar.file_uploader("Upload Scan File (Excel)", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File (Excel)", type=["xlsx"])

if scan_file and db_file:
    try:
        df_scan = pd.read_excel(scan_file, sheet_name=None)
        df_db = pd.read_excel(db_file, sheet_name=None)

        df_scan = df_scan[list(df_scan.keys())[0]]
        df_db = df_db[list(df_db.keys())[0]]

        # Standardize column names
        df_scan.columns = df_scan.columns.str.strip().str.lower()
        df_db.columns = df_db.columns.str.strip().str.lower()

        # Detect and rename essential columns
        scan_col_map = {
            'tanggal scan': 'tanggal_scan',
            'id outlet': 'id_outlet',
            'kode program': 'kode_program'
        }
        db_col_map = {
            'id outlet': 'id_outlet',
            'pic': 'pic',
            'pic / promotor': 'pic',
            'program': 'program'
        }
        df_scan = df_scan.rename(columns=scan_col_map)
        df_db = df_db.rename(columns=db_col_map)

        # Ensure required columns
        required_cols_scan = ['tanggal_scan', 'id_outlet', 'kode_program']
        required_cols_db = ['id_outlet', 'pic', 'program']

        for col in required_cols_scan:
            if col not in df_scan.columns:
                st.error(f"Missing column '{col}' in Scan File")
                st.stop()

        for col in required_cols_db:
            if col not in df_db.columns:
                st.error(f"Missing column '{col}' in Database File")
                st.stop()

        df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
        df_scan['minggu'] = df_scan['tanggal_scan'].dt.isocalendar().week
        df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
        df_scan['kode_program'] = df_scan['kode_program'].astype(str).str.strip()
        df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()

        df_merged = pd.merge(df_db, df_scan, on='id_outlet', how='left')
        df_merged['is_active'] = df_merged['tanggal_scan'].notna()

        # --- Streamlit filter by program ---
        selected_program = st.sidebar.multiselect("Filter by Program", options=sorted(df_db['program'].dropna().unique()))
        if selected_program:
            df_merged = df_merged[df_merged['program'].isin(selected_program)]

        # --- Weekly active outlets ---
        weekly = df_merged.groupby(['pic', 'program', 'minggu', 'id_outlet']).agg({'is_active': 'max'}).reset_index()
        weekly_summary = weekly.groupby(['pic', 'program', 'minggu']).agg(
            total_outlet=('id_outlet', 'count'),
            aktif=('is_active', 'sum')
        ).reset_index()
        weekly_summary['% aktif'] = (weekly_summary['aktif'] / weekly_summary['total_outlet'] * 100).round(1).astype(str) + '%'

        # --- Pivot table like Excel ---
        pivot_table = weekly_summary.pivot_table(
            index=['pic', 'program'],
            columns='minggu',
            values='% aktif',
            aggfunc='first',
            fill_value='0%'
        ).reset_index()

        st.subheader("📌 Weekly Active % by PIC and Program")
        st.dataframe(pivot_table, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
else:
    st.info("Please upload both Scan and Database Excel files to begin.")
