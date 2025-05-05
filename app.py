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

        # Ensure columns exist
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
        df_scan['week_number'] = df_scan['tanggal_scan'].dt.isocalendar().week
        df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
        df_scan['kode_program'] = df_scan['kode_program'].astype(str).str.strip()
        df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()

        df_merged = pd.merge(df_db, df_scan, on='id_outlet', how='left')
        df_merged['is_active'] = df_merged['tanggal_scan'].notna()

        # --- Program filter widget ---
        program_options = df_merged['program'].dropna().unique()
        selected_programs = st.sidebar.multiselect("Filter by Program", sorted(program_options), default=sorted(program_options))
        df_merged = df_merged[df_merged['program'].isin(selected_programs)]

        # --- KPI calculations ---
        active_summary = df_merged.groupby(['pic', 'id_outlet']).agg({'is_active': 'max'}).reset_index()
        percent_active = active_summary.groupby('pic')['is_active'].mean().reset_index()
        percent_active['% active outlets'] = (percent_active['is_active'] * 100).round(1)
        percent_active = percent_active.drop(columns=['is_active'])

        st.subheader("📌 % Active Outlet by PIC / Promotor")
        st.dataframe(percent_active, use_container_width=True)

        chart = alt.Chart(percent_active).mark_bar().encode(
            x=alt.X('% active outlets', title='% Active Outlets'),
            y=alt.Y('pic', sort='-x'),
            tooltip=['pic', '% active outlets']
        ).properties(
            width=700,
            height=400,
            title="% Active Outlets by PIC / Promotor"
        )

        st.altair_chart(chart, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
else:
    st.info("Please upload both Scan and Database Excel files to begin.")
