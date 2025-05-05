import streamlit as st
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import altair as alt
from io import BytesIO

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
        df_scan['minggu'] = df_scan['tanggal_scan'].dt.isocalendar().week
        df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
        df_scan['kode_program'] = df_scan['kode_program'].astype(str).str.strip()
        df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()

        df_merged = pd.merge(df_db, df_scan, on='id_outlet', how='left')
        df_merged['is_active'] = df_merged['tanggal_scan'].notna()

        # Group by minggu, pic, program
        grouped = df_merged.groupby(['pic', 'program', 'minggu', 'id_outlet']).agg({'is_active': 'max'}).reset_index()
        weekly = grouped.groupby(['pic', 'program', 'minggu'])['is_active'].mean().reset_index()
        weekly['active_percent'] = (weekly['is_active'] * 100).round(1)

        # Pivot table
        df_pivot = weekly.pivot_table(index=['pic', 'program'], columns='minggu', values='active_percent').reset_index()

        st.subheader(":pushpin: Weekly Active % by PIC and Program")
        st.dataframe(df_pivot, use_container_width=True)

        # --- Convert to long format for trend visualization ---
        df_long = df_pivot.melt(id_vars=['pic', 'program'], var_name='minggu', value_name='active_percent')
        df_long['active_percent'] = df_long['active_percent'].fillna(0)

        # --- Sidebar Filters ---
        st.sidebar.markdown("### 🔎 Additional Filters")
        available_programs = sorted(df_long['program'].dropna().unique())
        selected_programs = st.sidebar.multiselect("Filter by Program", available_programs, default=available_programs)

        available_weeks = sorted(df_long['minggu'].dropna().unique())
        selected_weeks = st.sidebar.multiselect("Select Weeks", available_weeks, default=available_weeks)

        available_pics = sorted(df_long['pic'].dropna().unique())
        selected_pics = st.sidebar.multiselect("Select PIC(s)", available_pics, default=available_pics)

        # --- Apply Filters ---
        df_filtered = df_long[
            df_long['program'].isin(selected_programs) &
            df_long['minggu'].isin(selected_weeks) &
            df_long['pic'].isin(selected_pics)
        ]

        # --- Line Chart ---
        st.subheader("📈 Weekly Active % Trend by PIC")
        line_chart = alt.Chart(df_filtered).mark_line(point=True).encode(
            x=alt.X('minggu:N', title='Minggu'),
            y=alt.Y('active_percent:Q', title='% Aktif', scale=alt.Scale(domain=[0, 100])),
            color='pic:N',
            tooltip=['pic', 'program', 'minggu', 'active_percent']
        ).properties(
            width=800,
            height=400
        )
        st.altair_chart(line_chart, use_container_width=True)

        # --- Download Excel ---
        def to_excel(df):
            output = BytesIO()
            df.to_excel(output, index=False)
            return output.getvalue()

        st.download_button(
            label="📅 Download Filtered Data",
            data=to_excel(df_filtered),
            file_name="filtered_kpi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
else:
    st.info("Please upload both Scan and Database Excel files to begin.")
