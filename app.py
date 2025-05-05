import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

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

        df_scan.columns = df_scan.columns.str.strip().str.lower()
        df_db.columns = df_db.columns.str.strip().str.lower()

        scan_col_map = {
            'tanggal scan': 'tanggal_scan',
            'id outlet': 'id_outlet',
            'kode program': 'kode_program'
        }
        db_col_map = {
            'id outlet': 'id_outlet',
            'pic': 'pic',
            'pic / promotor': 'pic',
            'program': 'program',
            'dso': 'dso'
        }

        df_scan = df_scan.rename(columns=scan_col_map)
        df_db = df_db.rename(columns=db_col_map)

        required_cols_scan = ['tanggal_scan', 'id_outlet', 'kode_program']
        required_cols_db = ['id_outlet', 'pic', 'program', 'dso']

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

        # Filter UI
        st.sidebar.header("🔍 Additional Filters")
        selected_dso = st.sidebar.selectbox("Filter by DSO", options=sorted(df_merged['dso'].dropna().unique()))
        df_filtered = df_merged[df_merged['dso'] == selected_dso]

        selected_programs = st.sidebar.multiselect("Filter by Program", options=sorted(df_filtered['program'].dropna().unique()), default=sorted(df_filtered['program'].dropna().unique()))
        df_filtered = df_filtered[df_filtered['program'].isin(selected_programs)]

        selected_weeks = st.sidebar.multiselect("Select Weeks", options=sorted(df_filtered['week_number'].dropna().unique()), default=sorted(df_filtered['week_number'].dropna().unique()))
        df_filtered = df_filtered[df_filtered['week_number'].isin(selected_weeks)]

        selected_pics = st.sidebar.multiselect("Select PIC(s)", options=sorted(df_filtered['pic'].dropna().unique()), default=sorted(df_filtered['pic'].dropna().unique()))
        df_filtered = df_filtered[df_filtered['pic'].isin(selected_pics)]

        # Compute active % per week
        summary = df_filtered.groupby(['pic', 'program', 'week_number'])['is_active'].mean().reset_index()
        summary['% active'] = (summary['is_active'] * 100).round(1)
        pivot_df = summary.pivot_table(index=['pic', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()

        st.subheader("📌 Weekly Active % by PIC and Program")

        def highlight_low(val):
            if isinstance(val, (int, float)) and val < 50:
                return 'background-color: #ffcccc'
            return ''

        styled_df = pivot_df.style.format("{:.1f}%").applymap(highlight_low, subset=pd.IndexSlice[:, pivot_df.columns[2]:])
        st.dataframe(styled_df, use_container_width=True)

        # Export button
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
            pivot_df.to_excel(writer, index=False, sheet_name="Weekly KPI")
        st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="weekly_kpi_report.xlsx")

        # Weekly trend line chart
        st.subheader("📈 Weekly Active % Trend by PIC")
        trend_data = summary.copy()
        chart = alt.Chart(trend_data).mark_line(point=True).encode(
            x=alt.X('week_number:O', title='Week'),
            y=alt.Y('% active', title='% Active'),
            color='pic',
            tooltip=['pic', 'program', 'week_number', '% active']
        ).properties(
            width=800,
            height=400,
            title="Active % Trend by PIC"
        )

        st.altair_chart(chart, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
else:
    st.info("Please upload both Scan and Database Excel files to begin.")
