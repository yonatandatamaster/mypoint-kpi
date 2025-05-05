import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# --- Upload section ---
st.sidebar.header("Upload Data Files")
scan_file = st.sidebar.file_uploader("Upload Scan File (Excel)", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File (Excel)", type=["xlsx"])

if scan_file and db_file:
    try:
        df_scan = pd.read_excel(scan_file, sheet_name=0)
        df_db = pd.read_excel(db_file, sheet_name=0)

        # Clean & Rename
        df_scan.columns = df_scan.columns.str.strip().str.lower()
        df_db.columns = df_db.columns.str.strip().str.lower()

        df_scan = df_scan.rename(columns={
            'tanggal scan': 'tanggal_scan',
            'id outlet': 'id_outlet',
            'kode program': 'kode_program'
        })
        df_db = df_db.rename(columns={
            'id outlet': 'id_outlet',
            'pic / promotor': 'pic',
            'pic': 'pic',
            'program': 'program',
            'dso': 'dso'
        })

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

        # Standardize
        df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
        df_scan['week_number'] = df_scan['tanggal_scan'].dt.isocalendar().week
        df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
        df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()

        # Build full grid for all outlets x all weeks
        df_merged = pd.merge(df_db, df_scan[['id_outlet', 'week_number']], on='id_outlet', how='left')
        df_merged['week_number'] = df_merged['week_number'].astype('Int64')

        df_all = df_db.copy()
        df_all['key'] = 1
        all_weeks = pd.DataFrame({'week_number': sorted(df_scan['week_number'].dropna().unique()), 'key': 1})
        full_grid = pd.merge(df_all, all_weeks, on='key').drop(columns='key')

        df_status = pd.merge(full_grid, df_scan[['id_outlet', 'week_number']], 
                             on=['id_outlet', 'week_number'], how='left', indicator=True)
        df_status['is_active'] = (df_status['_merge'] == 'both').astype(int)

        # --- Sidebar Filters ---
        st.sidebar.header("🔍 Additional Filters")
        selected_dso = st.sidebar.selectbox("Filter by DSO", sorted(df_status['dso'].dropna().unique()))
        df_status = df_status[df_status['dso'] == selected_dso]

        selected_programs = st.sidebar.multiselect("Filter by Program", 
            sorted(df_status['program'].dropna().unique()), default=sorted(df_status['program'].dropna().unique()))
        df_status = df_status[df_status['program'].isin(selected_programs)]

        selected_weeks = st.sidebar.multiselect("Select Weeks", 
            sorted(df_status['week_number'].dropna().unique()), default=sorted(df_status['week_number'].dropna().unique()))
        df_status = df_status[df_status['week_number'].isin(selected_weeks)]

        selected_pics = st.sidebar.multiselect("Select PIC(s)", 
            sorted(df_status['pic'].dropna().unique()), default=sorted(df_status['pic'].dropna().unique()))
        df_status = df_status[df_status['pic'].isin(selected_pics)]

        # --- Calculate summary ---
        df_summary = df_status.groupby(['pic', 'program', 'week_number']).agg(
            total_outlet=('id_outlet', 'nunique'),
            active_outlet=('is_active', 'sum')
        ).reset_index()
        df_summary['% active'] = (df_summary['active_outlet'] / df_summary['total_outlet'] * 100).round(1)

        pivot_df = df_summary.pivot_table(index=['pic', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()

        # --- Tabs ---
        tab1, tab2, tab3 = st.tabs(["📋 Dashboard Table", "📈 Trends & Charts", "📂 Raw Data"])

        with tab1:
            st.subheader("📌 Weekly Active % by PIC and Program")

            def highlight_low(val):
                try:
                    return 'background-color: #ffcccc' if float(val) < 50 else ''
                except:
                    return ''

            styled_df = pivot_df.style.format("{:.1f}%").applymap(
                highlight_low, subset=pd.IndexSlice[:, pivot_df.columns[2]:]
            )
            st.dataframe(styled_df, use_container_width=True)

            # Excel Export
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                pivot_df.to_excel(writer, index=False, sheet_name="Weekly KPI")
            st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="weekly_kpi_report.xlsx")

        with tab2:
            st.subheader("📈 Weekly Active % Trend by PIC")
            chart = alt.Chart(df_summary).mark_line(point=True).encode(
                x=alt.X('week_number:O', title='Week'),
                y=alt.Y('% active', title='% Active'),
                color='pic',
                tooltip=['pic', 'program', 'week_number', '% active']
            ).properties(width=800, height=400)
            st.altair_chart(chart, use_container_width=True)

            st.subheader("📊 Average Active % by PIC")
            avg_active = df_summary.groupby('pic')['% active'].mean().reset_index().sort_values('% active', ascending=False)
            bar_chart = alt.Chart(avg_active).mark_bar().encode(
                x=alt.X('% active', title='Avg % Active'),
                y=alt.Y('pic', sort='-x'),
                tooltip=['pic', '% active']
            ).properties(width=700, height=400)
            st.altair_chart(bar_chart, use_container_width=True)

            st.subheader("📊 Weekly Active % Distribution by Program")
            stacked_data = df_summary.groupby(['program', 'week_number'])['% active'].mean().reset_index()
            stacked_chart = alt.Chart(stacked_data).mark_bar().encode(
                x=alt.X('week_number:O', title='Week'),
                y=alt.Y('% active', stack='normalize'),
                color='program',
                tooltip=['program', 'week_number', '% active']
            ).properties(width=800, height=400)
            st.altair_chart(stacked_chart, use_container_width=True)

        with tab3:
            st.subheader("📂 Merged Raw Outlet-Scan Data")
            st.dataframe(df_status.drop(columns=['key', '_merge']), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
else:
    st.info("📂 Please upload both Scan and Database Excel files to begin.")
