import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io
from datetime import datetime, timedelta

# Load fixed master database once
df_db = pd.read_excel("Master_Database_Outlet.xlsx")
df_db.columns = df_db.columns.str.strip().str.lower()
df_db = df_db.rename(columns={
    'id outlet': 'id_outlet',
    'pic / promotor': 'pic',
    'program': 'program',
    'dso ': 'dso'
})
df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()
required_cols_db = ['id_outlet', 'pic', 'program', 'dso']
df_db = df_db[[col for col in df_db.columns if col in required_cols_db]]

# Define week number with Saturday as week start (Excel WEEKNUM(..., 16))
def custom_weeknum(date):
    ref_sat = datetime(2025, 1, 4)  # First Saturday in 2025
    if pd.isna(date):
        return np.nan
    delta_days = (date - ref_sat).days
    return delta_days // 7 + 1 if delta_days >= 0 else 0

# UI Setup
st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# Upload section
with st.sidebar:
    st.header("📤 Upload Scan File")
    scan_file = st.file_uploader("Upload Scan File (Excel)", type=["xlsx"])

# Tabs
tabs = st.tabs(["📊 Dashboard Table", "📈 Trends & Charts", "📂 Raw Data", "📞 Multi-Outlet Scans"])

if scan_file:
    try:
        df_scan = pd.read_excel(scan_file, sheet_name=None)
        df_scan = df_scan[list(df_scan.keys())[0]]
        df_scan.columns = df_scan.columns.str.strip().str.lower()

        col_map = {
            'tanggal scan': 'tanggal_scan',
            'id outlet': 'id_outlet',
            'kode program': 'kode_program',
            'no wa': 'no_hp',
            'no_hp': 'no_hp'  # fallback
        }
        df_scan = df_scan.rename(columns={k: v for k, v in col_map.items() if k in df_scan.columns})
        required_cols_scan = ['tanggal_scan', 'id_outlet']
        for col in required_cols_scan:
            if col not in df_scan.columns:
                st.error(f"Missing column '{col}' in Scan File")
                st.stop()

        # Clean & prep
        df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
        df_scan['week_number'] = df_scan['tanggal_scan'].apply(custom_weeknum)
        df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
        df_scan['no_hp'] = df_scan['no_hp'].astype(str).str.strip() if 'no_hp' in df_scan.columns else ""

        df_merged = pd.merge(df_db, df_scan[['id_outlet', 'week_number', 'no_hp']], on='id_outlet', how='left')

        # Filter sidebar
        with st.sidebar:
            st.header("🔍 Additional Filters")
            selected_dso = st.selectbox("Filter by DSO", sorted(df_db['dso'].dropna().unique()))
            df_filtered = df_merged[df_merged['dso'] == selected_dso]

            programs = sorted(df_filtered['program'].dropna().unique())
            selected_programs = st.multiselect("Filter by Program", options=programs, default=programs)
            df_filtered = df_filtered[df_filtered['program'].isin(selected_programs)]

            weeks = sorted(df_filtered['week_number'].dropna().unique())
            selected_weeks = st.multiselect("Select Weeks", options=weeks, default=weeks)
            df_filtered = df_filtered[df_filtered['week_number'].isin(selected_weeks)]

            pics = sorted(df_filtered['pic'].dropna().unique())
            selected_pics = st.multiselect("Select PIC(s)", options=pics, default=pics)
            df_filtered = df_filtered[df_filtered['pic'].isin(selected_pics)]

        # ============ Weekly % Active by PIC & Program ============
        df_filtered['is_active'] = df_filtered['week_number'].notna()
        weekly_summary = df_filtered.groupby(['pic', 'program', 'week_number']).agg(
            active_outlet=('id_outlet', 'nunique')
        ).reset_index()

        total_outlets = df_filtered[['pic', 'program', 'id_outlet']].drop_duplicates()
        total_outlets = total_outlets.groupby(['pic', 'program'])['id_outlet'].nunique().reset_index(name='total_outlets')

        df_all = pd.merge(weekly_summary, total_outlets, on=['pic', 'program'])
        df_all['% active'] = (df_all['active_outlet'] / df_all['total_outlets'] * 100).round(1)

        pivot_df = df_all.pivot_table(index=['pic', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()

        def highlight_low(val):
            try:
                return 'background-color: #ffcccc' if float(val) < 50 else ''
            except:
                return ''

        styled_df = pivot_df.style.format("{:.1f}%", na_rep="").applymap(highlight_low, subset=pd.IndexSlice[:, pivot_df.columns[2]:])

        with tabs[0]:
            st.subheader("📌 Weekly Active % by PIC and Program")
            st.dataframe(styled_df, use_container_width=True)

            # Export
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                pivot_df.to_excel(writer, index=False, sheet_name="Weekly KPI")
            st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="weekly_kpi_report.xlsx")

        # ============ Trend Charts ============
        with tabs[1]:
            st.subheader("📈 Weekly Active % Trend by PIC")
            line_chart = alt.Chart(df_all).mark_line(point=True).encode(
                x=alt.X('week_number:O', title='Week'),
                y=alt.Y('% active', title='% Active'),
                color='pic',
                tooltip=['pic', 'program', 'week_number', '% active']
            ).properties(width=800, height=400)
            st.altair_chart(line_chart, use_container_width=True)

            st.subheader("📊 Avg % Active by PIC")
            avg_active = df_all.groupby('pic')['% active'].mean().reset_index().sort_values('% active', ascending=False)
            bar_chart = alt.Chart(avg_active).mark_bar().encode(
                x=alt.X('% active', title='Avg % Active'),
                y=alt.Y('pic', sort='-x'),
                tooltip=['pic', '% active']
            ).properties(width=700, height=400)
            st.altair_chart(bar_chart, use_container_width=True)

        # ============ Raw Merged ============
        with tabs[2]:
            st.subheader("📂 Raw Merged Data")
            st.dataframe(df_filtered, use_container_width=True)

        # ============ Multi-Outlet Scans ============
        with tabs[3]:
            st.subheader("📞 Phone Numbers Scanning Multiple Outlets")
            df_multi = df_scan[['no_hp', 'id_outlet']].dropna()
            df_multi = df_multi.groupby('no_hp')['id_outlet'].nunique().reset_index(name='num_outlets')
            df_multi = df_multi[df_multi['num_outlets'] > 1].sort_values(by='num_outlets', ascending=False)
            st.dataframe(df_multi, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
else:
    st.info("Please upload a Scan Excel file to begin.")
