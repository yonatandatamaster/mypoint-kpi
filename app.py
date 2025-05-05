import streamlit as st
import pandas as pd
import altair as alt
import io

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# Uploads
st.sidebar.header("Upload Data Files")
scan_file = st.sidebar.file_uploader("Upload Scan File (Excel)", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File (Excel)", type=["xlsx"])

if scan_file and db_file:
    try:
        df_scan = pd.read_excel(scan_file, sheet_name=0)
        df_db = pd.read_excel(db_file, sheet_name=0)

        # Clean columns
        df_scan.columns = df_scan.columns.str.strip().str.lower()
        df_db.columns = df_db.columns.str.strip().str.lower()

        # Rename relevant columns
        df_scan.rename(columns={
            'tanggal scan': 'tanggal_scan',
            'id outlet': 'id_outlet',
            'kode program': 'kode_program'
        }, inplace=True)

        df_db.rename(columns={
            'id outlet': 'id_outlet',
            'pic / promotor': 'pic',
            'pic': 'pic',
            'program': 'program',
            'dso': 'dso'
        }, inplace=True)

        # Validation
        required_cols_scan = ['tanggal_scan', 'id_outlet']
        required_cols_db = ['id_outlet', 'pic', 'program', 'dso']
        for col in required_cols_scan + required_cols_db:
            if col not in df_scan.columns and col not in df_db.columns:
                st.error(f"Missing column: {col}")
                st.stop()

        df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
        df_scan['week_number'] = df_scan['tanggal_scan'].dt.isocalendar().week
        df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
        df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()

        df_merged = pd.merge(df_db, df_scan[['id_outlet', 'week_number']], on='id_outlet', how='left')
        df_merged['week_number'] = df_merged['week_number'].fillna(0).astype(int)

        # Filter
        st.sidebar.header("🔍 Additional Filters")
        selected_dso = st.sidebar.selectbox("Filter by DSO", sorted(df_merged['dso'].dropna().unique()))
        df_filtered = df_merged[df_merged['dso'] == selected_dso]

        selected_programs = st.sidebar.multiselect("Filter by Program", sorted(df_filtered['program'].unique()), default=sorted(df_filtered['program'].unique()))
        df_filtered = df_filtered[df_filtered['program'].isin(selected_programs)]

        selected_weeks = st.sidebar.multiselect("Select Weeks", sorted(df_filtered['week_number'].unique()), default=sorted(df_filtered['week_number'].unique()))
        df_filtered = df_filtered[df_filtered['week_number'].isin(selected_weeks)]

        selected_pics = st.sidebar.multiselect("Select PIC(s)", sorted(df_filtered['pic'].unique()), default=sorted(df_filtered['pic'].unique()))
        df_filtered = df_filtered[df_filtered['pic'].isin(selected_pics)]

        # Pivot Table
        total_outlets = df_filtered.groupby(['pic', 'program'])['id_outlet'].nunique().reset_index().rename(columns={'id_outlet': 'total_outlets'})
        active_counts = df_filtered[df_filtered['week_number'] > 0].groupby(['pic', 'program', 'week_number'])['id_outlet'].nunique().reset_index().rename(columns={'id_outlet': 'active_outlets'})

        result = pd.merge(active_counts, total_outlets, on=['pic', 'program'], how='left')
        result['% active'] = (result['active_outlets'] / result['total_outlets'] * 100).round(1)

        pivot_df = result.pivot_table(index=['pic', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()

        st.subheader("📌 Weekly Active % by PIC and Program")

        def highlight_low(val):
            try:
                return 'background-color: #ffcccc' if float(val) < 50 else ''
            except:
                return ''

        week_cols = pivot_df.columns[2:]
        styled_df = pivot_df.style.format({col: "{:.1f}%" for col in week_cols}).applymap(highlight_low, subset=pd.IndexSlice[:, week_cols])
        st.dataframe(styled_df, use_container_width=True)

        # Excel Export
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
            pivot_df.to_excel(writer, index=False, sheet_name="Weekly KPI")
        st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="weekly_kpi_report.xlsx")

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("Please upload both Excel files to begin.")
