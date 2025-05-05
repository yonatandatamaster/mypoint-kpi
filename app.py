import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# Upload Section
st.sidebar.header("📤 Upload Files")
scan_file = st.sidebar.file_uploader("Upload Scan File", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File", type=["xlsx"])

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
            'kode program': 'kode_program',
            'no hp': 'no_hp'
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

        df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
        df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
        df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()

        # Custom Weeknum (Saturday start - Excel WEEKNUM(..., 16))
        df_scan['custom_week'] = ((df_scan['tanggal_scan'] - pd.to_timedelta((df_scan['tanggal_scan'].dt.weekday + 2) % 7, unit='d'))
                                  .dt.isocalendar().week)

        df_merged = pd.merge(df_db, df_scan, on='id_outlet', how='left')
        df_merged['is_active'] = df_merged['tanggal_scan'].notna()

        # Filter UI
        st.sidebar.header("🔍 Filter Data")
        selected_dso = st.sidebar.selectbox("Filter by DSO", options=sorted(df_merged['dso'].dropna().unique()))
        df_filtered = df_merged[df_merged['dso'] == selected_dso]

        selected_programs = st.sidebar.multiselect("Filter by Program", sorted(df_filtered['program'].dropna().unique()), default=sorted(df_filtered['program'].dropna().unique()))
        df_filtered = df_filtered[df_filtered['program'].isin(selected_programs)]

        selected_pics = st.sidebar.multiselect("Filter by PIC", sorted(df_filtered['pic'].dropna().unique()), default=sorted(df_filtered['pic'].dropna().unique()))
        df_filtered = df_filtered[df_filtered['pic'].isin(selected_pics)]

        selected_weeks = st.sidebar.multiselect("Filter by Week", sorted(df_filtered['custom_week'].dropna().unique()), default=sorted(df_filtered['custom_week'].dropna().unique()))
        df_filtered = df_filtered[df_filtered['custom_week'].isin(selected_weeks)]

        # Compute weekly active %
        summary = df_filtered.groupby(['pic', 'program', 'custom_week']).agg(
            active_outlets=('id_outlet', lambda x: x[df_filtered.loc[x.index, 'is_active']].nunique()),
            total_outlets=('id_outlet', 'nunique')
        ).reset_index()
        summary['% active'] = (summary['active_outlets'] / summary['total_outlets'] * 100).round(1)

        pivot_df = summary.pivot_table(index=['pic', 'program'], columns='custom_week', values='% active', fill_value=0).reset_index()

        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📌 Dashboard", "📈 Trends & Charts", "🗂 Raw Data", "📱 Multi-Outlet Scans"])

        with tab1:
            st.subheader("📌 Weekly Active % by PIC and Program")

            def highlight_low(val):
                return 'background-color: #ffcccc' if isinstance(val, (int, float)) and val < 50 else ''

            styled_df = pivot_df.style.format("{:.1f}%").applymap(highlight_low, subset=pd.IndexSlice[:, pivot_df.columns[2]:])
            st.dataframe(styled_df, use_container_width=True)

            # Export button
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                pivot_df.to_excel(writer, index=False, sheet_name="Weekly KPI")
            st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="weekly_kpi_report.xlsx")

        with tab2:
            st.subheader("📈 Weekly Trend by PIC")
            trend_chart = alt.Chart(summary).mark_line(point=True).encode(
                x=alt.X('custom_week:O', title='Week'),
                y=alt.Y('% active', title='% Active'),
                color='pic',
                tooltip=['pic', 'program', 'custom_week', '% active']
            ).properties(width=800, height=400)
            st.altair_chart(trend_chart, use_container_width=True)

            st.subheader("📊 Average Active % by PIC")
            avg_df = summary.groupby('pic')['% active'].mean().reset_index().sort_values('% active', ascending=False)
            bar_chart = alt.Chart(avg_df).mark_bar().encode(
                x=alt.X('% active', title='Avg % Active'),
                y=alt.Y('pic', sort='-x'),
                tooltip=['pic', '% active']
            ).properties(width=700, height=400)
            st.altair_chart(bar_chart, use_container_width=True)

            st.subheader("📊 Weekly Distribution by Program")
            program_chart_data = summary.groupby(['program', 'custom_week'])['% active'].mean().reset_index()
            stacked = alt.Chart(program_chart_data).mark_bar().encode(
                x=alt.X('custom_week:O', title='Week'),
                y=alt.Y('% active', stack='normalize'),
                color='program',
                tooltip=['program', 'custom_week', '% active']
            ).properties(width=800, height=400)
            st.altair_chart(stacked, use_container_width=True)

        with tab3:
            st.subheader("🗂 Raw Merged Data")
            st.dataframe(df_filtered, use_container_width=True)

        with tab4:
            st.subheader("📱 Phone Numbers Scanning Multiple Outlets")
            multi_outlet = df_filtered.dropna(subset=['no_hp'])
            scan_counts = multi_outlet.groupby(['no_hp'])['id_outlet'].nunique().reset_index()
            scan_counts = scan_counts[scan_counts['id_outlet'] > 1].sort_values('id_outlet', ascending=False)
            st.dataframe(scan_counts.rename(columns={'no_hp': 'Phone Number', 'id_outlet': 'Outlets Scanned'}), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("Please upload both Scan and Database Excel files to proceed.")
