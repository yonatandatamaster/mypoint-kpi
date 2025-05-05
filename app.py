import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# Sidebar upload
st.sidebar.header("Upload Data Files")
scan_file = st.sidebar.file_uploader("Upload Scan File (Excel)", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File (Excel)", type=["xlsx"])

if scan_file and db_file:
    try:
        # Load and clean data
        df_scan = pd.read_excel(scan_file)
        df_db = pd.read_excel(db_file)

        df_scan.columns = df_scan.columns.str.strip().str.lower()
        df_db.columns = df_db.columns.str.strip().str.lower()

        df_scan = df_scan.rename(columns={
            'tanggal scan': 'tanggal_scan',
            'id outlet': 'id_outlet',
            'nomor hp': 'nomor_hp'
        })
        df_db = df_db.rename(columns={
            'id outlet': 'id_outlet',
            'pic / promotor': 'pic',
            'program': 'program',
            'dso': 'dso'
        })

        df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
        df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()

        df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
        df_scan['adjusted_date'] = df_scan['tanggal_scan'] - pd.to_timedelta((df_scan['tanggal_scan'].dt.dayofweek + 2) % 7, unit='D')
        df_scan['week_number'] = df_scan['adjusted_date'].dt.isocalendar().week

        df_merged = pd.merge(df_db, df_scan, on='id_outlet', how='left')
        df_merged['is_active'] = df_merged['tanggal_scan'].notna()

        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📌 Dashboard Table", "📈 Trends & Charts", "📂 Raw Data", "📱 Multi-Outlet Scans"])

        with tab1:
            st.subheader("📌 Weekly Active % by PIC and Program")

            selected_dso = st.selectbox("Select DSO", sorted(df_merged['dso'].dropna().unique()))
            df_filtered = df_merged[df_merged['dso'] == selected_dso]

            programs = sorted(df_filtered['program'].dropna().unique())
            selected_programs = st.multiselect("Filter Program", programs, default=programs)
            df_filtered = df_filtered[df_filtered['program'].isin(selected_programs)]

            # Weekly active calculation (true logic)
            total_outlets = df_filtered[['pic', 'id_outlet']].drop_duplicates().groupby('pic').count().rename(columns={'id_outlet': 'total_outlets'}).reset_index()
            weekly_scan = df_filtered.dropna(subset=['tanggal_scan'])[['pic', 'week_number', 'id_outlet']].drop_duplicates()
            scanned = weekly_scan.groupby(['pic', 'week_number']).count().reset_index().rename(columns={'id_outlet': 'scanned_outlets'})
            result = pd.merge(scanned, total_outlets, on='pic', how='left')
            result['% active'] = (result['scanned_outlets'] / result['total_outlets'] * 100).round(1)

            pivot_df = result.pivot_table(index='pic', columns='week_number', values='% active', fill_value=0).reset_index()

            def highlight(val):
                return 'background-color: #ffcccc' if val < 50 else ''

            styled = pivot_df.style.format("{:.1f}%").applymap(highlight, subset=pd.IndexSlice[:, pivot_df.columns[1]:])
            st.dataframe(styled, use_container_width=True)

            # Excel export
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                pivot_df.to_excel(writer, index=False, sheet_name="KPI Summary")
            st.download_button("📥 Download Excel Summary", towrite.getvalue(), "weekly_kpi_summary.xlsx")

        with tab2:
            st.subheader("📈 Weekly Active % Trend by PIC")
            trend_chart = alt.Chart(result).mark_line(point=True).encode(
                x=alt.X('week_number:O', title="Week"),
                y=alt.Y('% active', title="% Active"),
                color='pic',
                tooltip=['pic', 'week_number', '% active']
            ).properties(width=800, height=400)
            st.altair_chart(trend_chart, use_container_width=True)

            st.subheader("📊 Avg % Active by PIC")
            avg_active = result.groupby('pic')['% active'].mean().reset_index()
            bar_chart = alt.Chart(avg_active).mark_bar().encode(
                x=alt.X('% active', title="Avg % Active"),
                y=alt.Y('pic', sort='-x'),
                tooltip=['pic', '% active']
            ).properties(width=700, height=400)
            st.altair_chart(bar_chart, use_container_width=True)

        with tab3:
            st.subheader("📂 Merged Raw Data")
            st.dataframe(df_merged.head(100), use_container_width=True)

        with tab4:
            st.subheader("📱 Phone Numbers Scanning in Multiple Outlets")
            df_phone = df_scan[['nomor_hp', 'id_outlet']].dropna()
            dupes = df_phone.drop_duplicates().groupby('nomor_hp')['id_outlet'].nunique().reset_index()
            dupes = dupes[dupes['id_outlet'] > 1].sort_values('id_outlet', ascending=False)

            st.dataframe(dupes.rename(columns={'id_outlet': 'unique_outlets_scanned'}), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("Please upload both Scan and Database files to begin.")
