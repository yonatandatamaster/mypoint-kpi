import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

DB_FILE = "Master_Database_Outlet.xlsx"
WEEK_TAG_FILE = "Date_week_tag.xlsx"

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")

@st.cache_data
def load_database():
    df = pd.read_excel(DB_FILE)
    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={
        'id outlet': 'id_outlet',
        'nama outlet (dsca)': 'nama_outlet',
        'pic / promotor': 'pic',
        'program': 'program',
        'dso ': 'dso'
    })
    df['id_outlet'] = df['id_outlet'].astype(str).str.strip()
    return df

@st.cache_data
def load_week_tags():
    df = pd.read_excel(WEEK_TAG_FILE)
    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={'date': 'tanggal_scan', 'week': 'week_number'})
    df['tanggal_scan'] = pd.to_datetime(df['tanggal_scan'], errors='coerce')
    return df

def highlight_low(val):
    try:
        return 'background-color: #ffcccc' if float(val) < 50 else ''
    except:
        return ''

st.sidebar.header("📂 Upload Scan File")
scan_file = st.sidebar.file_uploader("Upload Scan File (.xlsx)", type=["xlsx"])

if scan_file:
    df_db = load_database()
    week_map = load_week_tags()

    df_scan = pd.read_excel(scan_file, sheet_name=None)
    df_scan = df_scan[list(df_scan.keys())[0]]
    df_scan.columns = df_scan.columns.str.strip().str.lower()
    df_scan = df_scan.rename(columns={
        'tanggal scan': 'tanggal_scan',
        'id outlet': 'id_outlet',
        'kode program': 'kode_program',
        'site': 'site',
        'no wa': 'no_hp'
    })
    df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
    df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
    df_scan['no_hp'] = df_scan['no_hp'].astype(str).str.strip()
    df_scan = pd.merge(df_scan, week_map, on='tanggal_scan', how='left')

    df_merged = pd.merge(df_db, df_scan, on='id_outlet', how='left')
    df_merged['is_active'] = df_merged['tanggal_scan'].notna()

    # Sidebar filters
    st.sidebar.header("🔍 Filter Data")
    selected_dso = st.sidebar.selectbox("Filter by DSO", sorted(df_db['dso'].dropna().unique()))
    df_filtered = df_merged[df_merged['dso'] == selected_dso]

    selected_programs = st.sidebar.multiselect("Filter by Program", sorted(df_filtered['program'].dropna().unique()), default=sorted(df_filtered['program'].dropna().unique()))
    df_filtered = df_filtered[df_filtered['program'].isin(selected_programs)]

    all_weeks = sorted(df_filtered['week_number'].dropna().unique())
    selected_weeks = st.sidebar.multiselect("Select Weeks", all_weeks, default=all_weeks)

    selected_pics = st.sidebar.multiselect("Select PIC(s)", sorted(df_filtered['pic'].dropna().unique()), default=sorted(df_filtered['pic'].dropna().unique()))
    df_filtered = df_filtered[df_filtered['pic'].isin(selected_pics)]

    # Tabs (sorted alphabetically)
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Dashboard Table Unique Konsumen",
        "📊 DSO Summary Table",
        "🚫 List of Inactive Outlets",
        "📊 Dashboard Table",
        "📈 Trends & Charts",
        "📋 Multi-Outlet Scans"
    ])

    # --- Tab 1: Unique Konsumen Table ---
    with tab1:
        st.subheader("📊 Weekly Unique Consumers per Outlet")
        df_weeks = df_filtered[df_filtered['week_number'].isin(selected_weeks)]
        df_unique = df_weeks[df_weeks['no_hp'].notna()]
        
        # Weekly unique per outlet
        weekly_unique = df_unique.groupby(['week_number', 'pic', 'program', 'id_outlet'])['no_hp'].nunique().reset_index(name='unique_users')
        pivot_unique = weekly_unique.pivot_table(index=['pic', 'program', 'id_outlet'], columns='week_number', values='unique_users', fill_value=0).reset_index()
        
        # Total unique konsumen across all weeks
        total_unique = df_unique.groupby(['id_outlet'])['no_hp'].nunique().reset_index(name='total_unique_konsumen')
        pivot_unique = pd.merge(pivot_unique, total_unique, on='id_outlet', how='left')

        st.dataframe(pivot_unique, use_container_width=True)
        
        # Excel download for unique konsumen
        towrite1 = io.BytesIO()
        with pd.ExcelWriter(towrite1, engine='xlsxwriter') as writer:
            pivot_unique.to_excel(writer, index=False, sheet_name="Unique Konsumen")
        st.download_button(
            label="📥 Download Unique Konsumen",
            data=towrite1.getvalue(),
            file_name="unique_konsumen_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        

    # --- Tab 2: DSO Summary Table ---
    with tab2:
        st.subheader("📊 Active % by DSO / Site")
        df_dso_filtered = df_filtered.copy()
        df_dso_filtered['scanned'] = df_dso_filtered['tanggal_scan'].notna()

        # Apply DSO + Program filters
        summary = df_dso_filtered.groupby('dso').agg(
            total_outlets=('id_outlet', 'nunique'),
            active_outlets=('id_outlet', lambda x: x[df_dso_filtered['scanned']].nunique())
        ).reset_index()
        summary['% active'] = (summary['active_outlets'] / summary['total_outlets'] * 100).round(1)
        st.dataframe(summary, use_container_width=True)

    # --- Tab 3: List of Inactive Outlets ---
    with tab3:
        st.subheader("🚫 Outlets with Zero User Scans")
        scanned_ids = df_filtered[df_filtered['is_active']]['id_outlet'].unique()
        inactive_df = df_db[~df_db['id_outlet'].isin(scanned_ids)]
        inactive_df = inactive_df[['pic', 'id_outlet', 'nama_outlet', 'program']]
        st.dataframe(inactive_df.sort_values(by=['pic', 'id_outlet']), use_container_width=True)

    # --- Tab 4: Weekly Active % Table ---
    with tab4:
        st.subheader("📌 Weekly Active % by PIC and Program")
        total_outlets = df_filtered.drop_duplicates(subset=['id_outlet', 'pic', 'program']) \
            .groupby(['pic', 'program'])['id_outlet'].nunique().reset_index(name='total_outlets')
        df_weeks = df_filtered[df_filtered['week_number'].isin(selected_weeks)]
        active_counts = df_weeks[df_weeks['is_active']].drop_duplicates(subset=['id_outlet', 'week_number']) \
            .groupby(['pic', 'program', 'week_number'])['id_outlet'].count().reset_index(name='active_count')
        merged = pd.merge(active_counts, total_outlets, on=['pic', 'program'], how='left')
        merged['% active'] = (merged['active_count'] / merged['total_outlets'] * 100).round(1)
        pivot_df = merged.pivot_table(index=['pic', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()
        styled_df = pivot_df.style.format({col: "{:.1f}%" for col in pivot_df.columns[2:]}) \
            .applymap(highlight_low, subset=pd.IndexSlice[:, pivot_df.columns[2:]])
        st.dataframe(styled_df, use_container_width=True)

        # Export
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
            pivot_df.to_excel(writer, index=False, sheet_name="Weekly KPI")
        st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="weekly_kpi_report.xlsx")

    # --- Tab 5: Trend Charts ---
    with tab5:
        st.subheader("📈 Weekly Active % Trend by PIC")
        chart = alt.Chart(merged).mark_line(point=True).encode(
            x=alt.X('week_number:O', title='Week'),
            y=alt.Y('% active', title='% Active'),
            color='pic',
            tooltip=['pic', 'program', 'week_number', '% active']
        ).properties(width=800, height=400)
        st.altair_chart(chart, use_container_width=True)

        st.subheader("📊 Avg % Active by PIC")
        avg_active = merged.groupby('pic')['% active'].mean().reset_index().sort_values('% active', ascending=False)
        bar_chart = alt.Chart(avg_active).mark_bar().encode(
            x=alt.X('% active', title='Average % Active'),
            y=alt.Y('pic', sort='-x'),
            tooltip=['pic', '% active']
        ).properties(width=700, height=400)
        st.altair_chart(bar_chart, use_container_width=True)

    # --- Tab 6: Multi-Outlet Scans ---
    with tab6:
        st.subheader("📋 Phone Numbers Scanning in Multiple Outlets")
        df_valid = df_weeks[df_weeks['no_hp'].notna()]
        phone_outlets = df_valid.groupby('no_hp')['id_outlet'].nunique().reset_index(name='unique_outlets')
        multi = phone_outlets[phone_outlets['unique_outlets'] > 1].sort_values(by='unique_outlets', ascending=False)
        merged_multi = pd.merge(multi, df_weeks[['no_hp', 'id_outlet']].drop_duplicates(), on='no_hp', how='left')
        st.dataframe(merged_multi, use_container_width=True)

else:
    st.info("📂 Please upload a Scan File to begin.")
