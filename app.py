import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")

# File constants
DB_FILE = "Master_Database_Outlet.xlsx"
WEEK_TAG_FILE = "Date_week_tag.xlsx"

@st.cache_data
def load_database():
    df_db = pd.read_excel(DB_FILE, sheet_name=None)
    df_db = df_db[list(df_db.keys())[0]]
    df_db.columns = df_db.columns.str.strip().str.lower()
    df_db = df_db.rename(columns={
        'id outlet': 'id_outlet',
        'nama outlet (dsca)': 'nama_outlet',  # ← fixed here
        'pic / promotor': 'pic',
        'program': 'program',
        'dso': 'dso'
    })
    df_db['id_outlet'] = df_db['id_outlet'].astype(str).str.strip()
    return df_db


@st.cache_data
def load_week_tags():
    week_df = pd.read_excel(WEEK_TAG_FILE)
    week_df.columns = week_df.columns.str.strip().str.lower()
    week_df = week_df.rename(columns={'date': 'tanggal_scan', 'week': 'week_number'})
    week_df['tanggal_scan'] = pd.to_datetime(week_df['tanggal_scan'], errors='coerce')
    return week_df

def highlight_low(val):
    try:
        return 'background-color: #ffcccc' if float(val) < 50 else ''
    except:
        return ''

# Upload Scan File
st.sidebar.header("📁 Upload Scan File")
scan_file = st.sidebar.file_uploader("Upload only the Scan File (.xlsx)", type=["xlsx"])

if scan_file:
    df_db = load_database()
    week_map = load_week_tags()

    df_scan = pd.read_excel(scan_file, sheet_name=None)
    df_scan = df_scan[list(df_scan.keys())[0]]
    df_scan.columns = df_scan.columns.str.strip().str.lower()
    df_scan = df_scan.rename(columns={
        'tanggal scan': 'tanggal_scan',
        'id outlet': 'id_outlet',
        'no wa': 'no_hp'
    })
    df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
    df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
    df_scan['no_hp'] = df_scan['no_hp'].astype(str).str.strip()
    df_scan = pd.merge(df_scan, week_map, on='tanggal_scan', how='left')

    df = pd.merge(df_db, df_scan, on='id_outlet', how='left')
    df['is_active'] = df['tanggal_scan'].notna()

    # Sidebar filters
    st.sidebar.header("🔍 Filter Data")
    selected_dso = st.sidebar.selectbox("Filter by DSO", sorted(df['dso'].dropna().unique()))
    df = df[df['dso'] == selected_dso]

    all_programs = sorted(df['program'].dropna().unique())
    selected_programs = st.sidebar.multiselect("Filter by Program", options=all_programs, default=all_programs)
    df = df[df['program'].isin(selected_programs)]

    all_weeks = sorted(df['week_number'].dropna().unique())
    selected_weeks = st.sidebar.multiselect("Select Weeks", options=all_weeks, default=all_weeks)
    df = df[df['week_number'].isin(selected_weeks)]

    all_pics = sorted(df['pic'].dropna().unique())
    selected_pics = st.sidebar.multiselect("Select PIC(s)", options=all_pics, default=all_pics)
    df = df[df['pic'].isin(selected_pics)]

    # Tabs
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Dashboard Table Unique Konsumen",
        "📊 DSO Summary Table",
        "🚫 List of Inactive Outlets",
        "📊 Dashboard Table",
        "📈 Trends & Charts",
        "📋 Multi-Outlet Scans"
    ])

    # Tab 1: Unique Konsumen
    with tab1:
        st.subheader("📊 Weekly Unique Consumers per Outlet")
        weekly_unique = df.dropna(subset=['no_hp']).groupby(['pic', 'program', 'id_outlet', 'week_number'])['no_hp'].nunique().reset_index(name='unique_konsumen')
        pivot = weekly_unique.pivot_table(index=['pic', 'program', 'id_outlet'], columns='week_number', values='unique_konsumen', fill_value=0).reset_index()
        total_unique = df.dropna(subset=['no_hp']).groupby(['id_outlet'])['no_hp'].nunique().reset_index(name='total_unique_konsumen')
        pivot = pd.merge(pivot, total_unique, on='id_outlet', how='left')
        st.dataframe(pivot, use_container_width=True)

        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
            pivot.to_excel(writer, index=False, sheet_name="Weekly Unique Konsumen")
        st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="unique_konsumen.xlsx")

    # Tab 2: DSO Summary
    with tab2:
        st.subheader("📊 Active % by DSO / Site")
        total_outlets = df_db[df_db['dso'] == selected_dso].groupby(['dso', 'program'])['id_outlet'].nunique().reset_index(name='total_outlets')
        active_weekly = df[df['is_active']].drop_duplicates(['id_outlet', 'week_number']).groupby(['dso', 'program', 'week_number'])['id_outlet'].nunique().reset_index(name='active_outlets')
        merged = pd.merge(active_weekly, total_outlets, on=['dso', 'program'], how='left')
        merged['% active'] = (merged['active_outlets'] / merged['total_outlets'] * 100).round(1)
        pivot_dso = merged.pivot_table(index=['dso', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()
        styled_dso = pivot_dso.style.format({col: "{:.1f}%" for col in pivot_dso.columns[2:]}).applymap(highlight_low, subset=pd.IndexSlice[:, pivot_dso.columns[2:]])
        st.dataframe(styled_dso, use_container_width=True)

    # Tab 3: Inactive Outlets (Filtered + Download)
    with tab3:
        st.subheader("🚫 Outlets with Zero User Scans (per Program & PIC)")
        active_ids = df[df['is_active']]['id_outlet'].unique()
        inactive_df = df_db[~df_db['id_outlet'].isin(active_ids)]
        inactive_df = inactive_df[inactive_df['dso'] == selected_dso]
        inactive_df = inactive_df[inactive_df['program'].isin(selected_programs)]
        inactive_df = inactive_df[inactive_df['pic'].isin(selected_pics)]
        display_cols = ['pic', 'id_outlet', 'nama_outlet', 'program']
        st.dataframe(inactive_df[display_cols], use_container_width=True)

        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
            inactive_df[display_cols].to_excel(writer, index=False, sheet_name="Inactive Outlets")
        st.download_button("📥 Download Inactive Outlets", data=towrite.getvalue(), file_name="inactive_outlets.xlsx")

    # Tab 4: Weekly Active %
    with tab4:
        st.subheader("📌 Weekly Active % by PIC and Program")
        total_outlets = df_db[df_db['dso'] == selected_dso].groupby(['pic', 'program'])['id_outlet'].nunique().reset_index(name='total_outlets')
        active_weekly = df[df['is_active']].drop_duplicates(['id_outlet', 'week_number']).groupby(['pic', 'program', 'week_number'])['id_outlet'].nunique().reset_index(name='active_count')
        merged = pd.merge(active_weekly, total_outlets, on=['pic', 'program'], how='left')
        merged['% active'] = (merged['active_count'] / merged['total_outlets'] * 100).round(1)
        pivot_df = merged.pivot_table(index=['pic', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()
        styled_df = pivot_df.style.format({col: "{:.1f}%" for col in pivot_df.columns[2:]}).applymap(highlight_low, subset=pd.IndexSlice[:, pivot_df.columns[2:]])
        st.dataframe(styled_df, use_container_width=True)

    # Tab 5: Trend Charts
    with tab5:
        st.subheader("📈 Weekly Active % Trend by PIC")
        chart = alt.Chart(merged).mark_line(point=True).encode(
            x=alt.X('week_number:O'),
            y=alt.Y('% active'),
            color='pic',
            tooltip=['pic', 'program', 'week_number', '% active']
        ).properties(width=800, height=400)
        st.altair_chart(chart, use_container_width=True)

    # Tab 6: Multi-Outlet Scans
    with tab6:
        st.subheader("📋 Phone Numbers Scanning in Multiple Outlets")
        df_valid = df[df['no_hp'].notna()]
        phone_outlets = df_valid.groupby('no_hp')['id_outlet'].nunique().reset_index(name='unique_outlets')
        multi = phone_outlets[phone_outlets['unique_outlets'] > 1].sort_values(by='unique_outlets', ascending=False)
        merged_multi = pd.merge(multi, df[['no_hp', 'id_outlet']].drop_duplicates(), on='no_hp', how='left')
        st.dataframe(merged_multi, use_container_width=True)

else:
    st.info("📂 Please upload the Scan File (.xlsx) to begin.")
