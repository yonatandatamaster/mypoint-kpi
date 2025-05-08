import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")

# --- Constants ---
DB_FILE = "Master_Database_Outlet.xlsx"
WEEK_TAG_FILE = "Date_week_tag.xlsx"

@st.cache_data
def load_database():
    df_db = pd.read_excel(DB_FILE)
    df_db.columns = df_db.columns.str.strip().str.lower()
    df_db = df_db.rename(columns={
        'id outlet': 'id_outlet',
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
    week_df = week_df.rename(columns={'tanggal': 'tanggal_scan', 'minggu': 'week_number'})
    week_df['tanggal_scan'] = pd.to_datetime(week_df['tanggal_scan'], errors='coerce')
    return week_df

def highlight_low(val):
    try:
        return 'background-color: #ffcccc' if float(val) < 50 else ''
    except:
        return ''

# --- Sidebar Upload ---
st.sidebar.header("Upload Scan File")
scan_file = st.sidebar.file_uploader("Upload only the Scan File (.xlsx)", type=["xlsx"])

if scan_file:
    df_db = load_database()
    week_map = load_week_tags()

    # --- Load Scan File ---
    df_scan = pd.read_excel(scan_file, sheet_name=None)
    df_scan = df_scan[list(df_scan.keys())[0]]  # use first sheet
    df_scan.columns = df_scan.columns.str.strip().str.lower()
    df_scan = df_scan.rename(columns={
        'tanggal scan': 'tanggal_scan',
        'id outlet': 'id_outlet',
        'kode program': 'kode_program',
        'no wa': 'no_hp'
    })

    df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal_scan'], errors='coerce')
    df_scan['id_outlet'] = df_scan['id_outlet'].astype(str).str.strip()
    df_scan['no_hp'] = df_scan['no_hp'].astype(str).str.strip()

    # --- Merge with Week Mapping ---
    df_scan = pd.merge(df_scan, week_map, on='tanggal_scan', how='left')
    df_scan = df_scan.dropna(subset=['week_number'])

    # --- Merge with Master DB ---
    df_merged = pd.merge(df_db, df_scan, on='id_outlet', how='left')
    df_merged['is_active'] = df_merged['tanggal_scan'].notna()

    # --- Filters ---
    st.sidebar.header("🔍 Additional Filters")
    selected_dso = st.sidebar.selectbox("Filter by DSO", options=sorted(df_merged['dso'].dropna().unique()))
    df_filtered = df_merged[df_merged['dso'] == selected_dso]

    selected_programs = st.sidebar.multiselect("Filter by Program", options=sorted(df_filtered['program'].dropna().unique()), default=sorted(df_filtered['program'].dropna().unique()))
    df_filtered = df_filtered[df_filtered['program'].isin(selected_programs)]

    all_weeks = sorted(df_filtered['week_number'].dropna().unique())
    selected_weeks = st.sidebar.multiselect("Select Weeks", options=all_weeks, default=all_weeks)
    df_filtered = df_filtered[df_filtered['week_number'].isin(selected_weeks)]

    selected_pics = st.sidebar.multiselect("Select PIC(s)", options=sorted(df_filtered['pic'].dropna().unique()), default=sorted(df_filtered['pic'].dropna().unique()))
    df_filtered = df_filtered[df_filtered['pic'].isin(selected_pics)]

    # --- Tabs ---
    tab1, tab2, tab3 = st.tabs(["📊 Dashboard Table", "📈 Trends & Charts", "📋 Multi-Outlet Scans"])

    with tab1:
        st.subheader("📌 Weekly Active % by PIC and Program")

        total_outlets = df_filtered.drop_duplicates(subset=['id_outlet', 'pic', 'program']) \
            .groupby(['pic', 'program'])['id_outlet'].count().reset_index(name='total_outlets')

        active_weekly = df_filtered[df_filtered['is_active']].drop_duplicates(subset=['id_outlet', 'week_number']) \
            .groupby(['pic', 'program', 'week_number'])['id_outlet'].count().reset_index(name='active_count')

        merged = pd.merge(active_weekly, total_outlets, on=['pic', 'program'], how='left')
        merged['% active'] = (merged['active_count'] / merged['total_outlets'] * 100).round(1)

        pivot_df = merged.pivot_table(index=['pic', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()

        styled_df = pivot_df.style.format({col: "{:.1f}%" for col in pivot_df.columns[2:]}) \
            .applymap(highlight_low, subset=pd.IndexSlice[:, pivot_df.columns[2:]])

        st.dataframe(styled_df, use_container_width=True)

        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
            pivot_df.to_excel(writer, index=False, sheet_name="Weekly KPI")
        st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="weekly_kpi_report.xlsx")

    with tab2:
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

    with tab3:
        st.subheader("📋 Phone Numbers Scanning in Multiple Outlets")
        df_valid = df_filtered[df_filtered['no_hp'].notna()]
        phone_outlets = df_valid.groupby('no_hp')['id_outlet'].nunique().reset_index(name='unique_outlets')
        multi = phone_outlets[phone_outlets['unique_outlets'] > 1].sort_values(by='unique_outlets', ascending=False)
        merged_multi = pd.merge(multi, df_filtered[['no_hp', 'id_outlet']].drop_duplicates(), on='no_hp', how='left')
        st.dataframe(merged_multi, use_container_width=True)

else:
    st.info("📂 Please upload the Scan File (.xlsx) to begin.")
