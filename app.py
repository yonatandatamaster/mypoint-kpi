import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

st.set_page_config(page_title="MyPoint Outlet KPI Dashboard", layout="wide")

DB_FILE = "Master_Database_Outlet.xlsx"
WEEK_TAG_FILE = "Date_week_tag.xlsx"

@st.cache_data
def load_database():
    df = pd.read_excel(DB_FILE)
    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={
        'id outlet': 'id_outlet',
        'pic / promotor': 'pic',
        'program': 'program',
        'nama outlet': 'nama_outlet',
        'dso': 'dso'
    })
    df['id_outlet'] = df['id_outlet'].astype(str).str.strip()
    return df

@st.cache_data
def load_week_tags():
    week_df = pd.read_excel(WEEK_TAG_FILE)
    week_df.columns = week_df.columns.str.strip().str.lower()
    week_df = week_df.rename(columns={'tanggal': 'tanggal_scan', 'minggu': 'week_number'})
    week_df['tanggal_scan'] = pd.to_datetime(week_df['tanggal_scan'])
    return week_df

@st.cache_data
def prepare_scan_data(scan_file, week_map):
    df = pd.read_excel(scan_file, sheet_name=None)
    df = df[list(df.keys())[0]]
    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={
        'tanggal scan': 'tanggal_scan',
        'id outlet': 'id_outlet',
        'no wa': 'no_hp'
    })
    df['tanggal_scan'] = pd.to_datetime(df['tanggal_scan'], errors='coerce')
    df['id_outlet'] = df['id_outlet'].astype(str).str.strip()
    df['no_hp'] = df['no_hp'].astype(str).str.strip()
    df = pd.merge(df, week_map, on='tanggal_scan', how='left')
    return df

def highlight_low(val):
    try:
        return 'background-color: #ffcccc' if float(val) < 50 else ''
    except:
        return ''

# --- UI Upload ---
st.sidebar.header("📁 Upload Scan File")
scan_file = st.sidebar.file_uploader("Upload only the Scan File (.xlsx)", type=["xlsx"])

if scan_file:
    df_db = load_database()
    week_map = load_week_tags()
    df_scan = prepare_scan_data(scan_file, week_map)

    df_scan = pd.merge(df_scan, df_db[['id_outlet', 'pic', 'program', 'dso']], on='id_outlet', how='left')

    # --- Filters ---
    st.sidebar.header("🔎 Filter Data")
    dso_options = sorted(df_db['dso'].dropna().unique())
    selected_dso = st.sidebar.selectbox("Filter by DSO", options=dso_options)
    selected_programs = st.sidebar.multiselect("Filter by Program", sorted(df_db['program'].unique()), default=sorted(df_db['program'].unique()))
    selected_weeks = st.sidebar.multiselect("Select Weeks", sorted(week_map['week_number'].unique()), default=sorted(week_map['week_number'].unique()))
    selected_pics = st.sidebar.multiselect("Select PIC(s)", sorted(df_db['pic'].dropna().unique()), default=sorted(df_db['pic'].dropna().unique()))

    df_db = df_db[df_db['dso'] == selected_dso]
    df_db = df_db[df_db['program'].isin(selected_programs)]
    df_db = df_db[df_db['pic'].isin(selected_pics)]

    df_scan = df_scan[df_scan['dso'] == selected_dso]
    df_scan = df_scan[df_scan['program'].isin(selected_programs)]
    df_scan = df_scan[df_scan['week_number'].isin(selected_weeks)]
    df_scan = df_scan[df_scan['pic'].isin(selected_pics)]

    tabs = st.tabs([
        "📊 Dashboard Table Unique Konsumen",
        "📊 DSO Summary Table",
        "🚫 List of Inactive Outlets",
        "📊 Dashboard Table",
        "📈 Trends & Charts",
        "📋 Multi-Outlet Scans"
    ])

    # ---------------- TAB 1 ----------------
    with tabs[0]:
        st.subheader("📊 Weekly Unique Consumers per Outlet")
        weekly_unique = df_scan.groupby(['pic', 'program', 'id_outlet', 'week_number'])['no_hp'].nunique().reset_index(name='unique_users')
        pivot_uniq = weekly_unique.pivot_table(index=['pic', 'program', 'id_outlet'], columns='week_number', values='unique_users', fill_value=0).reset_index()

        total_unique = df_scan.groupby('id_outlet')['no_hp'].nunique().reset_index(name='total_unique_konsumen')
        pivot_uniq = pd.merge(pivot_uniq, total_unique, on='id_outlet', how='left')

        st.dataframe(pivot_uniq, use_container_width=True)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            pivot_uniq.to_excel(writer, index=False, sheet_name="Unique Consumers")
        st.download_button("📥 Download Excel Report", data=out.getvalue(), file_name="unique_consumers_report.xlsx")

    # ---------------- TAB 2 ----------------
    with tabs[1]:
        st.subheader("📊 Active % by DSO / Site")

        base = df_db[['dso', 'id_outlet']].drop_duplicates().groupby('dso').count().reset_index(name='total')
        active = df_scan[['dso', 'id_outlet', 'week_number']].drop_duplicates().groupby(['dso', 'week_number']).count().reset_index(name='active')
        merged = pd.merge(active, base, on='dso', how='left')
        merged['% active'] = (merged['active'] / merged['total'] * 100).round(1)

        pivot_dso = merged.pivot_table(index='dso', columns='week_number', values='% active', fill_value=0).reset_index()
        styled_dso = pivot_dso.style.format({col: "{:.1f}%" for col in pivot_dso.columns[2:]}).applymap(highlight_low, subset=pd.IndexSlice[:, pivot_dso.columns[1:]])
        st.dataframe(styled_dso, use_container_width=True)

    # ---------------- TAB 3 ----------------
    with tabs[2]:
        st.subheader("🚫 Outlets with Zero User Scans (per Program & PIC & Week)")

        # Create all combinations
        all_combinations = (
            df_db[['id_outlet', 'nama_outlet', 'pic', 'program']]
            .assign(key=1)
            .merge(pd.DataFrame({'week_number': selected_weeks, 'key': 1}), on='key')
            .drop('key', axis=1)
        )

        scanned_outlets = df_scan[['id_outlet', 'week_number']].drop_duplicates()
        inactive_df = pd.merge(all_combinations, scanned_outlets, on=['id_outlet', 'week_number'], how='left', indicator=True)
        inactive_df = inactive_df[inactive_df['_merge'] == 'left_only'].drop(columns='_merge')

        inactive_df = inactive_df[inactive_df['pic'].isin(selected_pics)]
        inactive_df = inactive_df[inactive_df['program'].isin(selected_programs)]

        display_cols = ['pic', 'id_outlet', 'nama_outlet', 'program', 'week_number']
        st.dataframe(inactive_df[display_cols], use_container_width=True)

        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
            inactive_df[display_cols].to_excel(writer, index=False, sheet_name="Inactive Outlets")
        st.download_button("📥 Download Inactive Outlets", data=towrite.getvalue(), file_name="inactive_outlets.xlsx")

    # ---------------- TAB 4 ----------------
    with tabs[3]:
        st.subheader("📌 Weekly Active % by PIC and Program")

        total_outlets = df_db.drop_duplicates(subset=['id_outlet', 'pic', 'program']) \
            .groupby(['pic', 'program'])['id_outlet'].count().reset_index(name='total_outlets')

        active_weekly = df_scan[df_scan['no_hp'].notna()].drop_duplicates(subset=['id_outlet', 'week_number']) \
            .groupby(['pic', 'program', 'week_number'])['id_outlet'].count().reset_index(name='active_count')

        merged = pd.merge(active_weekly, total_outlets, on=['pic', 'program'], how='left')
        merged['% active'] = (merged['active_count'] / merged['total_outlets'] * 100).round(1)

        pivot_df = merged.pivot_table(index=['pic', 'program'], columns='week_number', values='% active', fill_value=0).reset_index()

        styled_df = pivot_df.style.format({col: "{:.1f}%" for col in pivot_df.columns[2:]}) \
            .applymap(highlight_low, subset=pd.IndexSlice[:, pivot_df.columns[2:]])
        st.dataframe(styled_df, use_container_width=True)

    # ---------------- TAB 5 ----------------
    with tabs[4]:
        st.subheader("📈 Weekly Active % Trend by PIC")
        chart = alt.Chart(merged).mark_line(point=True).encode(
            x=alt.X('week_number:O', title='Week'),
            y=alt.Y('% active', title='% Active'),
            color='pic',
            tooltip=['pic', 'program', 'week_number', '% active']
        ).properties(width=800, height=400)
        st.altair_chart(chart, use_container_width=True)

    # ---------------- TAB 6 ----------------
    with tabs[5]:
        st.subheader("📋 Phone Numbers Scanning in Multiple Outlets")
        df_valid = df_scan[df_scan['no_hp'].notna()]
        phone_outlets = df_valid.groupby('no_hp')['id_outlet'].nunique().reset_index(name='unique_outlets')
        multi = phone_outlets[phone_outlets['unique_outlets'] > 1].sort_values(by='unique_outlets', ascending=False)
        merged_multi = pd.merge(multi, df_valid[['no_hp', 'id_outlet']].drop_duplicates(), on='no_hp', how='left')
        st.dataframe(merged_multi, use_container_width=True)

else:
    st.info("📂 Please upload the Scan File (.xlsx) to begin.")
