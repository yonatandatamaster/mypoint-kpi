import streamlit as st
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import altair as alt

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# --- File upload section ---
st.sidebar.header("Upload Data Files")
scan_file = st.sidebar.file_uploader("Upload Scan File (Excel)", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File (Excel)", type=["xlsx"])

if scan_file and db_file:
    # Load sheets
    df_scan = pd.read_excel(scan_file, sheet_name=None)
    df_db = pd.read_excel(db_file, sheet_name=None)

    # Extract expected sheets
    df_scan = df_scan[list(df_scan.keys())[0]]
    df_db = df_db[list(df_db.keys())[0]]

    # --- Data Cleaning ---
    df_scan['Tanggal Scan'] = pd.to_datetime(df_scan['Tanggal Scan'])
    df_scan['Week Number'] = df_scan['Tanggal Scan'].dt.isocalendar().week
    df_scan['ID Outlet'] = df_scan['ID Outlet'].astype(str).str.strip()
    df_scan['Kode Program'] = df_scan['Kode Program'].astype(str).str.strip()
    df_db['ID Outlet'] = df_db['ID Outlet'].astype(str).str.strip()

    # --- Main KPI Table ---
    df_merged = pd.merge(df_db, df_scan, on='ID Outlet', how='left')
    df_merged['Is Active'] = df_merged['Tanggal Scan'].notna()

    active_summary = df_merged.groupby(['PIC', 'ID Outlet']).agg({'Is Active': 'max'}).reset_index()
    percent_active = active_summary.groupby('PIC')['Is Active'].mean().reset_index()
    percent_active['% Active Outlets'] = percent_active['Is Active'] * 100
    percent_active = percent_active.drop(columns=['Is Active'])

    # --- Show result ---
    st.subheader("% Active Outlets per PIC")
    st.dataframe(percent_active)

    chart = alt.Chart(percent_active).mark_bar().encode(
        x=alt.X('% Active Outlets', title='% Active Outlets'),
        y=alt.Y('PIC', sort='-x'),
        tooltip=['PIC', '% Active Outlets']
    ).properties(
        width=700,
        height=400
    )

    st.altair_chart(chart, use_container_width=True)
else:
    st.info("Please upload both Scan and Database Excel files to begin.")
