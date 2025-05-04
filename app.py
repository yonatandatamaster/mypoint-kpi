
import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# --- File upload section ---
st.sidebar.header("Upload Data Files")
scan_file = st.sidebar.file_uploader("Upload Scan File (Excel)", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File (Excel)", type=["xlsx"])

if scan_file and db_file:
    # Load sheets
    df_scan = pd.read_excel(scan_file, sheet_name='scan')
    df_db = pd.read_excel(db_file, sheet_name='database')

    # Clean and prep
    df_scan['Tanggal Scan'] = pd.to_datetime(df_scan['Tanggal Scan'])
    df_scan['Week Number'] = df_scan['Tanggal Scan'].dt.isocalendar().week
    df_scan['ID Outlet'] = df_scan['ID Outlet'].astype(str).str.strip()
    df_scan['Kode Program'] = df_scan['Kode Program'].astype(str).str.strip()

    df_db['ID OUTLET'] = df_db['ID OUTLET'].astype(str).str.strip()
    df_db['PROGRAM'] = df_db['PROGRAM'].astype(str).str.strip()
    df_db['DSO'] = df_db['DSO '].astype(str).str.strip()

    # --- KPI Summary ---
    scan_weekly = df_scan[['ID Outlet', 'Kode Program', 'Week Number']].drop_duplicates()
    scan_weekly['Scanned'] = 1

    scan_kpi = pd.merge(scan_weekly, df_db[['ID OUTLET', 'DSO', 'PROGRAM']],
                        left_on='ID Outlet', right_on='ID OUTLET', how='left')

    kpi_summary = scan_kpi.groupby(['DSO', 'PROGRAM', 'Week Number'])['ID Outlet'].nunique().reset_index()
    kpi_summary.rename(columns={'ID Outlet': 'Total Active Outlets'}, inplace=True)

    assigned_outlets = df_db.groupby(['DSO', 'PROGRAM'])['ID OUTLET'].nunique().reset_index()
    assigned_outlets.rename(columns={'ID OUTLET': 'Total Assigned Outlets'}, inplace=True)

    kpi_full = pd.merge(kpi_summary, assigned_outlets, on=['DSO', 'PROGRAM'], how='left')
    kpi_full['% Active Outlets'] = (kpi_full['Total Active Outlets'] / kpi_full['Total Assigned Outlets']) * 100

    # --- Consumer Count ---
    consumer_counts = df_scan[['No WA', 'ID Outlet', 'Kode Program', 'Week Number']].drop_duplicates()
    consumer_counts = pd.merge(consumer_counts, df_db[['ID OUTLET', 'DSO', 'PROGRAM']],
                               left_on='ID Outlet', right_on='ID OUTLET', how='left')
    consumer_summary = consumer_counts.groupby(['DSO', 'PROGRAM', 'Week Number'])['No WA'].nunique().reset_index()
    consumer_summary.rename(columns={'No WA': 'Unique Consumers'}, inplace=True)

    # --- Final KPI Table ---
    final_kpi = pd.merge(kpi_full, consumer_summary, on=['DSO', 'PROGRAM', 'Week Number'], how='left')

    # --- Sidebar Filters ---
    selected_week = st.sidebar.selectbox("Select Week", sorted(final_kpi['Week Number'].unique()))
    filtered_kpi = final_kpi[final_kpi['Week Number'] == selected_week]

    st.subheader(f"KPI Metrics - Week {selected_week}")
    st.dataframe(filtered_kpi)

    # --- Visualization ---
    st.subheader("📈 % Active Outlets by Program")
    fig1, ax1 = plt.subplots()
    chart_data = filtered_kpi.groupby('PROGRAM')['% Active Outlets'].mean().reset_index()
    sns.barplot(data=chart_data, x='% Active Outlets', y='PROGRAM', ax=ax1)
    st.pyplot(fig1)

    st.subheader("🏪 Unique Consumers by Program")
    fig2, ax2 = plt.subplots()
    chart_data2 = filtered_kpi.groupby('PROGRAM')['Unique Consumers'].sum().reset_index()
    sns.barplot(data=chart_data2, x='Unique Consumers', y='PROGRAM', ax=ax2, palette='Blues_d')
    st.pyplot(fig2)

else:
    st.info("Please upload both Scan and Database Excel files to begin.")
