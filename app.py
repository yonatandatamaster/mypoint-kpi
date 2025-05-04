import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# Upload files
st.sidebar.header("Upload Data Files")
scan_file = st.sidebar.file_uploader("Upload Scan File (Excel)", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File (Excel)", type=["xlsx"])

if scan_file and db_file:
    # Read Excel files
    df_scan = pd.read_excel(scan_file, sheet_name=0)
    df_db = pd.read_excel(db_file, sheet_name=0)

    # Clean & Prep Scan Data
    df_scan['Tanggal Scan'] = pd.to_datetime(df_scan['Tanggal Scan'])
    df_scan['Week Number'] = df_scan['Tanggal Scan'].dt.isocalendar().week
    df_scan['ID Outlet'] = df_scan['ID Outlet'].astype(str).str.strip()

    # Clean & Prep DB Data
    df_db['ID Outlet'] = df_db['ID Outlet'].astype(str).str.strip()

    # Merge
    merged = df_db.merge(df_scan, on='ID Outlet', how='left')
    merged['Scan Flag'] = ~merged['Tanggal Scan'].isna()

    # Summary: %Active per PIC / Promotor
    summary = merged.groupby('PIC / Promotor').agg(
        Total_Outlet=('ID Outlet', 'nunique'),
        Active_Outlet=('Scan Flag', 'sum')
    ).reset_index()
    summary['% Active'] = (summary['Active_Outlet'] / summary['Total_Outlet'] * 100).round(1).astype(str) + '%'

    st.subheader("📌 % Active Outlet by PIC / Promotor")
    st.dataframe(summary, use_container_width=True)

    # Visual
    st.subheader("📈 Bar Chart of % Active Outlets")
    fig, ax = plt.subplots(figsize=(10, 5))
    summary_sorted = summary.sort_values('Active_Outlet', ascending=False)
    sns.barplot(
        x='Active_Outlet', y='PIC / Promotor', data=summary_sorted,
        palette='viridis', ax=ax
    )
    ax.set_xlabel('Number of Active Outlets')
    ax.set_ylabel('PIC / Promotor')
    st.pyplot(fig)

else:
    st.info("Please upload both Scan and Database Excel files to begin.")
