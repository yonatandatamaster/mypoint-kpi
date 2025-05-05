import streamlit as st
import pandas as pd

st.set_page_config(page_title="KPI Report Table", layout="wide")
st.title("📊 Outlet Report Table - Mingguan")

scan_file = st.sidebar.file_uploader("Upload Scan File", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File", type=["xlsx"])

if scan_file and db_file:
    try:
        df_scan = pd.read_excel(scan_file)
        df_db = pd.read_excel(db_file)

        # Normalisasi kolom
        df_scan.columns = df_scan.columns.str.strip().str.lower()
        df_db.columns = df_db.columns.str.strip().str.lower()

        df_scan['tanggal_scan'] = pd.to_datetime(df_scan['tanggal scan'], errors='coerce')
        df_scan['minggu'] = df_scan['tanggal_scan'].dt.isocalendar().week
        df_scan['id outlet'] = df_scan['id outlet'].astype(str).str.strip()
        df_db['id outlet'] = df_db['id outlet'].astype(str).str.strip()

        df_merged = pd.merge(df_db, df_scan[['id outlet', 'minggu']], on='id outlet', how='left')

        # Hitung jumlah outlet aktif per minggu
        summary = (
            df_merged
            .groupby(['dso', 'pic / promotor', 'program', 'nama program', 'minggu'])
            .agg(jumlah_outlet=('id outlet', 'nunique'))
            .reset_index()
        )

        pivot = summary.pivot_table(
            index=['dso', 'pic / promotor', 'program', 'nama program'],
            columns='minggu',
            values='jumlah_outlet',
            fill_value=0
        ).reset_index()

        st.dataframe(pivot, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal memproses data: {e}")
else:
    st.info("Silakan upload kedua file Excel terlebih dahulu.")
