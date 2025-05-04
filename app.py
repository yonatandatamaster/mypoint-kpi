
import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(page_title="MyPoint KPI Dashboard", layout="wide")
st.title("📊 MyPoint KPI Dashboard")

st.sidebar.header("Upload Excel Files")
scan_file = st.sidebar.file_uploader("Upload SCAN File", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload DATABASE File", type=["xlsx"])

if scan_file and db_file:
    df_scan = pd.read_excel(scan_file, engine='openpyxl')
    df_db = pd.read_excel(db_file, engine='openpyxl')

    scan_cols = {col.lower(): col for col in df_scan.columns}
    date_col = [col for col in scan_cols if "tanggal" in col][0]
    program_col = [col for col in scan_cols if "program" in col][0]
    idoutlet_col = [col for col in scan_cols if "outlet" in col and "id" in col][0]
    nowa_col = [col for col in scan_cols if "wa" in col][0]

    df_scan[date_col] = pd.to_datetime(df_scan[date_col])
    df_scan["Week Number"] = df_scan[date_col].dt.isocalendar().week
    df_scan["ID Outlet"] = df_scan[idoutlet_col].astype(str).str.strip()
    df_scan["Kode Program"] = df_scan[program_col].astype(str).str.strip()
    df_scan["No WA"] = df_scan[nowa_col].astype(str).str.strip()

    df_db["ID OUTLET"] = df_db["ID OUTLET"].astype(str).str.strip()
    df_db["PROGRAM"] = df_db["PROGRAM"].astype(str).str.strip()
    df_db["DSO"] = df_db["DSO "].astype(str).str.strip()
    df_db["PIC"] = df_db["PIC"].astype(str).str.strip()

    scan_merge = df_scan[["ID Outlet", "Kode Program", "Week Number", "No WA"]].drop_duplicates()
    scan_merge = pd.merge(scan_merge, df_db[["ID OUTLET", "DSO", "PROGRAM", "PIC"]],
                          left_on="ID Outlet", right_on="ID OUTLET", how="left")

    outlet_active = scan_merge.groupby(["DSO", "PROGRAM", "Week Number"])["ID Outlet"].nunique().reset_index()
    outlet_active = outlet_active.rename(columns={"ID Outlet": "Total Active Outlets"})

    outlet_total = df_db.groupby(["DSO", "PROGRAM"])["ID OUTLET"].nunique().reset_index()
    outlet_total = outlet_total.rename(columns={"ID OUTLET": "Total Assigned Outlets"})

    kpi = pd.merge(outlet_active, outlet_total, on=["DSO", "PROGRAM"], how="left")
    kpi["% Active Outlets"] = (kpi["Total Active Outlets"] / kpi["Total Assigned Outlets"]) * 100

    consumer = scan_merge.groupby(["DSO", "PROGRAM", "Week Number"])["No WA"].nunique().reset_index()
    consumer = consumer.rename(columns={"No WA": "Unique Consumers"})

    kpi_final = pd.merge(kpi, consumer, on=["DSO", "PROGRAM", "Week Number"], how="left")

    pic_summary = scan_merge.groupby(["PIC", "Week Number"])["ID Outlet"].nunique().reset_index()
    pic_total = df_db.groupby("PIC")["ID OUTLET"].nunique().reset_index().rename(
        columns={"ID OUTLET": "Total Assigned Outlets"})
    pic_summary = pd.merge(pic_summary, pic_total, left_on="PIC", right_on="PIC", how="left")
    pic_summary["% Active Outlets"] = (pic_summary["ID Outlet"] / pic_summary["Total Assigned Outlets"]) * 100

    week_selected = st.sidebar.selectbox("Select Week", sorted(kpi_final["Week Number"].unique()))
    filtered_kpi = kpi_final[kpi_final["Week Number"] == week_selected]
    filtered_pic = pic_summary[pic_summary["Week Number"] == week_selected]

    st.subheader(f"KPI Summary - Week {week_selected}")
    st.dataframe(filtered_kpi)

    def convert_df(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="KPI Summary", index=False)
        output.seek(0)
        return output

    st.download_button("📥 Download KPI Summary (Excel)", data=convert_df(filtered_kpi),
                       file_name="KPI_Summary.xlsx")

    st.subheader("📊 % Active Outlets by Program")
    prog_chart = filtered_kpi.groupby("PROGRAM")["% Active Outlets"].mean().reset_index()
    st.bar_chart(prog_chart.set_index("PROGRAM"))

    st.subheader("👥 % Active Outlets by PIC / Promotor")
    pic_chart = filtered_pic[["PIC", "% Active Outlets"]].set_index("PIC")
    st.bar_chart(pic_chart)
else:
    st.info("Upload both Scan and Database Excel files to begin.")
