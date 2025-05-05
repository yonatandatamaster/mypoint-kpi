import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import io

st.set_page_config(page_title="Outlet KPI Dashboard", layout="wide")
st.title("📊 MyPoint Outlet KPI Dashboard")

# Upload section
st.sidebar.header("Upload Data Files")
scan_file = st.sidebar.file_uploader("Upload Scan File (Excel)", type=["xlsx"])
db_file = st.sidebar.file_uploader("Upload Database File (Excel)", type=["xlsx"])

if scan_file and db_file:
    try:
        df_scan = pd.read_excel(scan_file)
        df_db = pd.read_excel(db_file)

        df_scan.columns = df_scan.columns.str.strip().str.lower()
        df_db.columns = df_db.columns.str.strip().str.lower()

        df_scan = df_scan.rename(columns={"tanggal scan": "tanggal_scan", "id outlet": "id_outlet", "no hp": "no_hp"})
        df_db = df_db.rename(columns={"id outlet": "id_outlet", "pic / promotor": "pic"})

        df_scan["tanggal_scan"] = pd.to_datetime(df_scan["tanggal_scan"], errors="coerce")
        df_scan["id_outlet"] = df_scan["id_outlet"].astype(str).str.strip()
        df_db["id_outlet"] = df_db["id_outlet"].astype(str).str.strip()
        df_scan["no_hp"] = df_scan["no_hp"].astype(str).str.strip()

        # WEEKNUM(..., 16) compatibility: week starts on Saturday
        df_scan["week_start"] = df_scan["tanggal_scan"] - pd.to_timedelta((df_scan["tanggal_scan"].dt.weekday + 2) % 7, unit="d")
        df_scan["week_number"] = df_scan["week_start"].dt.isocalendar().week

        # Merge
        df_merged = df_db.merge(df_scan[["id_outlet", "week_number"]], on="id_outlet", how="left")
        df_merged["is_active"] = df_merged["week_number"].notna()

        # Sidebar Filters
        st.sidebar.header("🔍 Additional Filters")
        selected_dso = st.sidebar.selectbox("Filter by DSO", options=sorted(df_db["dso"].dropna().unique()))
        filtered_db = df_db[df_db["dso"] == selected_dso]

        selected_programs = st.sidebar.multiselect("Filter by Program", options=sorted(filtered_db["program"].unique()), default=sorted(filtered_db["program"].unique()))
        filtered_db = filtered_db[filtered_db["program"].isin(selected_programs)]

        selected_weeks = st.sidebar.multiselect("Select Weeks", options=sorted(df_scan["week_number"].dropna().unique()), default=sorted(df_scan["week_number"].dropna().unique()))
        df_scan_filtered = df_scan[df_scan["week_number"].isin(selected_weeks)]

        # --- Core Pivot Calculation ---
        df_combined = filtered_db.merge(df_scan_filtered[["id_outlet", "week_number"]], on="id_outlet", how="left")
        df_combined["active"] = df_combined["week_number"].notna()

        summary = (
            df_combined.groupby(["pic", "program", "week_number"])
            .agg(total_outlets=("id_outlet", "count"), active_outlets=("active", "sum"))
            .reset_index()
        )
        summary["% active"] = (summary["active_outlets"] / summary["total_outlets"] * 100).round(1)
        pivot_df = summary.pivot(index=["pic", "program"], columns="week_number", values="% active").fillna(0).reset_index()

        def highlight_low(val):
            try:
                return 'background-color: #ffcccc' if float(val) < 50 else ''
            except:
                return ''

        # --- Tabs ---
        tab1, tab2, tab3, tab4 = st.tabs(["📋 Dashboard Table", "📈 Trends & Charts", "📄 Raw Data", "📞 Multi-Outlet Scans"])

        with tab1:
            st.subheader("📌 Weekly Active % by PIC and Program")
            week_cols = pivot_df.columns[2:]
            styled_df = pivot_df.style.format({col: "{:.1f}%" for col in week_cols}).applymap(highlight_low, subset=pd.IndexSlice[:, week_cols])
            st.dataframe(styled_df, use_container_width=True)

            # Excel download
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                pivot_df.to_excel(writer, index=False, sheet_name="Weekly KPI")
            st.download_button("📥 Download Excel Report", data=towrite.getvalue(), file_name="weekly_kpi_report.xlsx")

        with tab2:
            st.subheader("📈 Weekly Active % Trend by PIC")
            st.altair_chart(
                alt.Chart(summary).mark_line(point=True).encode(
                    x=alt.X("week_number:O", title="Week"),
                    y=alt.Y("% active", title="% Active"),
                    color="pic",
                    tooltip=["pic", "program", "week_number", "% active"]
                ).properties(width=800, height=400), use_container_width=True
            )

        with tab3:
            st.subheader("📄 Raw Database + Scan Preview")
            st.write("Filtered Master Database", filtered_db)
            st.write("Filtered Scan Data", df_scan_filtered)

        with tab4:
            st.subheader("📞 Phone Numbers Scanning in Multiple Outlets")
            multi = df_scan.groupby("no_hp")["id_outlet"].nunique().reset_index()
            multi = multi[multi["id_outlet"] > 1].sort_values("id_outlet", ascending=False).rename(columns={"id_outlet": "unique_outlets"})
            st.dataframe(multi, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
else:
    st.info("📤 Please upload both Scan and Database Excel files to begin.")
