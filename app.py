import pandas as pd
import streamlit as st
import altair as alt
from datetime import date, timedelta

st.set_page_config(page_title="Primavera Planning Dashboard", layout="wide")
st.title("Primavera Planning Dashboard (Instant)")
st.caption("Planning dashboard. Data source = extracted CSV committed by GitHub Actions.")

# ---- Load CSV safely ----
try:
    df = pd.read_csv("data/primavera.csv")
except Exception as e:
    st.error("Could not read data/primavera.csv. Make sure GitHub Actions created it and committed it.")
    st.stop()

# Ensure expected columns exist
required_cols = ["major_group","package_code","work_type","activity_id","activity_name","start","finish"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"CSV is missing required columns: {missing}. Re-run the GitHub Actions extractor.")
    st.stop()

# Parse dates robustly
df["start"] = pd.to_datetime(df["start"], errors="coerce").dt.date
df["finish"] = pd.to_datetime(df["finish"], errors="coerce").dt.date

# Drop rows with invalid dates
df = df.dropna(subset=["start", "finish"]).copy()

if df.empty:
    st.warning("No valid rows found in CSV (start/finish dates are missing). Re-run extractor and check PDF text extraction.")
    st.stop()

# Enforce start <= finish
df = df[df["start"] <= df["finish"]].copy()
if df.empty:
    st.warning("All rows have start > finish after cleaning. Check the PDF extraction output.")
    st.stop()

# Sidebar filters
st.sidebar.header("Filters")

major_options = ["(All)"] + sorted(df["major_group"].dropna().astype(str).unique().tolist())
major = st.sidebar.selectbox("Major Group", major_options, index=0)

pkg_options = ["(All)"] + sorted(df["package_code"].dropna().astype(str).unique().tolist())
pkg = st.sidebar.selectbox("Area / Package", pkg_options, index=0)

wt_options = ["(All)"] + sorted(df["work_type"].dropna().astype(str).unique().tolist())
wt = st.sidebar.selectbox("Work Type", wt_options, index=0)

# Safe date range defaults
min_d = df["start"].min()
max_d = df["finish"].max()

# If something still odd, fallback to a sane window
if not isinstance(min_d, date) or not isinstance(max_d, date):
    today = date.today()
    min_d = today - timedelta(days=30)
    max_d = today + timedelta(days=365)

if min_d > max_d:
    min_d, max_d = max_d, min_d

d_range = st.sidebar.date_input(
    "Planned date range",
    value=(min_d, max_d),
    min_value=min_d,
    max_value=max_d,
)

# Apply filters
f = df.copy()
if major != "(All)":
    f = f[f["major_group"].astype(str) == major]
if pkg != "(All)":
    f = f[f["package_code"].astype(str) == pkg]
if wt != "(All)":
    f = f[f["work_type"].astype(str) == wt]

# Apply date overlap filter
if isinstance(d_range, (tuple, list)) and len(d_range) == 2:
    d1, d2 = d_range[0], d_range[1]
    f = f[(f["start"] <= d2) & (f["finish"] >= d1)]

# KPIs
c1, c2, c3, c4 = st.columns(4)
c1.metric("Activities", int(len(f)))
c2.metric("Earliest start", str(f["start"].min()) if not f.empty else "-")
c3.metric("Latest finish", str(f["finish"].max()) if not f.empty else "-")
if "is_milestone" in f.columns:
    c4.metric("Milestones", int(pd.to_numeric(f["is_milestone"], errors="coerce").fillna(0).sum()))
else:
    c4.metric("Milestones", "-")

st.divider()

# Charts
left, right = st.columns(2)

with left:
    st.subheader("Activities by Package")
    if f.empty:
        st.info("No data for current filters.")
    else:
        by_pkg = f.groupby("package_code", as_index=False).size().sort_values("size", ascending=False)
        st.altair_chart(
            alt.Chart(by_pkg).mark_bar().encode(
                x=alt.X("package_code:N", sort="-y", title="Package"),
                y=alt.Y("size:Q", title="Count"),
                tooltip=["package_code", "size"],
            ).properties(height=320),
            use_container_width=True
        )

with right:
    st.subheader("Work Type distribution")
    if f.empty:
        st.info("No data for current filters.")
    else:
        by_wt = f.groupby("work_type", as_index=False).size().sort_values("size", ascending=False)
        st.altair_chart(
            alt.Chart(by_wt).mark_bar().encode(
                y=alt.Y("work_type:N", sort="-x", title="Work Type"),
                x=alt.X("size:Q", title="Count"),
                tooltip=["work_type", "size"],
            ).properties(height=320),
            use_container_width=True
        )

st.subheader("Schedule table")
st.dataframe(f, use_container_width=True, height=460)

st.download_button(
    "Download filtered CSV",
    data=f.to_csv(index=False).encode("utf-8"),
    file_name="primavera_planning_filtered.csv",
    mime="text/csv",
)
