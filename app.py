import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(page_title="Primavera Planning Dashboard", layout="wide")
st.title("Primavera Planning Dashboard (Instant)")
st.caption("Planning dashboard. Data source = extracted CSV committed by GitHub Actions.")

df = pd.read_csv("data/primavera.csv", parse_dates=["start", "finish"])
df["start"] = pd.to_datetime(df["start"]).dt.date
df["finish"] = pd.to_datetime(df["finish"]).dt.date

st.sidebar.header("Filters")
major = st.sidebar.selectbox("Major Group", ["(All)"] + sorted(df["major_group"].dropna().unique().tolist()))
pkg = st.sidebar.selectbox("Area / Package", ["(All)"] + sorted(df["package_code"].dropna().unique().tolist()))
wt = st.sidebar.selectbox("Work Type", ["(All)"] + sorted(df["work_type"].dropna().unique().tolist()))

min_d, max_d = df["start"].min(), df["finish"].max()
d1, d2 = st.sidebar.date_input("Planned date range", value=(min_d, max_d), min_value=min_d, max_value=max_d)

f = df.copy()
if major != "(All)":
    f = f[f["major_group"] == major]
if pkg != "(All)":
    f = f[f["package_code"] == pkg]
if wt != "(All)":
    f = f[f["work_type"] == wt]
f = f[(f["start"] <= d2) & (f["finish"] >= d1)]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Activities", len(f))
c2.metric("Earliest start", str(f["start"].min()) if len(f) else "-")
c3.metric("Latest finish", str(f["finish"].max()) if len(f) else "-")
c4.metric("Milestones", int(f["is_milestone"].sum()) if len(f) else 0)

st.divider()

left, right = st.columns(2)
with left:
    by_pkg = f.groupby("package_code", as_index=False).size()
    st.altair_chart(
        alt.Chart(by_pkg).mark_bar().encode(
            x=alt.X("package_code:N", sort="-y", title="Package"),
            y=alt.Y("size:Q", title="Count"),
            tooltip=["package_code", "size"],
        ).properties(height=320),
        use_container_width=True
    )

with right:
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
