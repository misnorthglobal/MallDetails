from pathlib import Path
import re

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="UAE Mall Intelligence",
    page_icon="🏙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_NAME = "2gis_all_malls_company_details.xlsx"
WORKBOOK_CANDIDATES = [
    BASE_DIR / "data" / WORKBOOK_NAME,
    BASE_DIR / WORKBOOK_NAME,
]


def bundled_workbook():
    """Return the workbook wherever it was placed in the deployed repository."""
    return next((path for path in WORKBOOK_CANDIDATES if path.is_file()), None)


def clean_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u200b", " ")).strip()


def first_number(value):
    match = re.search(r"[\d,]+(?:\.\d+)?", clean_text(value))
    return float(match.group(0).replace(",", "")) if match else 0.0


@st.cache_data(show_spinner=False)
def load_data(source):
    malls = pd.read_excel(source, sheet_name="Malls")
    companies = pd.read_excel(source, sheet_name="Companies")

    malls.columns = [clean_text(c) for c in malls.columns]
    companies.columns = [clean_text(c) for c in companies.columns]
    for frame in (malls, companies):
        for col in frame.select_dtypes(include="object").columns:
            frame[col] = frame[col].map(clean_text)

    malls["Mall Number"] = pd.to_numeric(malls["Mall Number"], errors="coerce").astype("Int64")
    companies["Mall Number"] = pd.to_numeric(companies["Mall Number"], errors="coerce").astype("Int64")
    malls["Size (sq ft)"] = malls.get("Approximate Size", "").map(first_number)
    malls["Parking Numeric"] = pd.to_numeric(malls.get("Total Parking"), errors="coerce").fillna(0)
    malls["Floors Numeric"] = pd.to_numeric(malls.get("Floors"), errors="coerce").fillna(0)
    malls["Listed Companies"] = pd.to_numeric(malls.get("Companies"), errors="coerce").fillna(0)

    if "Operator" not in malls.columns:
        malls["Operator"] = "Not available in source Excel"

    actual = companies.groupby("Mall Number")["Company Name"].nunique()
    malls["Companies in file"] = malls["Mall Number"].map(actual).fillna(0).astype(int)
    companies["Category"] = companies.get("Category", "").replace("", "Uncategorized")
    return malls, companies


def csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


st.markdown("""
<style>
    .stApp {background: #f5f7fb;}
    [data-testid="stSidebar"] {background: #10243e;}
    [data-testid="stSidebar"] * {color: #f8fafc;}
    .hero {padding: 1.6rem 1.8rem; border-radius: 20px; color: white;
           background: linear-gradient(120deg,#0f2745,#0c7183); margin-bottom: 1rem;}
    .hero h1 {margin:0; font-size:2.15rem;}
    .hero p {margin:.45rem 0 0; opacity:.9;}
    .mall-card {background:white; border:1px solid #e2e8f0; border-radius:16px;
                padding:1rem 1.2rem; margin:.45rem 0; box-shadow:0 4px 14px rgba(15,23,42,.04);}
    .muted {color:#64748b; font-size:.9rem;}
    div[data-testid="stMetric"] {background:white; border:1px solid #e2e8f0;
        padding:14px; border-radius:14px; box-shadow:0 3px 12px rgba(15,23,42,.04);}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero"><h1>UAE Mall Intelligence</h1>
<p>Search malls, discover companies, and compare ownership, size, floors and parking.</p></div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.subheader("Data source")
    uploaded = st.file_uploader("Replace the Excel workbook", type=["xlsx"])
    default_file = bundled_workbook()
    source = uploaded if uploaded is not None else default_file

if source is None:
    st.warning(
        "The mall workbook is not included in this deployment. "
        "Upload **2gis_all_malls_company_details.xlsx** in the sidebar to continue."
    )
    st.info(
        "For a permanent Render deployment, commit the workbook at "
        "`data/2gis_all_malls_company_details.xlsx` in the same GitHub repository as `app.py`."
    )
    st.stop()

try:
    malls, companies = load_data(source)
except Exception as exc:
    st.error(f"Could not read the workbook: {exc}")
    st.stop()

with st.sidebar:
    st.subheader("Search & filters")
    query = st.text_input("Mall or company", placeholder="e.g. Dubai Mall or Carrefour")
    emirates = sorted(x for x in malls["Emirates"].dropna().unique() if x)
    chosen_emirates = st.multiselect("Emirates", emirates)
    owners = sorted(x for x in malls["Property Owner"].dropna().unique() if x)
    chosen_owners = st.multiselect("Property owner", owners)
    confidences = sorted(x for x in malls["Confidence"].dropna().unique() if x)
    chosen_confidence = st.multiselect("Ownership confidence", confidences)
    max_size = int(malls["Size (sq ft)"].max()) if len(malls) else 0
    size_range = st.slider("Approx. size (sq ft)", 0, max(max_size, 1), (0, max(max_size, 1)), step=max(max_size // 100, 1))
    min_companies = st.number_input("Minimum companies", min_value=0, value=0, step=10)

filtered = malls.copy()
if chosen_emirates:
    filtered = filtered[filtered["Emirates"].isin(chosen_emirates)]
if chosen_owners:
    filtered = filtered[filtered["Property Owner"].isin(chosen_owners)]
if chosen_confidence:
    filtered = filtered[filtered["Confidence"].isin(chosen_confidence)]
filtered = filtered[filtered["Size (sq ft)"].between(*size_range)]
filtered = filtered[filtered["Companies in file"] >= min_companies]

if query.strip():
    q = query.strip().lower()
    mall_hit = filtered.apply(
        lambda row: q in " ".join(clean_text(row.get(c, "")).lower() for c in
                                  ["Name", "Location", "Emirates", "Property Owner", "Operator"]), axis=1
    )
    company_malls = companies[
        companies[["Company Name", "Category"]].fillna("").astype(str)
        .apply(lambda col: col.str.lower().str.contains(q, regex=False)).any(axis=1)
    ]["Mall Number"].unique()
    filtered = filtered[mall_hit | filtered["Mall Number"].isin(company_malls)]

tabs = st.tabs(["Overview", "Mall explorer", "Company search", "Compare malls", "Data quality"])

with tabs[0]:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Matching malls", f"{len(filtered):,}")
    c2.metric("Companies in file", f"{filtered['Companies in file'].sum():,}")
    c3.metric("Total parking", f"{int(filtered['Parking Numeric'].sum()):,}")
    c4.metric("Median mall size", f"{filtered['Size (sq ft)'].median():,.0f} sq ft" if len(filtered) else "—")

    left, right = st.columns(2)
    with left:
        by_emirate = filtered.groupby("Emirates", as_index=False).size().sort_values("size", ascending=False)
        st.plotly_chart(px.bar(by_emirate, x="Emirates", y="size", color="size",
                               color_continuous_scale=["#9bd5d9", "#0c7183"],
                               labels={"size": "Malls"}, title="Malls by emirate"), use_container_width=True)
    with right:
        top_categories = (companies[companies["Mall Number"].isin(filtered["Mall Number"])]
                          .groupby("Category").size().nlargest(12).sort_values())
        st.plotly_chart(px.bar(x=top_categories.values, y=top_categories.index, orientation="h",
                               labels={"x": "Companies", "y": "Category"},
                               title="Most common company categories",
                               color=top_categories.values, color_continuous_scale=["#b8e1dc", "#126b76"]),
                        use_container_width=True)

    st.subheader("Largest malls in the current results")
    chart_data = filtered.nlargest(15, "Size (sq ft)")
    st.plotly_chart(px.scatter(chart_data, x="Size (sq ft)", y="Companies in file", size="Parking Numeric",
                               color="Emirates", hover_name="Name", title="Size, companies and parking"),
                    use_container_width=True)

with tabs[1]:
    if filtered.empty:
        st.info("No malls match the current filters.")
    else:
        options = filtered.sort_values("Name").apply(lambda r: f"{r['Name']} — {r['Emirates']}", axis=1).tolist()
        selection = st.selectbox("Choose a mall", options)
        idx = options.index(selection)
        mall = filtered.sort_values("Name").iloc[idx]
        mall_companies = companies[companies["Mall Number"] == mall["Mall Number"]].copy()

        st.header(mall["Name"])
        st.caption(f"{mall['Emirates']} · {mall['Location']}")
        a, b, c, d, e = st.columns(5)
        a.metric("Approx. size", mall.get("Approximate Size") or "—")
        b.metric("Companies", f"{len(mall_companies):,}")
        c.metric("Parking", f"{int(mall['Parking Numeric']):,}" if mall["Parking Numeric"] else "—")
        d.metric("Floors", f"{int(mall['Floors Numeric'])}" if mall["Floors Numeric"] else "—")
        e.metric("Confidence", mall.get("Confidence") or "—")

        p1, p2 = st.columns(2)
        p1.info(f"**Property owner:** {mall.get('Property Owner') or 'Not available'}")
        p2.info(f"**Operator:** {mall.get('Operator') or 'Not available'}")
        st.write(f"**Size range:** {mall.get('Size Range') or 'Not available'}")
        if mall.get("Remarks"):
            st.write(f"**Remarks:** {mall['Remarks']}")
        if mall.get("Link"):
            st.link_button("Open mall in 2GIS", mall["Link"])

        cat_filter = st.multiselect("Filter the directory by category", sorted(mall_companies["Category"].unique()))
        if cat_filter:
            mall_companies = mall_companies[mall_companies["Category"].isin(cat_filter)]
        show_cols = [c for c in ["Company Name", "Category", "Floors", "Branch Counts", "Company URL"] if c in mall_companies]
        st.dataframe(mall_companies[show_cols], use_container_width=True, hide_index=True,
                     column_config={"Company URL": st.column_config.LinkColumn("Company link")})
        st.download_button("Download this mall's companies", csv_bytes(mall_companies[show_cols]),
                           f"{re.sub(r'[^A-Za-z0-9]+', '_', mall['Name']).strip('_')}_companies.csv", "text/csv")

with tabs[2]:
    st.subheader("Search all companies")
    company_query = st.text_input("Company name or category", key="company_search", placeholder="Type a company or business category")
    pool = companies[companies["Mall Number"].isin(filtered["Mall Number"])].copy()
    if company_query:
        cq = company_query.lower()
        pool = pool[pool[["Company Name", "Category"]].fillna("").astype(str)
                    .apply(lambda col: col.str.lower().str.contains(cq, regex=False)).any(axis=1)]
    company_emirates = st.multiselect("Company emirates", sorted(pool["Emirates"].unique()), key="company_emirates")
    if company_emirates:
        pool = pool[pool["Emirates"].isin(company_emirates)]
    st.write(f"**{len(pool):,} matching company records**")
    company_cols = [c for c in ["Company Name", "Category", "Mall Name", "Emirates", "Floors", "Branch Counts", "Company URL"] if c in pool]
    st.dataframe(pool[company_cols], use_container_width=True, hide_index=True, height=560,
                 column_config={"Company URL": st.column_config.LinkColumn("Company link")})
    st.download_button("Download company results", csv_bytes(pool[company_cols]), "company_search_results.csv", "text/csv")

with tabs[3]:
    st.subheader("Side-by-side mall comparison")
    labels = filtered.sort_values("Name").apply(lambda r: f"{r['Name']} — {r['Emirates']}", axis=1).tolist()
    selected = st.multiselect("Select up to 8 malls", labels, max_selections=8)
    if selected:
        names = [x.rsplit(" — ", 1)[0] for x in selected]
        comp = filtered[filtered["Name"].isin(names)].copy()
        cols = ["Name", "Emirates", "Approximate Size", "Size Range", "Property Owner", "Operator",
                "Confidence", "Total Parking", "Floors", "Companies in file"]
        st.dataframe(comp[cols], use_container_width=True, hide_index=True)
        metrics = comp.melt(id_vars="Name", value_vars=["Size (sq ft)", "Parking Numeric", "Companies in file"],
                            var_name="Metric", value_name="Value")
        st.plotly_chart(px.bar(metrics, x="Name", y="Value", color="Name", facet_col="Metric",
                               facet_col_wrap=3, title="Comparison metrics"), use_container_width=True)
    else:
        st.info("Select two or more malls to compare them.")

with tabs[4]:
    st.subheader("Coverage and source completeness")
    checks = pd.DataFrame({
        "Field": ["Property owner", "Operator", "Approximate size", "Floors", "Parking", "Company records"],
        "Available": [
            (malls["Property Owner"].ne("") & ~malls["Property Owner"].str.contains("not publicly", case=False, na=False)).sum(),
            (malls["Operator"].ne("") & ~malls["Operator"].str.contains("not available", case=False, na=False)).sum(),
            malls["Size (sq ft)"].gt(0).sum(), malls["Floors Numeric"].gt(0).sum(),
            malls["Parking Numeric"].gt(0).sum(), malls["Companies in file"].gt(0).sum(),
        ]
    })
    checks["Missing / undisclosed"] = len(malls) - checks["Available"]
    checks["Coverage %"] = checks["Available"] / len(malls) * 100
    st.dataframe(checks, hide_index=True, use_container_width=True,
                 column_config={"Coverage %": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100)})
    st.warning("Operator is not present in the uploaded workbook. Add an 'Operator' column to the Malls sheet and the dashboard will use it automatically.")
    st.download_button("Download filtered mall results", csv_bytes(filtered.drop(columns=["Parking Numeric", "Floors Numeric"], errors="ignore")),
                       "filtered_malls.csv", "text/csv")

st.caption("UAE Mall Intelligence · Source: uploaded 2GIS mall and company workbook")
