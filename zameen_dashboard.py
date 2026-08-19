"""
Zameen.com Real Estate Analytics Dashboard — Full Analyst Version
--------------------------------------------------------------------------
Pages:
  1. Overview              — simple market snapshot + city-wise trend comparison
  2. City Explorer          — city -> area -> subarea, + price-vs-size outliers
  3. Hierarchy View          — step-by-step drill-down bars
  4. Compare Areas           — overlay trend lines
  5. Area Rankings           — hottest/coolest/priciest + radar comparison
  6. Investment Insights     — opportunity quadrant, best value areas, ROI score
  7. Amenities Impact        — which features raise price + room/space features
  8. Market Activity          — listing freshness / turnover
  9. Browse Listings        — search + sortable table

Install once:
    pip install streamlit pandas plotly openpyxl numpy

Run:
    streamlit run zameen_dashboard.py
"""

import re
import numpy as np
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Real Estate Analysis", layout="wide", page_icon="🏠")

MASTER_FILE = "zameen_master.xlsx"
ACCENT = "#16a34a"
ACCENT2 = "#dc2626"
ACCENT3 = "#3b82f6"
CHART_TEMPLATE = "plotly_white"
PALETTE = px.colors.qualitative.Set2

AMENITY_KEYWORDS = [
    "Gym", "Swimming Pool", "Mosque", "Security Staff", "Community Centre",
    "Kids Play Area", "Barbeque Area", "Lawn or Garden", "Elevators",
    "Broadband Internet Access", "Maintenance Staff", "Day Care Centre",
    "Central Air Conditioning", "Servant Quarters",
]

# ---------------------------------------------------------------------------
# STYLING — clean white theme, polished sidebar & cards
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}

    html, body, [class*="css"] { font-family: 'Segoe UI', 'Inter', sans-serif; }
    .stApp { background-color: #f8fafc; }
    .block-container {padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1400px;}

    /* Header banner */
    .dash-header {
        background: linear-gradient(90deg, #16a34a 0%, #22c55e 100%);
        border-radius: 16px;
        padding: 22px 28px;
        margin-bottom: 22px;
        box-shadow: 0 6px 20px rgba(22,163,74,0.18);
    }
    .dash-header h1 {
        color: #ffffff !important; font-weight: 800 !important;
        font-size: 1.9rem !important; margin: 0 !important; line-height: 1.3;
    }
    .dash-header p {
        color: #ecfdf5 !important; margin: 4px 0 0 0 !important; font-size: 0.95rem;
    }

    /* KPI metric cards */
    [data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-left: 4px solid #16a34a;
        border-radius: 12px; padding: 14px 16px;
        box-shadow: 0 2px 8px rgba(15,23,42,0.05);
    }
    [data-testid="stMetricLabel"] { color: #64748b !important; font-weight: 600; font-size: 0.78rem;}
    [data-testid="stMetricValue"] { color: #0f172a !important; font-size: 1.25rem !important; font-weight: 700 !important;}

    h2, h3 { color: #0f172a; font-weight: 700 !important; }
    h3 { font-size: 1.05rem !important; margin-top: 0.3rem !important; }

    [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; border: 1px solid #e2e8f0; }

    hr { margin: 1.2rem 0 !important; border-color: #e2e8f0 !important; }

    /* Sidebar — clean white with subtle border */
    section[data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    section[data-testid="stSidebar"] .block-container { padding-top: 1.5rem; }
    section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #0f172a !important; }
    section[data-testid="stSidebar"] label { color: #334155 !important; font-weight: 500; }

    /* Sidebar nav radio -> pill style */
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 9px 12px;
        margin-bottom: 6px;
        width: 100%;
        transition: all 0.15s ease;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background: #ecfdf5; border-color: #16a34a;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label div:first-child {
        border-color: #16a34a !important;
    }

    /* Dropdowns / selectboxes / multiselects everywhere */
    div[data-baseweb="select"] > div {
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
    }
    div[data-baseweb="select"] > div:hover { border-color: #16a34a !important; }
    /* Multiselect: compact pills, scrollable box instead of expanding endlessly */
    div[data-baseweb="select"] {
        max-height: 110px;
        overflow-y: auto;
    }
    span[data-baseweb="tag"] {
        background-color: #16a34a !important; border-radius: 6px !important;
        font-size: 0.78rem !important; padding: 2px 6px !important; margin: 2px !important;
    }

    /* Sidebar filter section card */
    .filter-card {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 14px 14px 4px 14px; margin-bottom: 14px;
    }

    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] { border-radius: 10px 10px 0 0; padding: 10px 16px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)


ROOM_FEATURES = {
    "Kitchens": "count", "Parking Spaces": "count", "Servant Quarters": "count",
    "Store Rooms": "count", "Elevators": "count", "Floors": "count",
    "Drawing Room": "flag", "Dining Room": "flag", "Study Room": "flag",
    "Prayer Room": "flag", "Powder Room": "flag", "Lounge or Sitting Room": "flag",
    "Laundry Room": "flag", "Built in year": "year",
}


def parse_room_features(amenities_raw):
    """
    Parses tokens like: 'Kitchens; : 1; Servant Quarters; : 1; Drawing Room; ...'
    into a dict of {feature_name: numeric_value_or_True}.
    """
    result = {}
    if not isinstance(amenities_raw, str):
        return result
    tokens = [t.strip() for t in amenities_raw.split(";") if t.strip()]
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if tok.startswith(":"):
            i += 1
            continue
        value = True
        if i + 1 < len(tokens) and tokens[i + 1].startswith(":"):
            raw_val = tokens[i + 1].replace(":", "").strip()
            try:
                value = float(raw_val)
            except ValueError:
                value = True
            i += 2
        else:
            i += 1
        if tok in ROOM_FEATURES:
            result[tok] = value
    return result


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def fmt_price(value):
    if pd.isna(value):
        return "-"
    if value >= 1_00_00_000:
        return f"PKR {value/1_00_00_000:.2f} Crore"
    elif value >= 1_00_000:
        return f"PKR {value/1_00_000:.1f} Lakh"
    return f"PKR {value:,.0f}"


def fmt_price_short(value):
    if pd.isna(value):
        return "-"
    if value >= 1_00_00_000:
        return f"{value/1_00_00_000:.2f} Cr"
    elif value >= 1_00_000:
        return f"{value/1_00_000:.1f} Lac"
    return f"{value:,.0f}"


def parse_area(area_raw):
    if not isinstance(area_raw, str):
        return None, None
    m = re.match(r"^\s*([\d,\.]+)\s*(.+?)\s*$", area_raw)
    if not m:
        return None, None
    try:
        size = float(m.group(1).replace(",", ""))
    except ValueError:
        return None, None
    return size, m.group(2).strip()


def parse_days_ago(added_raw):
    if not isinstance(added_raw, str):
        return None
    s = added_raw.lower()
    if "today" in s or "hour" in s:
        return 0
    if "yesterday" in s:
        return 1
    m = re.search(r"(\d+)\s*day", s)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s*week", s)
    if m:
        return int(m.group(1)) * 7
    m = re.search(r"(\d+)\s*month", s)
    if m:
        return int(m.group(1)) * 30
    return None


def freshness_bucket(days):
    if pd.isna(days):
        return None
    if days <= 7: return "0-7 days"
    if days <= 14: return "1-2 weeks"
    if days <= 30: return "2-4 weeks"
    if days <= 90: return "1-3 months"
    return "3+ months"


FRESHNESS_ORDER = ["0-7 days", "1-2 weeks", "2-4 weeks", "1-3 months", "3+ months"]


def add_pct_labels(df, count_col="count"):
    total = df[count_col].sum()
    df = df.copy()
    df["pct"] = (df[count_col] / total * 100).round(1)
    df["label"] = df.apply(lambda r: f"{int(r[count_col]):,} ({r['pct']}%)", axis=1)
    return df


def dominant_unit(df):
    if "area_unit" not in df.columns or df["area_unit"].dropna().empty:
        return None
    return df["area_unit"].value_counts().idxmax()


@st.cache_data
def load_data():
    listings = pd.read_excel(MASTER_FILE, sheet_name="Listings")
    try:
        trend_history = pd.read_excel(MASTER_FILE, sheet_name="TrendHistory")
        trend_history["month_year"] = pd.to_datetime(trend_history["month_year"], errors="coerce")
    except Exception:
        trend_history = pd.DataFrame()

    if "area_raw" in listings.columns:
        parsed = listings["area_raw"].apply(parse_area)
        listings["area_size"] = parsed.apply(lambda x: x[0])
        listings["area_unit"] = parsed.apply(lambda x: x[1])

    if "added_raw" in listings.columns:
        listings["days_ago"] = listings["added_raw"].apply(parse_days_ago)
        listings["freshness"] = listings["days_ago"].apply(freshness_bucket)

    if "amenities_raw" in listings.columns:
        for kw in AMENITY_KEYWORDS:
            listings[f"has_{kw}"] = listings["amenities_raw"].str.contains(kw, case=False, na=False)
        listings["amenity_count"] = listings[[f"has_{kw}" for kw in AMENITY_KEYWORDS]].sum(axis=1)

        room_parsed = listings["amenities_raw"].apply(parse_room_features)
        for feature, kind in ROOM_FEATURES.items():
            listings[f"has_{feature}"] = room_parsed.apply(lambda d: feature in d)
            if kind in ("count", "year"):
                numeric_col = "built_year" if feature == "Built in year" else f"{feature.lower().replace(' ', '_')}_count"
                listings[numeric_col] = room_parsed.apply(
                    lambda d: d.get(feature) if isinstance(d.get(feature), float) else None)

    return listings, trend_history


def trend_line_for(trend_ids, trend_history):
    if trend_history.empty or not len(trend_ids):
        return pd.DataFrame()
    data = trend_history[trend_history["trend_id"].isin(trend_ids)]
    if data.empty:
        return pd.DataFrame()
    return data.groupby("month_year", as_index=False)["avg_price"].mean().sort_values("month_year")


def price_trend_figure(monthly_df, color=ACCENT2):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=monthly_df["month_year"], y=monthly_df["avg_price"],
        mode="lines", line=dict(color=color, width=3),
        fill="tozeroy", fillcolor="rgba(220,38,38,0.06)",
    ))
    fig.update_layout(
        template=CHART_TEMPLATE, height=400, xaxis_title=None, yaxis_title="Average Price (PKR)",
        xaxis=dict(
            rangeselector=dict(
                buttons=[
                    dict(count=6, label="6 Months", step="month", stepmode="backward"),
                    dict(count=1, label="1 Year", step="year", stepmode="backward"),
                    dict(count=2, label="2 Years", step="year", stepmode="backward"),
                    dict(step="all", label="Max"),
                ], bgcolor="#f1f5f9", activecolor=ACCENT,
            ), rangeslider=dict(visible=False),
        ), margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified",
    )
    return fig


def bar_with_labels(df, x, y, title_x=None, title_y=None, orientation="v",
                     color=ACCENT, text=None, height=380):
    if orientation == "h":
        fig = px.bar(df, x=y, y=x, text=text, orientation="h", color_discrete_sequence=[color])
        fig.update_layout(yaxis=dict(autorange="reversed"), xaxis_title=title_y, yaxis_title=title_x)
    else:
        fig = px.bar(df, x=x, y=y, text=text, color_discrete_sequence=[color])
        fig.update_layout(xaxis_title=title_x, yaxis_title=title_y)
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=12, color="#0f172a"),
        marker=dict(line=dict(width=0)),
        cliponaxis=False,
    )
    fig.update_layout(
        template=CHART_TEMPLATE, margin=dict(l=10, r=10, t=30, b=10), height=height,
        uniformtext_minsize=9, uniformtext_mode="hide",
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#334155", size=12),
        xaxis=dict(tickfont=dict(size=11)), yaxis=dict(tickfont=dict(size=11)),
    )
    return fig


def pie_with_labels(df, names, values, height=360, colors=None):
    fig = px.pie(df, names=names, values=values, hole=0.5,
                 color_discrete_sequence=colors or px.colors.sequential.Greens_r)
    fig.update_traces(
        textinfo="percent+value", textposition="outside",
        textfont=dict(size=12, color="#0f172a"),
        marker=dict(line=dict(color="#ffffff", width=2)),
    )
    fig.update_layout(
        template=CHART_TEMPLATE, margin=dict(l=10, r=10, t=10, b=10), height=height,
        legend=dict(orientation="h", yanchor="top", y=-0.05, font=dict(size=11)),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#334155"),
    )
    return fig


def price_bucket(price_numeric):
    if pd.isna(price_numeric):
        return None
    cr = price_numeric / 1_00_00_000
    if cr < 0.5: return "< 50 Lakh"
    if cr < 1: return "50 Lakh - 1 Cr"
    if cr < 2: return "1 - 2 Cr"
    if cr < 5: return "2 - 5 Cr"
    if cr < 10: return "5 - 10 Cr"
    return "10 Cr+"


BUCKET_ORDER = ["< 50 Lakh", "50 Lakh - 1 Cr", "1 - 2 Cr", "2 - 5 Cr", "5 - 10 Cr", "10 Cr+"]


def normalize(series):
    lo, hi = series.min(), series.max()
    if hi == lo:
        return series * 0 + 0.5
    return (series - lo) / (hi - lo)


# ---------------------------------------------------------------------------
# LOAD DATA
# ---------------------------------------------------------------------------
listings, trend_history = load_data()
n_cities = listings["city"].nunique() if "city" in listings.columns else 0

st.markdown(f"""
<div class="dash-header">
    <h1>🏠 Real Estate Analysis</h1>
    <p>Live analysis of {len(listings):,} properties across {n_cities} cities in Pakistan</p>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown("## 📑 Navigate")
page = st.sidebar.radio("Go to", [
    "🏠 Overview", "🏙️ City Explorer", "🌳 Hierarchy View",
    "⚖️ Compare Areas", "🏆 Area Rankings", "💡 Investment Insights",
    "🛋️ Amenities Impact", "⏱️ Market Activity",
    "🔍 Browse Listings",
], label_visibility="collapsed")

st.sidebar.markdown("---")
st.sidebar.markdown("## 🔎 Global Filters")
st.sidebar.caption("Applies across every page")
st.sidebar.markdown('<div class="filter-card">', unsafe_allow_html=True)

_cities_list = ["All Cities"] + sorted(listings["city"].dropna().unique())
GLOBAL_CITY = st.sidebar.selectbox("City", _cities_list, key="global_city_filter")
GLOBAL_POOL = listings if GLOBAL_CITY == "All Cities" else listings[listings["city"] == GLOBAL_CITY]

_areas_list = ["All Areas"] + sorted(GLOBAL_POOL["breadcrumb_area"].dropna().unique())
GLOBAL_AREA = st.sidebar.selectbox("Area", _areas_list, key="global_area_filter")
if GLOBAL_AREA != "All Areas":
    GLOBAL_POOL = GLOBAL_POOL[GLOBAL_POOL["breadcrumb_area"] == GLOBAL_AREA]

_subareas_list = ["All Sub-areas"] + sorted(GLOBAL_POOL["breadcrumb_subarea"].dropna().unique())
GLOBAL_SUBAREA = st.sidebar.selectbox("Sub-area", _subareas_list, key="global_subarea_filter")
if GLOBAL_SUBAREA != "All Sub-areas":
    GLOBAL_POOL = GLOBAL_POOL[GLOBAL_POOL["breadcrumb_subarea"] == GLOBAL_SUBAREA]

_types_list = sorted(GLOBAL_POOL["type"].dropna().unique())
GLOBAL_TYPES = st.sidebar.multiselect("Property Type", _types_list, default=_types_list, key="global_type_filter")
if GLOBAL_TYPES:
    GLOBAL_POOL = GLOBAL_POOL[GLOBAL_POOL["type"].isin(GLOBAL_TYPES)]

if GLOBAL_POOL["price_numeric"].notna().any():
    _pmin, _pmax = float(GLOBAL_POOL["price_numeric"].min()), float(GLOBAL_POOL["price_numeric"].max())
    if _pmin < _pmax:
        _price_range = st.sidebar.slider(
            "Price Range", _pmin, _pmax, (_pmin, _pmax),
            format="%.0f", key="global_price_filter",
            help="In PKR"
        )
        st.sidebar.caption(f"{fmt_price_short(_price_range[0])} — {fmt_price_short(_price_range[1])}")
        GLOBAL_POOL = GLOBAL_POOL[
            (GLOBAL_POOL["price_numeric"] >= _price_range[0]) & (GLOBAL_POOL["price_numeric"] <= _price_range[1])
        ]

_beds_list = sorted(GLOBAL_POOL["bedrooms"].dropna().unique()) if "bedrooms" in GLOBAL_POOL.columns else []
GLOBAL_BEDS = st.sidebar.multiselect("Bedrooms", _beds_list, key="global_beds_filter")
if GLOBAL_BEDS:
    GLOBAL_POOL = GLOBAL_POOL[GLOBAL_POOL["bedrooms"].isin(GLOBAL_BEDS)]

if "kitchens_count" in GLOBAL_POOL.columns:
    _kitchens_list = sorted(GLOBAL_POOL["kitchens_count"].dropna().unique())
    if _kitchens_list:
        GLOBAL_KITCHENS = st.sidebar.multiselect("Kitchens", [int(k) for k in _kitchens_list], key="global_kitchens_filter")
        if GLOBAL_KITCHENS:
            GLOBAL_POOL = GLOBAL_POOL[GLOBAL_POOL["kitchens_count"].isin(GLOBAL_KITCHENS)]

if "area_size" in GLOBAL_POOL.columns and GLOBAL_POOL["area_size"].notna().any():
    _dom_unit = dominant_unit(GLOBAL_POOL)
    _sized_pool = GLOBAL_POOL[GLOBAL_POOL["area_unit"] == _dom_unit] if _dom_unit else GLOBAL_POOL.dropna(subset=["area_size"])
    if not _sized_pool.empty:
        _smin, _smax = float(_sized_pool["area_size"].min()), float(_sized_pool["area_size"].max())
        if _smin < _smax:
            _size_range = st.sidebar.slider(
                f"Size ({_dom_unit or 'units'})", _smin, _smax, (_smin, _smax), key="global_size_filter"
            )
            GLOBAL_POOL = GLOBAL_POOL[
                GLOBAL_POOL["area_size"].isna() |
                ((GLOBAL_POOL["area_size"] >= _size_range[0]) & (GLOBAL_POOL["area_size"] <= _size_range[1]))
            ]

st.sidebar.caption(f"📊 **{len(GLOBAL_POOL):,}** listings match these filters")
st.sidebar.markdown('</div>', unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.caption("💡 These filters apply everywhere. Some pages also have their own extra dropdowns for deeper drill-down.")


# =============================================================================
# PAGE 1 — OVERVIEW
# =============================================================================
if page == "🏠 Overview":
    st.header("🏠 Overview")
    st.caption("A simple, at-a-glance summary of the entire market")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Listings", f"{len(listings):,}")
    c2.metric("Average Price", fmt_price(listings["price_numeric"].mean()))
    c3.metric("Median Price", fmt_price(listings["price_numeric"].median()))
    c4.metric("Cities Covered", n_cities)

    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("① Listings by City")
        city_counts = listings["city"].value_counts().reset_index()
        city_counts.columns = ["city", "count"]
        city_counts = add_pct_labels(city_counts)
        st.plotly_chart(bar_with_labels(city_counts, "city", "count", "City", "Listings", text="label"),
                        use_container_width=True)
    with c2:
        st.subheader("② Property Type Split")
        if "type" in listings.columns:
            tc = listings["type"].value_counts().reset_index()
            tc.columns = ["type", "count"]
            st.plotly_chart(pie_with_labels(tc, "type", "count"), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("③ Price Range Distribution")
        pb = listings["price_numeric"].apply(price_bucket).value_counts().reindex(BUCKET_ORDER).dropna().reset_index()
        pb.columns = ["bucket", "count"]
        pb = add_pct_labels(pb)
        st.plotly_chart(bar_with_labels(pb, "bucket", "count", "Price Range", "Listings", text="label"),
                        use_container_width=True)
    with c2:
        st.subheader("④ Bedrooms Distribution")
        if "bedrooms" in listings.columns:
            bc = listings["bedrooms"].dropna().value_counts().sort_index().reset_index()
            bc.columns = ["bedrooms", "count"]
            bc = add_pct_labels(bc)
            st.plotly_chart(bar_with_labels(bc, "bedrooms", "count", "Bedrooms", "Listings", text="label"),
                            use_container_width=True)

    st.subheader("⑤ Overall Market Price Trend")
    if "trend_id" in listings.columns:
        monthly = trend_line_for(listings["trend_id"].dropna().unique(), trend_history)
        if not monthly.empty:
            st.plotly_chart(price_trend_figure(monthly), use_container_width=True)
        else:
            st.info("No trend data available yet.")

    st.subheader("⑥ City-wise Price Trend Comparison")
    if "trend_id" in listings.columns and "city" in listings.columns:
        fig_city_trend = go.Figure()
        for i, city_name in enumerate(sorted(listings["city"].dropna().unique())):
            city_ids = listings[listings["city"] == city_name]["trend_id"].dropna().unique()
            city_monthly = trend_line_for(city_ids, trend_history)
            if not city_monthly.empty:
                fig_city_trend.add_trace(go.Scatter(
                    x=city_monthly["month_year"], y=city_monthly["avg_price"],
                    mode="lines", name=city_name,
                    line=dict(width=2.5, color=PALETTE[i % len(PALETTE)]),
                ))
        fig_city_trend.update_layout(
            template=CHART_TEMPLATE, height=500, yaxis_title="Average Price (PKR)",
            legend=dict(orientation="h", yanchor="top", y=-0.18, x=0.5, xanchor="center", font=dict(size=11)),
            margin=dict(l=10, r=10, t=60, b=10), hovermode="x unified",
            xaxis=dict(
                rangeselector=dict(
                    buttons=[
                        dict(count=6, label="6 Months", step="month", stepmode="backward"),
                        dict(count=1, label="1 Year", step="year", stepmode="backward"),
                        dict(count=2, label="2 Years", step="year", stepmode="backward"),
                        dict(step="all", label="Max"),
                    ], bgcolor="#f1f5f9", activecolor=ACCENT, y=1.15,
                ), rangeslider=dict(visible=False),
            ),
        )
        st.plotly_chart(fig_city_trend, use_container_width=True)

    st.subheader("⑦ Price Change by City — Last Month / 6M / 1Y / 2Y")
    if "trend_id" in listings.columns and "city" in listings.columns and not trend_history.empty:
        rows = []
        for city_name in sorted(listings["city"].dropna().unique()):
            city_ids = listings[listings["city"] == city_name]["trend_id"].dropna().unique()
            city_monthly = trend_line_for(city_ids, trend_history)
            last_month_pct = None
            if len(city_monthly) >= 2:
                prev, latest = city_monthly["avg_price"].iloc[-2], city_monthly["avg_price"].iloc[-1]
                if prev:
                    last_month_pct = round((latest - prev) / prev * 100, 1)
            city_rows = listings[listings["city"] == city_name]
            rows.append({
                "City": city_name,
                "Last Month %": last_month_pct,
                "6-Month %": round(city_rows["trend_change_pct_6mo"].mean(), 1) if "trend_change_pct_6mo" in city_rows.columns else None,
                "1-Year %": round(city_rows["trend_change_pct_1yr"].mean(), 1) if "trend_change_pct_1yr" in city_rows.columns else None,
                "2-Year %": round(city_rows["trend_change_pct_2yr"].mean(), 1) if "trend_change_pct_2yr" in city_rows.columns else None,
            })
        change_df = pd.DataFrame(rows)
        st.dataframe(change_df, use_container_width=True, hide_index=True)
        st.caption("Last Month % is computed directly from the monthly price history; "
                   "6M/1Y/2Y % come from Zameen's own price index for each listing, averaged per city.")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⑧ Average Price by Property Type")
        if "type" in listings.columns:
            tp = listings.dropna(subset=["type", "price_numeric"]).groupby("type")["price_numeric"].mean().reset_index()
            tp["label"] = tp["price_numeric"].apply(fmt_price_short)
            st.plotly_chart(bar_with_labels(tp, "type", "price_numeric", "Property Type", "Avg Price", text="label"),
                            use_container_width=True)
    with c2:
        st.subheader("⑨ Average Price by City")
        cp = listings.dropna(subset=["city", "price_numeric"]).groupby("city")["price_numeric"].mean().reset_index()
        cp["label"] = cp["price_numeric"].apply(fmt_price_short)
        cp = cp.sort_values("price_numeric", ascending=False)
        st.plotly_chart(bar_with_labels(cp, "city", "price_numeric", "City", "Avg Price", text="label", color=ACCENT2),
                        use_container_width=True)

    st.subheader("⑩ Top 10 Areas by Number of Listings")
    if "breadcrumb_area" in listings.columns:
        ac = listings["breadcrumb_area"].dropna().value_counts().head(10).reset_index()
        ac.columns = ["area", "count"]
        ac = add_pct_labels(ac)
        st.plotly_chart(bar_with_labels(ac, "area", "count", "Listings", "Area", orientation="h",
                                        text="label", height=420), use_container_width=True)

    st.subheader("⑪ Bathrooms Distribution")
    if "bathrooms" in listings.columns:
        bc = listings["bathrooms"].dropna().value_counts().sort_index().reset_index()
        bc.columns = ["bathrooms", "count"]
        bc = add_pct_labels(bc)
        st.plotly_chart(bar_with_labels(bc, "bathrooms", "count", "Bathrooms", "Listings", text="label", color=ACCENT2),
                        use_container_width=True)


# =============================================================================
# PAGE 2 — CITY EXPLORER (+ price-vs-size outliers / best value listings)
# =============================================================================
elif page == "🏙️ City Explorer":
    st.header("🏙️ City Explorer")

    sel_city = GLOBAL_CITY
    city_df = GLOBAL_POOL

    areas = sorted(city_df["breadcrumb_area"].dropna().unique())
    sel_area = st.selectbox("Select an Area (optional)", ["All Areas"] + areas)
    scope_df = city_df if sel_area == "All Areas" else city_df[city_df["breadcrumb_area"] == sel_area]

    subareas = sorted(scope_df["breadcrumb_subarea"].dropna().unique())
    sel_subarea = st.selectbox("Select a Sub-area (optional)", ["All Sub-areas"] + subareas)
    if sel_subarea != "All Sub-areas":
        scope_df = scope_df[scope_df["breadcrumb_subarea"] == sel_subarea]

    unit = dominant_unit(scope_df)
    scope_label = sel_subarea if sel_subarea != "All Sub-areas" else (
        sel_area if sel_area != "All Areas" else sel_city)

    st.write("")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Listings", f"{len(scope_df):,}")
    c2.metric("Avg Price", fmt_price(scope_df["price_numeric"].mean()))
    c3.metric("Median Price", fmt_price(scope_df["price_numeric"].median()))
    if "trend_change_pct_overall" in scope_df.columns and scope_df["trend_change_pct_overall"].notna().any():
        c4.metric("Market Trend", f"{scope_df['trend_change_pct_overall'].mean():+.1f}%")
    c5.metric("Common Size Unit", unit or "-")

    st.write("")
    st.subheader(f"① Price Trend — {scope_label}")
    if "trend_id" in scope_df.columns:
        monthly = trend_line_for(scope_df["trend_id"].dropna().unique(), trend_history)
        if not monthly.empty:
            st.plotly_chart(price_trend_figure(monthly), use_container_width=True)
        else:
            st.info("No trend history available for this selection.")

    c1, c2 = st.columns(2)
    group_col = "city" if sel_city == "All Cities" else ("breadcrumb_area" if sel_area == "All Areas" else "breadcrumb_subarea")
    group_label = "City" if sel_city == "All Cities" else ("Area" if sel_area == "All Areas" else "Sub-area")

    with c1:
        st.subheader(f"② {group_label}-wise Listings")
        gc = scope_df[group_col].dropna().value_counts().head(15).reset_index()
        gc.columns = ["name", "count"]
        gc = add_pct_labels(gc)
        st.plotly_chart(bar_with_labels(gc, "name", "count", "Listings", None, orientation="h",
                                        text="label", height=420), use_container_width=True)
    with c2:
        st.subheader(f"③ Average Price by {group_label}")
        gp = scope_df.dropna(subset=[group_col, "price_numeric"]).groupby(group_col)["price_numeric"].mean()
        gp = gp.sort_values(ascending=False).head(15).reset_index()
        gp.columns = ["name", "price_numeric"]
        gp["label"] = gp["price_numeric"].apply(fmt_price_short)
        st.plotly_chart(bar_with_labels(gp, "name", "price_numeric", "Avg Price", None, orientation="h",
                                        text="label", height=420, color=ACCENT2), use_container_width=True)

    st.subheader("④ Property Type Split")
    if "type" in scope_df.columns:
        tc = scope_df["type"].value_counts().reset_index()
        tc.columns = ["type", "count"]
        st.plotly_chart(pie_with_labels(tc, "type", "count"), use_container_width=True)

    st.subheader("⑤ Price Distribution")
    if scope_df["price_numeric"].notna().any():
        fig = px.histogram(scope_df, x="price_numeric", nbins=30, color_discrete_sequence=[ACCENT])
        fig.update_traces(texttemplate="%{y}", textposition="outside")
        fig.update_layout(template=CHART_TEMPLATE, xaxis_title="Price (PKR)", yaxis_title="Listings",
                           margin=dict(l=10, r=10, t=10, b=10), height=360)
        st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⑥ Average Price by Bedrooms")
        if "bedrooms" in scope_df.columns:
            bp = scope_df.dropna(subset=["bedrooms", "price_numeric"]).groupby("bedrooms")["price_numeric"].mean().reset_index()
            bp = bp.sort_values("bedrooms")
            bp["label"] = bp["price_numeric"].apply(fmt_price_short)
            st.plotly_chart(bar_with_labels(bp, "bedrooms", "price_numeric", "Bedrooms", "Avg Price", text="label"),
                            use_container_width=True)
    with c2:
        title_unit = unit or "Size"
        st.subheader(f"⑦ Average Price by Size ({title_unit})")
        sized = scope_df.dropna(subset=["area_size", "price_numeric"])
        if unit:
            sized = sized[sized["area_unit"] == unit]
        if len(sized) >= 5:
            try:
                sized = sized.copy()
                sized["size_bucket"] = pd.qcut(sized["area_size"], q=min(5, sized["area_size"].nunique()), duplicates="drop")
                sb = sized.groupby("size_bucket")["price_numeric"].mean().reset_index()
                sb["size_bucket"] = sb["size_bucket"].astype(str)
                sb["label"] = sb["price_numeric"].apply(fmt_price_short)
                st.plotly_chart(bar_with_labels(sb, "size_bucket", "price_numeric", f"{title_unit} Range", "Avg Price", text="label"),
                                use_container_width=True)
            except Exception:
                st.info("Not enough size data to build this chart.")
        else:
            st.info("Not enough size data to build this chart.")

    # NEW: Price vs Size scatter with outlier detection
    st.subheader(f"⑧ Price vs Size — Spot the Outliers ({unit or 'Size'})")
    sized = scope_df.dropna(subset=["area_size", "price_numeric"])
    if unit:
        sized = sized[sized["area_unit"] == unit]
    if len(sized) >= 8:
        x = sized["area_size"].values
        y = sized["price_numeric"].values
        coeffs = np.polyfit(x, y, 1)
        predicted = np.polyval(coeffs, x)
        sized = sized.copy()
        sized["predicted_price"] = predicted
        sized["deal_pct"] = ((sized["price_numeric"] - sized["predicted_price"]) / sized["predicted_price"] * 100).round(1)
        sized["deal_type"] = np.where(sized["deal_pct"] < -15, "Good Deal (below market)",
                                np.where(sized["deal_pct"] > 15, "Overpriced (above market)", "Fair Price"))
        fig = px.scatter(sized, x="area_size", y="price_numeric", color="deal_type",
                         hover_data=["title"] if "title" in sized.columns else None,
                         color_discrete_map={"Good Deal (below market)": ACCENT,
                                             "Overpriced (above market)": ACCENT2,
                                             "Fair Price": "#94a3b8"})
        x_line = np.linspace(x.min(), x.max(), 50)
        fig.add_trace(go.Scatter(x=x_line, y=np.polyval(coeffs, x_line), mode="lines",
                                 name="Market Average Line", line=dict(color="#1e293b", dash="dash")))
        fig.update_layout(template=CHART_TEMPLATE, xaxis_title=f"Size ({unit})", yaxis_title="Price (PKR)",
                          margin=dict(l=10, r=10, t=10, b=10), height=440)
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("⑨ 💎 Best Value Listings (priced below market rate for their size)")
        deals = sized[sized["deal_type"] == "Good Deal (below market)"].sort_values("deal_pct").head(10)
        deal_cols = [c for c in ["title", "price_raw", "area_raw", "bedrooms", "deal_pct", "url"] if c in deals.columns]
        st.dataframe(deals[deal_cols], use_container_width=True, hide_index=True,
                    column_config={"url": st.column_config.LinkColumn("Link"), "deal_pct": "% Below Market"})
    else:
        st.info("Not enough data points to detect outliers for this selection.")

    st.subheader("⑩ Room & Space Features Present")
    room_flag_cols = [c for c in ["has_Kitchens", "has_Drawing Room", "has_Dining Room",
                                   "has_Servant Quarters", "has_Store Rooms"] if c in scope_df.columns]
    if room_flag_cols:
        rows = [{"Feature": c.replace("has_", ""), "count": scope_df[c].sum()} for c in room_flag_cols]
        rf_df = pd.DataFrame(rows).sort_values("count", ascending=False)
        rf_df = add_pct_labels(rf_df)
        st.plotly_chart(bar_with_labels(rf_df, "Feature", "count", "Listings", None, orientation="h",
                                        text="label", height=340), use_container_width=True)


# =============================================================================
# PAGE 3 — HIERARCHY VIEW
# =============================================================================
elif page == "🌳 Hierarchy View":
    st.header("🌳 Hierarchy View")
    st.caption("Simple step-by-step breakdown: pick a city, then an area, to drill deeper")

    st.subheader("① Listings by City")
    city_counts = listings["city"].value_counts().reset_index()
    city_counts.columns = ["city", "count"]
    city_counts = add_pct_labels(city_counts)
    st.plotly_chart(bar_with_labels(city_counts, "city", "count", "City", "Listings", text="label"),
                    use_container_width=True)

    st.subheader("② Average Price by City")
    cp = listings.dropna(subset=["city", "price_numeric"]).groupby("city")["price_numeric"].mean().reset_index()
    cp["label"] = cp["price_numeric"].apply(fmt_price_short)
    cp = cp.sort_values("price_numeric", ascending=False)
    st.plotly_chart(bar_with_labels(cp, "city", "price_numeric", "City", "Avg Price", text="label", color=ACCENT2),
                    use_container_width=True)

    st.markdown("---")
    cities = ["All Cities"] + sorted(listings["city"].dropna().unique())
    sel_city = st.selectbox("🔽 Drill into a City", cities)
    pool = listings if sel_city == "All Cities" else listings[listings["city"] == sel_city]

    st.subheader(f"③ Top 15 Areas in {sel_city}")
    ac = pool["breadcrumb_area"].dropna().value_counts().head(15).reset_index()
    ac.columns = ["area", "count"]
    ac = add_pct_labels(ac)
    st.plotly_chart(bar_with_labels(ac, "area", "count", "Listings", None, orientation="h",
                                    text="label", height=450), use_container_width=True)

    st.subheader(f"④ Average Price by Area in {sel_city}")
    ap = pool.dropna(subset=["breadcrumb_area", "price_numeric"]).groupby("breadcrumb_area")["price_numeric"].mean()
    ap = ap.sort_values(ascending=False).head(15).reset_index()
    ap["label"] = ap["price_numeric"].apply(fmt_price_short)
    st.plotly_chart(bar_with_labels(ap, "breadcrumb_area", "price_numeric", "Avg Price", None, orientation="h",
                                    text="label", height=450, color=ACCENT2), use_container_width=True)

    st.markdown("---")
    areas = ["All Areas"] + sorted(pool["breadcrumb_area"].dropna().unique())
    sel_area = st.selectbox("🔽 Drill into an Area", areas)
    area_pool = pool if sel_area == "All Areas" else pool[pool["breadcrumb_area"] == sel_area]

    st.subheader(f"⑤ Top Sub-areas in {sel_area}")
    sc = area_pool["breadcrumb_subarea"].dropna().value_counts().head(15).reset_index()
    sc.columns = ["subarea", "count"]
    if not sc.empty:
        sc = add_pct_labels(sc)
        st.plotly_chart(bar_with_labels(sc, "subarea", "count", "Listings", None, orientation="h",
                                        text="label", height=420), use_container_width=True)
    else:
        st.info("No sub-area breakdown available for this selection.")

    st.subheader(f"⑥ Average Price by Sub-area in {sel_area}")
    sp = area_pool.dropna(subset=["breadcrumb_subarea", "price_numeric"]).groupby("breadcrumb_subarea")["price_numeric"].mean()
    sp = sp.sort_values(ascending=False).head(15).reset_index()
    if not sp.empty:
        sp["label"] = sp["price_numeric"].apply(fmt_price_short)
        st.plotly_chart(bar_with_labels(sp, "breadcrumb_subarea", "price_numeric", "Avg Price", None, orientation="h",
                                        text="label", height=420, color=ACCENT2), use_container_width=True)
    else:
        st.info("No sub-area price data available for this selection.")

    st.markdown("---")
    st.subheader("⑦ Property Type Mix by City")
    if "type" in listings.columns:
        mix = listings.groupby(["city", "type"]).size().reset_index(name="count")
        fig = px.bar(mix, x="city", y="count", color="type", barmode="stack",
                    color_discrete_sequence=px.colors.qualitative.Set2, text="count")
        fig.update_traces(textposition="inside")
        fig.update_layout(template=CHART_TEMPLATE, xaxis_title=None, yaxis_title="Listings",
                          margin=dict(l=10, r=10, t=10, b=10), height=420)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("⑧ Full Breakdown Table")
    summary = pool.dropna(subset=["breadcrumb_area"]).groupby(
        ["breadcrumb_area", "breadcrumb_subarea"], dropna=False
    ).agg(listings=("listing_id", "count"), avg_price=("price_numeric", "mean"),
          avg_trend_pct=("trend_change_pct_overall", "mean")).reset_index().sort_values("listings", ascending=False)
    summary["avg_price"] = summary["avg_price"].apply(fmt_price)
    summary["avg_trend_pct"] = summary["avg_trend_pct"].round(1)
    st.dataframe(summary, use_container_width=True, hide_index=True, height=400)


# =============================================================================
# PAGE 4 — COMPARE AREAS
# =============================================================================
elif page == "⚖️ Compare Areas":
    st.header("⚖️ Compare Areas")

    sel_city = GLOBAL_CITY
    pool = GLOBAL_POOL

    all_areas = sorted(pool["breadcrumb_area"].dropna().unique())
    compare_areas = st.multiselect("Select areas to compare", all_areas,
                                    default=all_areas[:3] if len(all_areas) >= 3 else all_areas)

    if compare_areas:
        st.subheader("① Price Trend Comparison")
        fig_cmp = go.Figure()
        for i, area in enumerate(compare_areas):
            ids = pool[pool["breadcrumb_area"] == area]["trend_id"].dropna().unique()
            monthly = trend_line_for(ids, trend_history)
            if not monthly.empty:
                fig_cmp.add_trace(go.Scatter(x=monthly["month_year"], y=monthly["avg_price"],
                                             mode="lines", name=area,
                                             line=dict(width=2.5, color=PALETTE[i % len(PALETTE)])))
        fig_cmp.update_layout(template=CHART_TEMPLATE, height=440, yaxis_title="Average Price (PKR)",
                              legend=dict(orientation="h", yanchor="top", y=-0.15, x=0.5, xanchor="center"),
                              margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified")
        st.plotly_chart(fig_cmp, use_container_width=True)

        st.subheader("② Average Price Comparison")
        rows = [{"Area": a, "Avg Price": pool[pool["breadcrumb_area"] == a]["price_numeric"].mean(),
                 "Listings": len(pool[pool["breadcrumb_area"] == a])} for a in compare_areas]
        cmp_df = pd.DataFrame(rows)
        cmp_df["label"] = cmp_df["Avg Price"].apply(fmt_price_short)
        st.plotly_chart(bar_with_labels(cmp_df, "Area", "Avg Price", "Area", "Avg Price", text="label"),
                        use_container_width=True)

        st.subheader("③ Listings Count Comparison")
        cmp_df["pct"] = (cmp_df["Listings"] / cmp_df["Listings"].sum() * 100).round(1)
        cmp_df["label2"] = cmp_df.apply(lambda r: f"{int(r['Listings']):,} ({r['pct']}%)", axis=1)
        st.plotly_chart(bar_with_labels(cmp_df, "Area", "Listings", "Area", "Listings", text="label2", color=ACCENT2),
                        use_container_width=True)

        st.subheader("④ Full Comparison Table")
        cmp_df["Avg Price"] = cmp_df["Avg Price"].apply(fmt_price)
        st.dataframe(cmp_df[["Area", "Listings", "Avg Price"]], use_container_width=True, hide_index=True)
    else:
        st.info("Select at least one area to compare.")


# =============================================================================
# PAGE 5 — AREA RANKINGS (+ radar comparison)
# =============================================================================
elif page == "🏆 Area Rankings":
    st.header("🏆 Area Rankings")

    sel_city = GLOBAL_CITY
    pool = GLOBAL_POOL

    area_stats = pool.dropna(subset=["breadcrumb_area", "price_numeric"]).groupby("breadcrumb_area").agg(
        listings=("listing_id", "count"), avg_price=("price_numeric", "mean"),
        trend_pct=("trend_change_pct_overall", "mean"), avg_size=("area_size", "mean")).reset_index()

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("① 🔥 Hottest Areas (highest growth)")
        top = area_stats.sort_values("trend_pct", ascending=False).head(10).copy()
        top["avg_price"] = top["avg_price"].apply(fmt_price)
        st.dataframe(top[["breadcrumb_area", "listings", "avg_price", "trend_pct"]],
                    use_container_width=True, hide_index=True)
    with c2:
        st.subheader("② ❄️ Coolest Areas (lowest growth)")
        low = area_stats.sort_values("trend_pct", ascending=True).head(10).copy()
        low["avg_price"] = low["avg_price"].apply(fmt_price)
        st.dataframe(low[["breadcrumb_area", "listings", "avg_price", "trend_pct"]],
                    use_container_width=True, hide_index=True)

    st.subheader("③ 💎 Most Expensive Areas")
    exp = area_stats.sort_values("avg_price", ascending=False).head(15)
    exp = exp.assign(label=exp["avg_price"].apply(fmt_price_short))
    st.plotly_chart(bar_with_labels(exp, "breadcrumb_area", "avg_price", "Avg Price", None, orientation="h",
                                    text="label", height=450), use_container_width=True)

    st.subheader("④ 💰 Most Affordable Areas")
    aff = area_stats.sort_values("avg_price", ascending=True).head(15)
    aff = aff.assign(label=aff["avg_price"].apply(fmt_price_short))
    st.plotly_chart(bar_with_labels(aff, "breadcrumb_area", "avg_price", "Avg Price", None, orientation="h",
                                    text="label", height=450, color=ACCENT3), use_container_width=True)

    st.subheader("⑤ 📋 Most Listed Areas")
    listed = area_stats.sort_values("listings", ascending=False).head(15)
    listed = add_pct_labels(listed, "listings")
    st.plotly_chart(bar_with_labels(listed, "breadcrumb_area", "listings", "Listings", None, orientation="h",
                                    text="label", height=450, color=ACCENT2), use_container_width=True)

    # NEW: Radar comparison
    st.markdown("---")
    st.subheader("⑥ 🕸️ Multi-Dimensional Area Comparison (Radar)")
    radar_options = sorted(area_stats["breadcrumb_area"].dropna().unique())
    radar_areas = st.multiselect("Select 2-4 areas to compare across dimensions",
                                 radar_options, default=radar_options[:3] if len(radar_options) >= 3 else radar_options,
                                 key="radar_areas")
    if len(radar_areas) >= 2:
        radar_data = area_stats[area_stats["breadcrumb_area"].isin(radar_areas)].copy()
        radar_data["Affordability"] = 1 - normalize(radar_data["avg_price"])
        radar_data["Growth"] = normalize(radar_data["trend_pct"].fillna(0))
        radar_data["Popularity"] = normalize(radar_data["listings"])
        radar_data["Size"] = normalize(radar_data["avg_size"].fillna(0))

        dims = ["Affordability", "Growth", "Popularity", "Size"]
        fig_radar = go.Figure()
        for i, (_, row) in enumerate(radar_data.iterrows()):
            fig_radar.add_trace(go.Scatterpolar(
                r=[row[d] for d in dims] + [row[dims[0]]],
                theta=dims + [dims[0]],
                fill="toself", name=row["breadcrumb_area"],
                line=dict(color=PALETTE[i % len(PALETTE)]),
            ))
        fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 1], showticklabels=False)),
                                height=480, margin=dict(l=40, r=40, t=30, b=30),
                                legend=dict(orientation="h", yanchor="bottom", y=-0.15))
        st.plotly_chart(fig_radar, use_container_width=True)
        st.caption("Each axis is scaled 0-1 relative to the selected areas — bigger shape = better overall profile")
    else:
        st.info("Select at least 2 areas to compare.")


# =============================================================================
# PAGE 6 — INVESTMENT INSIGHTS
# =============================================================================
elif page == "💡 Investment Insights":
    st.header("💡 Investment Insights")
    st.caption("Where might the best opportunities be, based on price and growth trends")

    sel_city = GLOBAL_CITY
    pool = GLOBAL_POOL

    area_stats = pool.dropna(subset=["breadcrumb_area", "price_numeric", "trend_change_pct_overall"]).groupby("breadcrumb_area").agg(
        listings=("listing_id", "count"), avg_price=("price_numeric", "mean"),
        trend_pct=("trend_change_pct_overall", "mean"),
        avg_price_per_sqft=("trend_avg_price_per_sqft", "mean")).reset_index()
    area_stats = area_stats[area_stats["listings"] >= 3]

    if not area_stats.empty:
        price_median = area_stats["avg_price"].median()
        growth_median = area_stats["trend_pct"].median()

        def quadrant(row):
            if row["avg_price"] <= price_median and row["trend_pct"] >= growth_median:
                return "💎 Best Buy (cheap + growing)"
            if row["avg_price"] > price_median and row["trend_pct"] >= growth_median:
                return "🚀 Premium Growth"
            if row["avg_price"] <= price_median and row["trend_pct"] < growth_median:
                return "🏠 Budget / Stable"
            return "⚠️ Overpriced / Slow"

        area_stats["Category"] = area_stats.apply(quadrant, axis=1)

        st.subheader("① Opportunity Map — Price vs Growth")
        fig = px.scatter(area_stats, x="avg_price", y="trend_pct", size="listings", color="Category",
                         hover_name="breadcrumb_area",
                         color_discrete_map={
                             "💎 Best Buy (cheap + growing)": ACCENT,
                             "🚀 Premium Growth": ACCENT3,
                             "🏠 Budget / Stable": "#94a3b8",
                             "⚠️ Overpriced / Slow": ACCENT2,
                         }, size_max=40)
        fig.add_hline(y=growth_median, line_dash="dot", line_color="gray")
        fig.add_vline(x=price_median, line_dash="dot", line_color="gray")
        fig.update_layout(template=CHART_TEMPLATE, xaxis_title="Average Price (PKR)",
                          yaxis_title="Price Growth %", margin=dict(l=10, r=10, t=10, b=10), height=480)
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Bubble size = number of listings · Dotted lines = median price / median growth")

        st.subheader("② 💎 Best Buy Areas (below-median price, above-median growth)")
        best_buys = area_stats[area_stats["Category"] == "💎 Best Buy (cheap + growing)"].sort_values("trend_pct", ascending=False)
        best_buys_display = best_buys.copy()
        best_buys_display["avg_price"] = best_buys_display["avg_price"].apply(fmt_price)
        st.dataframe(best_buys_display[["breadcrumb_area", "listings", "avg_price", "trend_pct"]],
                    use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("③ 📐 Cheapest Price/Sqft Areas")
            if area_stats["avg_price_per_sqft"].notna().any():
                cheap_sqft = area_stats.dropna(subset=["avg_price_per_sqft"]).sort_values("avg_price_per_sqft").head(10)
                cheap_sqft = cheap_sqft.assign(label=cheap_sqft["avg_price_per_sqft"].round(0).astype(int).astype(str))
                st.plotly_chart(bar_with_labels(cheap_sqft, "breadcrumb_area", "avg_price_per_sqft",
                                                "PKR/Sqft", None, orientation="h", text="label", height=400),
                                use_container_width=True)
            else:
                st.info("Price/sqft data not available for this scope.")
        with c2:
            st.subheader("④ 📐 Priciest Price/Sqft Areas")
            if area_stats["avg_price_per_sqft"].notna().any():
                exp_sqft = area_stats.dropna(subset=["avg_price_per_sqft"]).sort_values("avg_price_per_sqft", ascending=False).head(10)
                exp_sqft = exp_sqft.assign(label=exp_sqft["avg_price_per_sqft"].round(0).astype(int).astype(str))
                st.plotly_chart(bar_with_labels(exp_sqft, "breadcrumb_area", "avg_price_per_sqft",
                                                "PKR/Sqft", None, orientation="h", text="label", height=400, color=ACCENT2),
                                use_container_width=True)

        st.subheader("⑤ 🏆 ROI Potential Score (growth-weighted ranking)")
        area_stats["roi_score"] = (normalize(area_stats["trend_pct"]) * 0.7 +
                                    (1 - normalize(area_stats["avg_price"])) * 0.3) * 100
        roi_ranked = area_stats.sort_values("roi_score", ascending=False).head(15)
        roi_display = roi_ranked.copy()
        roi_display["avg_price"] = roi_display["avg_price"].apply(fmt_price)
        roi_display["roi_score"] = roi_display["roi_score"].round(1)
        st.dataframe(roi_display[["breadcrumb_area", "avg_price", "trend_pct", "roi_score"]]
                    .rename(columns={"roi_score": "ROI Score (0-100)"}),
                    use_container_width=True, hide_index=True)
        st.caption("ROI Score blends growth trend (70% weight) and affordability (30% weight) — higher is better")
    else:
        st.info("Not enough data with trend information for this scope.")


# =============================================================================
# PAGE 7 — AMENITIES IMPACT
# =============================================================================
elif page == "🛋️ Amenities Impact":
    st.header("🛋️ Amenities Impact")
    st.caption("Which features actually correlate with higher prices?")

    sel_city = GLOBAL_CITY
    pool = GLOBAL_POOL

    st.subheader("① How Common Is Each Amenity?")
    freq_rows = []
    for kw in AMENITY_KEYWORDS:
        col = f"has_{kw}"
        if col in pool.columns:
            freq_rows.append({"Amenity": kw, "count": pool[col].sum()})
    freq_df = pd.DataFrame(freq_rows).sort_values("count", ascending=False)
    freq_df = add_pct_labels(freq_df)
    st.plotly_chart(bar_with_labels(freq_df, "Amenity", "count", "Listings", None, orientation="h",
                                    text="label", height=420), use_container_width=True)

    st.subheader("② Price Impact — With vs Without Each Amenity")
    impact_rows = []
    for kw in AMENITY_KEYWORDS:
        col = f"has_{kw}"
        if col in pool.columns:
            with_amenity = pool[pool[col]]["price_numeric"].mean()
            without_amenity = pool[~pool[col]]["price_numeric"].mean()
            if pd.notna(with_amenity) and pd.notna(without_amenity):
                impact_rows.append({"Amenity": kw, "With": with_amenity, "Without": without_amenity})
    impact_df = pd.DataFrame(impact_rows)
    if not impact_df.empty:
        impact_df["Premium %"] = ((impact_df["With"] - impact_df["Without"]) / impact_df["Without"] * 100).round(1)
        impact_df = impact_df.sort_values("Premium %", ascending=False)
        melted = impact_df.melt(id_vars="Amenity", value_vars=["With", "Without"], var_name="Has Amenity", value_name="Avg Price")
        fig = px.bar(melted, x="Amenity", y="Avg Price", color="Has Amenity", barmode="group",
                    color_discrete_sequence=[ACCENT, "#94a3b8"])
        fig.update_layout(template=CHART_TEMPLATE, xaxis_title=None, yaxis_title="Average Price (PKR)",
                          margin=dict(l=10, r=10, t=10, b=10), height=440)
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("③ Which Amenities Add the Most Price Premium?")
        prem = impact_df.assign(label=impact_df["Premium %"].apply(lambda v: f"{v:+.1f}%"))
        st.plotly_chart(bar_with_labels(prem, "Amenity", "Premium %", "Price Premium %", None,
                                        orientation="h", text="label", height=420,
                                        color=ACCENT), use_container_width=True)

    st.subheader("④ More Amenities = Higher Price?")
    if "amenity_count" in pool.columns:
        amenity_price = pool.dropna(subset=["amenity_count", "price_numeric"]).groupby("amenity_count")["price_numeric"].mean().reset_index()
        amenity_price["label"] = amenity_price["price_numeric"].apply(fmt_price_short)
        st.plotly_chart(bar_with_labels(amenity_price, "amenity_count", "price_numeric",
                                        "Number of Amenities", "Avg Price", text="label"), use_container_width=True)

    st.subheader("⑤ Listings Distribution by Number of Amenities")
    if "amenity_count" in pool.columns:
        ac = pool["amenity_count"].value_counts().sort_index().reset_index()
        ac.columns = ["amenity_count", "count"]
        ac = add_pct_labels(ac)
        st.plotly_chart(bar_with_labels(ac, "amenity_count", "count", "Number of Amenities", "Listings",
                                        text="label", color=ACCENT2), use_container_width=True)

    st.markdown("---")
    st.subheader("🛏️ Room & Space Features")
    st.caption("Kitchens, parking, servant quarters, and other room-level details from the listings")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⑥ Parking Spaces — How Many Listings Have Them")
        if "parking_spaces_count" in pool.columns:
            pk = pool.dropna(subset=["parking_spaces_count"])
            pk_dist = pk["parking_spaces_count"].astype(int).value_counts().sort_index().head(8).reset_index()
            pk_dist.columns = ["parking_spaces", "count"]
            pk_dist = add_pct_labels(pk_dist)
            st.plotly_chart(bar_with_labels(pk_dist, "parking_spaces", "count", "Parking Spaces", "Listings",
                                            text="label"), use_container_width=True)
        else:
            st.info("No parking space data available.")

    with c2:
        st.subheader("⑦ Number of Kitchens")
        if "kitchens_count" in pool.columns:
            kc = pool.dropna(subset=["kitchens_count"])
            kc_dist = kc["kitchens_count"].astype(int).value_counts().sort_index().reset_index()
            kc_dist.columns = ["kitchens", "count"]
            kc_dist = add_pct_labels(kc_dist)
            st.plotly_chart(bar_with_labels(kc_dist, "kitchens", "count", "Kitchens", "Listings",
                                            text="label", color=ACCENT2), use_container_width=True)
        else:
            st.info("No kitchen count data available.")

    st.subheader("⑧ Presence of Key Room Features")
    room_flags = ["Drawing Room", "Dining Room", "Servant Quarters", "Store Rooms",
                  "Study Room", "Prayer Room", "Powder Room", "Lounge or Sitting Room", "Laundry Room"]
    rows = []
    for feature in room_flags:
        col = f"has_{feature}"
        if col in pool.columns:
            rows.append({"Feature": feature, "count": pool[col].sum()})
    if rows:
        rf_df = pd.DataFrame(rows).sort_values("count", ascending=False)
        rf_df = add_pct_labels(rf_df)
        st.plotly_chart(bar_with_labels(rf_df, "Feature", "count", "Listings", None, orientation="h",
                                        text="label", height=420), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⑨ Price Impact — Servant Quarters")
        if "has_Servant Quarters" in pool.columns:
            with_sq = pool[pool["has_Servant Quarters"]]["price_numeric"].mean()
            without_sq = pool[~pool["has_Servant Quarters"]]["price_numeric"].mean()
            comp_df = pd.DataFrame({"Has Servant Quarters": ["Yes", "No"], "price": [with_sq, without_sq]})
            comp_df["label"] = comp_df["price"].apply(fmt_price_short)
            st.plotly_chart(bar_with_labels(comp_df, "Has Servant Quarters", "price", "Servant Quarters", "Avg Price",
                                            text="label"), use_container_width=True)

    with c2:
        st.subheader("⑩ Built Year Distribution")
        if "built_year" in pool.columns and pool["built_year"].notna().any():
            by = pool.dropna(subset=["built_year"])
            by = by[(by["built_year"] >= 1990) & (by["built_year"] <= 2030)]
            by["decade"] = (by["built_year"] // 5 * 5).astype(int).astype(str) + "s"
            by_dist = by["decade"].value_counts().sort_index().reset_index()
            by_dist.columns = ["period", "count"]
            by_dist = add_pct_labels(by_dist)
            st.plotly_chart(bar_with_labels(by_dist, "period", "count", "Built Period", "Listings", text="label"),
                            use_container_width=True)
        else:
            st.info("No built-year data available.")


# =============================================================================
# PAGE 8 — MARKET ACTIVITY
# =============================================================================
elif page == "⏱️ Market Activity":
    st.header("⏱️ Market Activity")
    st.caption("How fresh is the current inventory, and where is new supply concentrated")

    sel_city = GLOBAL_CITY
    pool = GLOBAL_POOL

    c1, c2, c3 = st.columns(3)
    fresh_pct = (pool["days_ago"] <= 7).mean() * 100 if "days_ago" in pool.columns else None
    c1.metric("Listed in Last 7 Days", f"{fresh_pct:.1f}%" if fresh_pct is not None else "-")
    c2.metric("Median Listing Age", f"{pool['days_ago'].median():.0f} days" if "days_ago" in pool.columns else "-")
    c3.metric("Total Listings", f"{len(pool):,}")

    st.write("")
    st.subheader("① Listing Freshness Distribution")
    if "freshness" in pool.columns:
        fc = pool["freshness"].value_counts().reindex(FRESHNESS_ORDER).dropna().reset_index()
        fc.columns = ["freshness", "count"]
        fc = add_pct_labels(fc)
        st.plotly_chart(bar_with_labels(fc, "freshness", "count", "Listing Age", "Listings", text="label"),
                        use_container_width=True)

    st.subheader("② Freshness Split (Pie)")
    if "freshness" in pool.columns:
        fp = pool["freshness"].value_counts().reindex(FRESHNESS_ORDER).dropna().reset_index()
        fp.columns = ["freshness", "count"]
        st.plotly_chart(pie_with_labels(fp, "freshness", "count"), use_container_width=True)

    st.subheader("③ 🔥 Most Active Areas (highest % fresh listings, min 5 listings)")
    if "freshness" in pool.columns and "breadcrumb_area" in pool.columns:
        area_activity = pool.dropna(subset=["breadcrumb_area"]).groupby("breadcrumb_area").agg(
            listings=("listing_id", "count"),
            fresh_count=("days_ago", lambda s: (s <= 7).sum())
        ).reset_index()
        area_activity = area_activity[area_activity["listings"] >= 5]
        area_activity["fresh_pct"] = (area_activity["fresh_count"] / area_activity["listings"] * 100).round(1)
        top_active = area_activity.sort_values("fresh_pct", ascending=False).head(15)
        top_active = top_active.assign(label=top_active["fresh_pct"].apply(lambda v: f"{v}%"))
        st.plotly_chart(bar_with_labels(top_active, "breadcrumb_area", "fresh_pct", "% Fresh Listings", None,
                                        orientation="h", text="label", height=450), use_container_width=True)

    st.subheader("④ Freshness by Property Type")
    if "freshness" in pool.columns and "type" in pool.columns:
        mix = pool.dropna(subset=["freshness"]).groupby(["type", "freshness"]).size().reset_index(name="count")
        fig = px.bar(mix, x="type", y="count", color="freshness", barmode="stack",
                    category_orders={"freshness": FRESHNESS_ORDER},
                    color_discrete_sequence=px.colors.sequential.Greens_r, text="count")
        fig.update_traces(textposition="inside")
        fig.update_layout(template=CHART_TEMPLATE, xaxis_title=None, yaxis_title="Listings",
                          margin=dict(l=10, r=10, t=10, b=10), height=420)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("⑤ Freshest Listings Right Now")
    if "days_ago" in pool.columns:
        newest = pool.sort_values("days_ago", na_position="last").head(15)
        cols = [c for c in ["title", "breadcrumb_area", "price_raw", "added_raw", "url"] if c in newest.columns]
        st.dataframe(newest[cols], use_container_width=True, hide_index=True,
                    column_config={"url": st.column_config.LinkColumn("Link")})


# =============================================================================
# PAGE 10 — BROWSE LISTINGS
# =============================================================================
elif page == "🔍 Browse Listings":
    st.header("🔍 Browse Listings")

    c1, c2 = st.columns(2)
    sel_city = GLOBAL_CITY
    table_data = GLOBAL_POOL

    with c1:
        types = sorted(table_data["type"].dropna().unique())
        sel_types = st.multiselect("Property Type", types, default=types)
        if sel_types:
            table_data = table_data[table_data["type"].isin(sel_types)]

    with c2:
        sort_choice = st.selectbox("Sort by", ["Newest", "Price: High to Low", "Price: Low to High"])

    search_term = st.text_input("🔎 Search by title or description")
    if search_term:
        mask = (table_data["title"].str.contains(search_term, case=False, na=False) |
                table_data["description"].str.contains(search_term, case=False, na=False))
        table_data = table_data[mask]

    if sort_choice == "Price: High to Low":
        table_data = table_data.sort_values("price_numeric", ascending=False, na_position="last")
    elif sort_choice == "Price: Low to High":
        table_data = table_data.sort_values("price_numeric", ascending=True, na_position="last")
    elif "days_ago" in table_data.columns:
        table_data = table_data.sort_values("days_ago", ascending=True, na_position="last")

    display_cols = [c for c in [
        "title", "city", "breadcrumb_area", "breadcrumb_subarea", "type",
        "price_raw", "bedrooms", "bathrooms", "area_raw", "agency_name", "url"
    ] if c in table_data.columns]

    st.dataframe(table_data[display_cols], use_container_width=True, hide_index=True, height=520,
                column_config={"url": st.column_config.LinkColumn("Link")})
    st.caption(f"Showing {len(table_data):,} of {len(listings):,} total listings")