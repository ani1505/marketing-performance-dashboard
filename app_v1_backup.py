from __future__ import annotations

import os
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from data_store import COMPANY_CURRENCY_RULES, WorkbookStore


# -----------------------------------------------------------------------------
# App configuration
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Marketing Performance Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = Path(
    os.getenv("DASHBOARD_WORKBOOK", str(BASE_DIR / "wolfable_master.xlsx"))
)
STORE = WorkbookStore(WORKBOOK_PATH)

COLORS = {
    "navy": "#17324D",
    "blue": "#3F7CAC",
    "blue2": "#6AA9D6",
    "sky": "#EAF4FB",
    "red": "#E63946",
    "text": "#263445",
    "muted": "#6B7A8C",
    "line": "#DCE5EC",
    "panel": "#F5F7F9",
    "white": "#FFFFFF",
}

PLOT_COLORS = ["#3F7CAC", "#E63946", "#71B48D", "#F2B134", "#8C6BB1"]


# -----------------------------------------------------------------------------
# Styling
# -----------------------------------------------------------------------------
st.markdown(
    """
<style>
:root {
  --bg: #F8FAFC;
  --panel: #FFFFFF;
  --soft: #F1F5F9;
  --line: #E2E8F0;
  --text: #25364A;
  --muted: #718096;
  --blue: #3F7CAC;
  --red: #E63946;
}

html, body, [class*="css"] { font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
.stApp { background: var(--bg); color: var(--text); }
.block-container { padding-top: 1.35rem; padding-bottom: 2rem; max-width: 1560px; }

/* Header */
.hero {
  padding: 1.15rem 1.3rem;
  background: linear-gradient(110deg, #FFFFFF 0%, #EEF6FC 100%);
  border: 1px solid var(--line);
  border-radius: 18px;
  margin-bottom: 0.9rem;
  box-shadow: 0 8px 26px rgba(35, 58, 86, 0.06);
}
.hero h1 { margin: 0; font-size: 1.75rem; color: #17324D; letter-spacing: -0.02em; }
.hero p { margin: 0.25rem 0 0; color: #718096; font-size: 0.93rem; }

/* Panels */
.panel-title { font-weight: 750; color: #27394E; font-size: 0.98rem; margin: 0 0 0.35rem; }
.panel-subtitle { color: #7B8796; font-size: 0.79rem; margin-bottom: 0.4rem; }
[data-testid="stVerticalBlockBorderWrapper"] {
  border-radius: 16px !important;
  border-color: #E1E8EF !important;
  background: #FFFFFF;
  box-shadow: 0 4px 20px rgba(35, 58, 86, 0.04);
}

/* KPI Cards */
.kpi-card {
  min-height: 118px;
  background: linear-gradient(145deg, #FFFFFF 0%, #F2F8FC 100%);
  border: 1px solid #DCE7EF;
  border-top: 4px solid #3F7CAC;
  border-radius: 16px;
  padding: 1rem 1.05rem;
  box-shadow: 0 5px 18px rgba(36, 62, 88, 0.05);
  display: flex;
  flex-direction: column;
  justify-content: space-between;
}
.kpi-label { color: #6C7B8C; font-size: 0.78rem; font-weight: 700; letter-spacing: 0.035em; text-transform: uppercase; }
.kpi-value { color: #17324D; font-size: 1.52rem; line-height: 1.14; font-weight: 800; letter-spacing: -0.025em; overflow-wrap: anywhere; }
.kpi-value.currency { font-size: 1.15rem; line-height: 1.28; }
.kpi-foot { color: #8A98A8; font-size: 0.72rem; margin-top: 0.35rem; }

/* Filter section */
.filter-strip {
  background: #FFFFFF;
  border: 1px solid #E1E8EF;
  border-radius: 16px;
  padding: 0.2rem 0.75rem 0.55rem;
  margin-bottom: 0.9rem;
  box-shadow: 0 4px 16px rgba(35, 58, 86, 0.035);
}
div[data-baseweb="select"] > div { border-radius: 10px !important; }
.stDateInput input { border-radius: 10px !important; }
.stButton > button, .stDownloadButton > button {
  border-radius: 10px;
  font-weight: 700;
  border: 1px solid #D8E2EA;
}
.stButton > button[kind="primary"] {
  background: #3F7CAC;
  border-color: #3F7CAC;
}

/* Dataframes */
[data-testid="stDataFrame"] { border: 1px solid #E1E8EF; border-radius: 12px; overflow: hidden; }

/* Sidebar */
[data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid #E6EDF2; }

.currency-note {
  display: inline-block;
  padding: 0.42rem 0.7rem;
  background: #FFF7ED;
  border: 1px solid #FED7AA;
  border-radius: 10px;
  color: #9A5A13;
  font-size: 0.78rem;
}
.success-note {
  padding: 0.75rem 0.9rem;
  background: #ECFDF5;
  border: 1px solid #A7F3D0;
  color: #166534;
  border-radius: 12px;
  font-size: 0.85rem;
}
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False, ttl=20)
def load_data(path_str: str, modified_time: float):
    # modified_time is intentionally part of the cache key.
    store = WorkbookStore(path_str)
    return store.load_dashboard_data()


def get_data():
    STORE.ensure_exists()
    return load_data(str(STORE.path), STORE.path.stat().st_mtime)


def clear_data_cache():
    load_data.clear()


def human_number(value: float | int) -> str:
    value = float(value or 0)
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.2f}M"
    if abs(value) >= 1_000:
        return f"{value / 1_000:.1f}K"
    return f"{value:,.0f}"


def pct(value: float) -> str:
    return f"{value * 100:.2f}%" if pd.notna(value) else "—"


def currency_breakdown(df: pd.DataFrame, symbols: dict[str, str], metric: str = "spend") -> str:
    if df.empty:
        return "—"
    parts = []
    for currency, group in df.groupby("Currency", dropna=False):
        currency = str(currency)
        symbol = symbols.get(currency, "")
        spend = pd.to_numeric(group["Total Spend"], errors="coerce").fillna(0).sum()
        clicks = pd.to_numeric(group["Clicks"], errors="coerce").fillna(0).sum()
        leads = pd.to_numeric(group["Total Leads"], errors="coerce").fillna(0).sum()
        if metric == "spend":
            value = spend
        elif metric == "cpl":
            value = spend / leads if leads else 0
        elif metric == "cpc":
            value = spend / clicks if clicks else 0
        else:
            raise ValueError(metric)
        parts.append(f"{symbol}{value:,.2f} {currency}")
    return "  ·  ".join(parts)


def metric_card(label: str, value: str, footer: str = "", currency: bool = False):
    value_class = "kpi-value currency" if currency else "kpi-value"
    st.markdown(
        f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="{value_class}">{value}</div>
          <div class="kpi-foot">{footer}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def style_plot(fig, height=340, legend=True):
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=20, r=18, t=46, b=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Inter, Arial", color="#425466", size=12),
        title_font=dict(size=15, color="#26384C"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        showlegend=legend,
        hoverlabel=dict(bgcolor="white", font_size=12),
    )
    fig.update_xaxes(showgrid=False, linecolor="#E4EAF0", tickfont=dict(color="#718096"))
    fig.update_yaxes(gridcolor="#EEF2F6", zeroline=False, tickfont=dict(color="#718096"))
    return fig


def pills(label: str, options: list[str], default, key: str, selection_mode: str = "single"):
    """Use st.pills when available; fallback to selectbox/multiselect."""
    if hasattr(st, "pills"):
        return st.pills(label, options, default=default, key=key, selection_mode=selection_mode)
    if selection_mode == "multi":
        return st.multiselect(label, options, default=default, key=key)
    index = options.index(default) if default in options else 0
    return st.selectbox(label, options, index=index, key=key)


def filter_date(df: pd.DataFrame, period: str, custom_range=None) -> pd.DataFrame:
    if df.empty or "Report Date" not in df.columns:
        return df
    out = df.copy()
    dates = pd.to_datetime(out["Report Date"], errors="coerce").dt.date
    today = date.today()

    if period == "Yesterday":
        target = today - timedelta(days=1)
        return out[dates == target]
    if period == "Last Week":
        this_monday = today - timedelta(days=today.weekday())
        start = this_monday - timedelta(days=7)
        end = this_monday - timedelta(days=1)
        return out[(dates >= start) & (dates <= end)]
    if period == "Custom" and custom_range:
        start, end = custom_range
        return out[(dates >= start) & (dates <= end)]
    return out


def available_currency_options(company: str, master: pd.DataFrame, currency_df: pd.DataFrame) -> list[str]:
    all_codes = [str(v).strip() for v in currency_df.get("Currency", pd.Series(dtype=str)).dropna().tolist()]
    if company != "All Clients":
        return COMPANY_CURRENCY_RULES.get(company, all_codes)
    present = [str(v).strip() for v in master.get("Currency", pd.Series(dtype=str)).dropna().unique().tolist()]
    return [c for c in all_codes if c in present] or all_codes


def primary_currency_text(company: str, companies: pd.DataFrame, symbols: dict[str, str]) -> str:
    if company == "All Clients":
        return "Multiple currencies — monetary values are always separated"
    allowed = COMPANY_CURRENCY_RULES.get(company)
    if allowed:
        return " · ".join(f"{symbols.get(c, '')} {c}".strip() for c in allowed)
    match = companies[companies["Company Name"].astype(str) == company]
    if match.empty:
        return ""
    c = str(match.iloc[0].get("Native Currency", ""))
    return f"{symbols.get(c, '')} {c}".strip()


# -----------------------------------------------------------------------------
# Dashboard
# -----------------------------------------------------------------------------
def render_dashboard():
    master, companies, currency_df = get_data()
    symbols = {
        str(row["Currency"]).strip(): str(row["Symbol"]).strip()
        for _, row in currency_df.iterrows()
        if pd.notna(row.get("Currency"))
    }

    st.markdown(
        """
        <div class="hero">
          <h1>Marketing Performance Hub</h1>
          <p>Cross-client campaign performance with currency-safe financial reporting.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    company_names = companies["Company Name"].dropna().astype(str).tolist()

    st.markdown('<div class="filter-strip">', unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns([1.25, 1.1, 1.05, 1.55], gap="medium")
    with f1:
        company = st.selectbox("Client", ["All Clients"] + company_names, index=0)

    company_scope = master.copy()
    if company != "All Clients":
        company_scope = company_scope[company_scope["Company"].astype(str) == company]

    platform_options = sorted(company_scope["Platform"].dropna().astype(str).unique().tolist())
    region_options = sorted(company_scope["Region"].dropna().astype(str).unique().tolist())

    with f2:
        platform = st.selectbox("Platform", ["All Platforms"] + platform_options)
    with f3:
        region = st.selectbox("Region", ["All Regions"] + region_options)
    with f4:
        period = pills(
            "Date",
            ["Yesterday", "Last Week", "Custom", "Lifetime"],
            default="Lifetime",
            key="date_period",
            selection_mode="single",
        ) or "Lifetime"

    custom_range = None
    if period == "Custom":
        c1, c2 = st.columns([3, 2])
        with c1:
            picked = st.date_input(
                "Custom date range",
                value=(date.today() - timedelta(days=30), date.today()),
                format="DD/MM/YYYY",
            )
            if isinstance(picked, (tuple, list)) and len(picked) == 2:
                custom_range = (picked[0], picked[1])
        with c2:
            st.caption("The dashboard filters MasterData using **Report Date**.")

    currency_options = available_currency_options(company, master, currency_df)
    currency_default = currency_options.copy()
    selected_currencies = pills(
        "Currency",
        currency_options,
        default=currency_default,
        key=f"currency_{company}",
        selection_mode="multi",
    )
    if not selected_currencies:
        selected_currencies = currency_options
    st.markdown("</div>", unsafe_allow_html=True)

    filtered = master.copy()
    if company != "All Clients":
        filtered = filtered[filtered["Company"].astype(str) == company]
    if platform != "All Platforms":
        filtered = filtered[filtered["Platform"].astype(str) == platform]
    if region != "All Regions":
        filtered = filtered[filtered["Region"].astype(str) == region]
    if selected_currencies:
        filtered = filtered[filtered["Currency"].astype(str).isin(selected_currencies)]
    filtered = filter_date(filtered, period, custom_range)

    currency_text = primary_currency_text(company, companies, symbols)
    st.markdown(
        f'<span class="currency-note">Primary / allowed currency: <b>{currency_text}</b>. No currency conversion is performed.</span>',
        unsafe_allow_html=True,
    )
    st.write("")

    if filtered.empty:
        st.warning("No rows match the current filters. Try Lifetime or a different currency/date selection.")
        return

    impressions = pd.to_numeric(filtered["Impressions"], errors="coerce").fillna(0).sum()
    clicks = pd.to_numeric(filtered["Clicks"], errors="coerce").fillna(0).sum()
    leads = pd.to_numeric(filtered["Total Leads"], errors="coerce").fillna(0).sum()
    ctr = clicks / impressions if impressions else 0
    conversion = leads / clicks if clicks else 0
    spend_text = currency_breakdown(filtered, symbols, "spend")
    cpl_text = currency_breakdown(filtered, symbols, "cpl")
    cpc_text = currency_breakdown(filtered, symbols, "cpc")

    # Eight metric cards: 2 rows x 4.
    cards = [
        ("Total Spend", spend_text, "Separated by transaction currency", True),
        ("Total Leads", human_number(leads), "Lead volume", False),
        ("Cost per Lead", cpl_text, "Spend ÷ leads, per currency", True),
        ("CTR", pct(ctr), "Clicks ÷ impressions", False),
        ("Impressions", human_number(impressions), "Ad impressions", False),
        ("Clicks", human_number(clicks), "Ad clicks", False),
        ("Avg. CPC", cpc_text, "Spend ÷ clicks, per currency", True),
        ("Conversion Rate", pct(conversion), "Leads ÷ clicks", False),
    ]
    for start in (0, 4):
        cols = st.columns(4, gap="medium")
        for col, item in zip(cols, cards[start : start + 4]):
            with col:
                metric_card(*item)
        st.write("")

    # Analysis row: main chart + currency table.
    left, right = st.columns([3.2, 1.05], gap="medium")
    with left:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Spend by platform and currency</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-subtitle">Currencies use separate facets and are never summed together.</div>', unsafe_allow_html=True)
            spend_platform = (
                filtered.groupby(["Currency", "Platform"], as_index=False)["Total Spend"].sum()
            )
            fig = px.bar(
                spend_platform,
                x="Platform",
                y="Total Spend",
                color="Platform",
                facet_col="Currency" if spend_platform["Currency"].nunique() > 1 else None,
                color_discrete_sequence=PLOT_COLORS,
                labels={"Total Spend": "Spend"},
            )
            fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
            fig.update_yaxes(matches=None)
            fig = style_plot(fig, height=360, legend=False)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with right:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Currency summary</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-subtitle">Every monetary metric stays in its native currency.</div>', unsafe_allow_html=True)
            summary_rows = []
            for cur, g in filtered.groupby("Currency"):
                spend = pd.to_numeric(g["Total Spend"], errors="coerce").fillna(0).sum()
                leads_cur = pd.to_numeric(g["Total Leads"], errors="coerce").fillna(0).sum()
                clicks_cur = pd.to_numeric(g["Clicks"], errors="coerce").fillna(0).sum()
                summary_rows.append(
                    {
                        "Currency": cur,
                        "Spend": f"{symbols.get(str(cur), '')}{spend:,.2f}",
                        "Leads": int(leads_cur),
                        "CPL": f"{symbols.get(str(cur), '')}{(spend / leads_cur if leads_cur else 0):,.2f}",
                        "CPC": f"{symbols.get(str(cur), '')}{(spend / clicks_cur if clicks_cur else 0):,.2f}",
                    }
                )
            st.dataframe(pd.DataFrame(summary_rows), hide_index=True, use_container_width=True, height=270)

    # Bottom row of analysis blocks.
    b1, b2, b3 = st.columns([1.55, 1.0, 1.25], gap="medium")
    with b1:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Spend trend</div>', unsafe_allow_html=True)
            trend = filtered.dropna(subset=["Report Date"]).copy()
            trend["Report Date"] = pd.to_datetime(trend["Report Date"]).dt.date
            trend = trend.groupby(["Report Date", "Currency"], as_index=False)["Total Spend"].sum()
            if trend.empty:
                st.info("No dated rows in this selection.")
            else:
                fig = px.line(
                    trend,
                    x="Report Date",
                    y="Total Spend",
                    color="Currency",
                    markers=True,
                    facet_row="Currency" if trend["Currency"].nunique() > 1 else None,
                    color_discrete_sequence=PLOT_COLORS,
                )
                fig.update_yaxes(matches=None)
                fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
                fig = style_plot(fig, height=350, legend=trend["Currency"].nunique() == 1)
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with b2:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Leads by platform</div>', unsafe_allow_html=True)
            lead_platform = filtered.groupby("Platform", as_index=False)["Total Leads"].sum().sort_values("Total Leads")
            fig = px.bar(
                lead_platform,
                x="Total Leads",
                y="Platform",
                orientation="h",
                color="Platform",
                color_discrete_sequence=PLOT_COLORS,
            )
            fig = style_plot(fig, height=350, legend=False)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with b3:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Campaign performance</div>', unsafe_allow_html=True)
            table_cols = [
                "Company",
                "Campaign Name",
                "Platform",
                "Region",
                "Campaign Status",
                "Currency",
                "Total Spend",
                "Total Leads",
                "Clicks",
            ]
            detail = filtered[[c for c in table_cols if c in filtered.columns]].copy()
            if "Total Spend" in detail:
                detail["Total Spend"] = pd.to_numeric(detail["Total Spend"], errors="coerce").round(2)
            st.dataframe(detail, hide_index=True, use_container_width=True, height=305)

    st.caption(
        f"Source: {STORE.path.name} · {len(filtered):,} filtered campaign rows · "
        "Financial metrics are grouped by currency without FX conversion."
    )


# -----------------------------------------------------------------------------
# Data Manager
# -----------------------------------------------------------------------------
def record_form(prefix: str, company: str, defaults: dict | None = None):
    defaults = defaults or {}
    allowed = STORE.allowed_currencies(company)
    platforms_default = ["Google Ads", "Meta Ads", "LinkedIn Ads", "YouTube Ads"]
    current_platform = str(defaults.get("Platform") or "Google Ads")
    platforms = list(dict.fromkeys(platforms_default + [current_platform]))

    col1, col2, col3 = st.columns(3, gap="medium")
    with col1:
        platform = st.selectbox(
            "Platform *",
            platforms,
            index=platforms.index(current_platform) if current_platform in platforms else 0,
            key=f"{prefix}_platform",
        )
        campaign = st.text_input(
            "Campaign Name *",
            value=str(defaults.get("Campaign Name") or ""),
            key=f"{prefix}_campaign",
        )
        status_value = str(defaults.get("Campaign Status") or "Active")
        status = st.selectbox(
            "Campaign Status *",
            ["Active", "Inactive", "Paused"],
            index=["Active", "Inactive", "Paused"].index(status_value)
            if status_value in ["Active", "Inactive", "Paused"]
            else 0,
            key=f"{prefix}_status",
        )
        region = st.text_input(
            "Region *",
            value=str(defaults.get("Region") or ""),
            key=f"{prefix}_region",
        )

    def to_date(value, fallback):
        parsed = pd.to_datetime(value, errors="coerce")
        return fallback if pd.isna(parsed) else parsed.date()

    with col2:
        report_date = st.date_input(
            "Report Date *",
            value=to_date(defaults.get("Report Date"), date.today()),
            key=f"{prefix}_report_date",
        )
        start_date = st.date_input(
            "Start Date *",
            value=to_date(defaults.get("Start Date"), date.today()),
            key=f"{prefix}_start_date",
        )
        current_end = pd.to_datetime(defaults.get("End Date"), errors="coerce")
        no_end_default = bool(pd.isna(current_end))
        no_end = st.checkbox(
            "No end date / ongoing",
            value=no_end_default,
            key=f"{prefix}_no_end",
        )
        end_date = None
        if not no_end:
            end_date = st.date_input(
                "End Date",
                value=date.today() if pd.isna(current_end) else current_end.date(),
                key=f"{prefix}_end_date",
            )
        current_currency = str(defaults.get("Currency") or allowed[0])
        if current_currency not in allowed:
            current_currency = allowed[0]
        currency = st.selectbox(
            "Currency *",
            allowed,
            index=allowed.index(current_currency),
            key=f"{prefix}_currency",
            help="Currency is restricted by company. No FX conversion is performed.",
        )
        daily_budget = st.number_input(
            "Daily Budget",
            min_value=0.0,
            value=float(defaults.get("Daily Budget") or 0.0),
            step=1.0,
            format="%.2f",
            key=f"{prefix}_budget",
        )

    with col3:
        impressions = st.number_input(
            "Impressions",
            min_value=0,
            value=int(float(defaults.get("Impressions") or 0)),
            step=1,
            key=f"{prefix}_impressions",
        )
        clicks = st.number_input(
            "Clicks",
            min_value=0,
            value=int(float(defaults.get("Clicks") or 0)),
            step=1,
            key=f"{prefix}_clicks",
        )
        leads = st.number_input(
            "Total Leads",
            min_value=0,
            value=int(float(defaults.get("Total Leads") or 0)),
            step=1,
            key=f"{prefix}_leads",
        )
        spend = st.number_input(
            "Total Spend *",
            min_value=0.0,
            value=float(defaults.get("Total Spend") or 0.0),
            step=1.0,
            format="%.2f",
            key=f"{prefix}_spend",
        )
        notes = st.text_area(
            "Notes",
            value="" if pd.isna(defaults.get("Notes")) else str(defaults.get("Notes") or ""),
            height=82,
            key=f"{prefix}_notes",
        )

    cpc = spend / clicks if clicks else 0.0
    cpl = spend / leads if leads else 0.0
    symbol = STORE.currency_symbols().get(currency, "")
    st.caption(f"Auto-calculated: Avg. CPC = {symbol}{cpc:,.2f} · Cost per Lead = {symbol}{cpl:,.2f}")

    record = {
        "Platform": platform,
        "Campaign Name": campaign.strip(),
        "Campaign Status": status,
        "Region": region.strip(),
        "Start Date": start_date,
        "End Date": end_date,
        "Currency": currency,
        "Daily Budget": daily_budget,
        "Impressions": impressions,
        "Clicks": clicks,
        "Total Leads": leads,
        "Total Spend": spend,
        "Notes": notes.strip(),
    }
    valid = bool(record["Campaign Name"] and record["Region"])
    return record, report_date, valid


def render_data_manager():
    master, companies, _ = get_data()
    company_names = companies["Company Name"].dropna().astype(str).tolist()

    st.markdown(
        """
        <div class="hero">
          <h1>Campaign Data Manager</h1>
          <p>Add, edit or remove campaign rows. Every save is written to both MasterData and the selected company sheet.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    company = st.selectbox("Company", company_names, key="manager_company")
    allowed = STORE.allowed_currencies(company)
    st.info(
        f"Allowed currencies for {company}: {', '.join(allowed)}. "
        "Average CPC and Cost per Lead are recalculated from spend/clicks/leads when you save."
    )

    current = STORE.get_company_rows(company)
    if not current.empty:
        view_cols = [
            "S.No",
            "Campaign Name",
            "Platform",
            "Campaign Status",
            "Region",
            "Report Date",
            "Currency",
            "Total Spend",
            "Total Leads",
        ]
        with st.expander(f"Current {company} records ({len(current)})", expanded=False):
            st.dataframe(current[[c for c in view_cols if c in current.columns]], hide_index=True, use_container_width=True)

    add_tab, edit_tab, delete_tab, file_tab = st.tabs(["＋ Add record", "✎ Update record", "Delete record", "Workbook"])

    with add_tab:
        st.subheader(f"Add a new {company} campaign")
        with st.form(f"add_form_{company}", clear_on_submit=False):
            record, report_date, valid = record_form(f"add_{company}", company)
            submitted = st.form_submit_button("Save new record", type="primary", use_container_width=True)
        if submitted:
            if not valid:
                st.error("Campaign Name and Region are required.")
            else:
                try:
                    result = STORE.save_record(company, record, report_date, s_no=None)
                    clear_data_cache()
                    st.success(
                        f"Created {company} S.No {result.s_no}. The individual sheet and MasterData were both updated."
                    )
                    st.caption(f"Automatic backup: {result.backup_path.name}")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Save failed: {exc}")

    with edit_tab:
        current = STORE.get_company_rows(company)
        if current.empty:
            st.warning("There are no records to update.")
        else:
            current["S.No"] = pd.to_numeric(current["S.No"], errors="coerce").astype("Int64")
            choice_map = {
                f"{int(row['S.No'])} — {row['Campaign Name']}": int(row["S.No"])
                for _, row in current.dropna(subset=["S.No"]).iterrows()
            }
            choice = st.selectbox("Choose campaign", list(choice_map.keys()), key=f"edit_choice_{company}")
            s_no = choice_map[choice]
            row = current[current["S.No"] == s_no].iloc[0].to_dict()
            with st.form(f"edit_form_{company}_{s_no}"):
                record, report_date, valid = record_form(f"edit_{company}_{s_no}", company, row)
                submitted = st.form_submit_button("Update record", type="primary", use_container_width=True)
            if submitted:
                if not valid:
                    st.error("Campaign Name and Region are required.")
                else:
                    try:
                        result = STORE.save_record(company, record, report_date, s_no=s_no)
                        clear_data_cache()
                        st.success(
                            f"Updated {company} S.No {result.s_no} in both the company sheet and MasterData."
                        )
                        st.caption(f"Automatic backup: {result.backup_path.name}")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Update failed: {exc}")

    with delete_tab:
        current = STORE.get_company_rows(company)
        if current.empty:
            st.warning("There are no records to delete.")
        else:
            current["S.No"] = pd.to_numeric(current["S.No"], errors="coerce").astype("Int64")
            delete_map = {
                f"{int(row['S.No'])} — {row['Campaign Name']}": int(row["S.No"])
                for _, row in current.dropna(subset=["S.No"]).iterrows()
            }
            choice = st.selectbox("Choose record to delete", list(delete_map.keys()), key=f"delete_choice_{company}")
            s_no = delete_map[choice]
            confirm = st.checkbox("I understand this deletes the row from MasterData and the individual company sheet.")
            if st.button("Delete selected record", type="primary", disabled=not confirm, use_container_width=True):
                try:
                    result = STORE.delete_record(company, s_no)
                    clear_data_cache()
                    st.success(f"Deleted {company} S.No {s_no} from both locations.")
                    st.caption(f"Automatic backup: {result.backup_path.name}")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Delete failed: {exc}")

    with file_tab:
        st.subheader("Workbook maintenance")
        st.write(f"**Active workbook:** `{STORE.path}`")
        st.write(f"**Last modified:** {pd.Timestamp(STORE.path.stat().st_mtime, unit='s'):%d %b %Y %H:%M:%S}")
        st.download_button(
            "Download current workbook",
            data=STORE.path.read_bytes(),
            file_name=STORE.path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(
            "Each write creates a timestamped backup in the backups folder. The newest 20 backups are retained."
        )


# -----------------------------------------------------------------------------
# Navigation
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Marketing Hub")
    page = st.radio("Navigation", ["Dashboard", "Data Manager"], label_visibility="collapsed")
    st.divider()
    st.caption("Excel-backed dashboard")
    st.caption("No FX conversion")

if page == "Dashboard":
    render_dashboard()
else:
    render_data_manager()
