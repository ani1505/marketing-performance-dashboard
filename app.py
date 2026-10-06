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
    "orange": "#3F7CAC",
    "orange_dark": "#17324D",
    "black": "#17324D",
    "charcoal": "#53677A",
    "text": "#25364A",
    "muted": "#718096",
    "line": "#DCE5EC",
    "panel": "#F8FAFC",
    "white": "#FFFFFF",
}

PLOT_COLORS = ["#3F7CAC", "#17324D", "#7895B2", "#A8C7DF", "#6B7A8C", "#DCE5EC"]


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
  --orange: #3F7CAC;
  --orange-soft: #EAF4FB;
  --black: #17324D;
  --grey: #6B7A8C;
}

html, body, [class*="css"] { font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
.stApp { background: var(--bg); color: var(--text); }
.block-container { padding-top: 1.25rem; padding-bottom: 2rem; max-width: 1580px; }

/* Header */
.hero {
  padding: 1.2rem 1.35rem;
  background: linear-gradient(120deg, #FFFFFF 0%, #EEF6FC 100%);
  border: 1px solid var(--line);
  border-left: 5px solid var(--orange);
  border-radius: 18px;
  margin-bottom: 0.9rem;
  box-shadow: 0 8px 24px rgba(17, 17, 17, 0.05);
}
.hero h1 { margin: 0; font-size: 1.78rem; color: var(--black); letter-spacing: -0.03em; }
.hero p { margin: 0.28rem 0 0; color: var(--muted); font-size: 0.93rem; }

/* Panels */
.panel-title { font-weight: 800; color: #202020; font-size: 1rem; margin: 0 0 0.28rem; }
.panel-subtitle { color: #7A7A7A; font-size: 0.78rem; margin-bottom: 0.4rem; }
[data-testid="stVerticalBlockBorderWrapper"] {
  border-radius: 16px !important;
  border-color: #DCE5EC !important;
  background: #FFFFFF;
  box-shadow: 0 5px 18px rgba(23,50,77,0.045);
}

/* KPI Cards */
.kpi-card {
  min-height: 118px;
  background: #FFFFFF;
  border: 1px solid #DCE5EC;
  border-top: 4px solid var(--orange);
  border-radius: 16px;
  padding: 1rem 1.05rem;
  box-shadow: 0 5px 16px rgba(17,17,17,0.04);
  display: flex;
  flex-direction: column;
  justify-content: space-between;
}
.kpi-label { color: #767676; font-size: 0.76rem; font-weight: 800; letter-spacing: 0.05em; text-transform: uppercase; }
.kpi-value { color: #17324D; font-size: 1.55rem; line-height: 1.14; font-weight: 850; letter-spacing: -0.03em; overflow-wrap: anywhere; }
.kpi-value.currency { font-size: 1.10rem; line-height: 1.3; }
.kpi-foot { color: #969696; font-size: 0.71rem; margin-top: 0.35rem; }

/* Filter section */
.filter-strip {
  background: #FFFFFF;
  border: 1px solid #DCE5EC;
  border-radius: 16px;
  padding: 0.2rem 0.75rem 0.55rem;
  margin-bottom: 0.9rem;
  box-shadow: 0 4px 14px rgba(17,17,17,0.025);
}
div[data-baseweb="select"] > div { border-radius: 10px !important; }
.stDateInput input { border-radius: 10px !important; }
.stButton > button, .stDownloadButton > button {
  border-radius: 10px;
  font-weight: 750;
  border: 1px solid #D5E0E8;
}
.stButton > button[kind="primary"] {
  background: #17324D;
  border-color: #17324D;
}

[data-testid="stDataFrame"] { border: 1px solid #DCE5EC; border-radius: 12px; overflow: hidden; }
[data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid #DCE5EC; }

.currency-note {
  display: inline-block;
  padding: 0.42rem 0.7rem;
  background: #EEF6FC;
  border: 1px solid #CFE0EE;
  border-radius: 10px;
  color: #315E83;
  font-size: 0.78rem;
}
.success-note {
  padding: 0.75rem 0.9rem;
  background: #F8FAFC;
  border: 1px solid #DCE5EC;
  color: #263445;
  border-left: 4px solid #3F7CAC;
  border-radius: 12px;
  font-size: 0.85rem;
}

.kpi-card, [data-testid="stVerticalBlockBorderWrapper"], .filter-strip {
  transition: transform 0.18s ease, box-shadow 0.18s ease;
}
.kpi-card:hover {
  transform: translateY(-2px);
  box-shadow: 0 10px 24px rgba(23,50,77,0.08);
}
.filter-strip {
  border-top: 3px solid #3F7CAC;
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
        margin=dict(l=20, r=18, t=44, b=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Inter, Arial", color="#53677A", size=12),
        title_font=dict(size=15, color="#17324D"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        showlegend=legend,
        hoverlabel=dict(bgcolor="white", font_size=12, font_color="#17324D"),
    )
    fig.update_xaxes(showgrid=False, linecolor="#DCE5EC", tickfont=dict(color="#6B7A8C"))
    fig.update_yaxes(gridcolor="#EEF3F7", zeroline=False, tickfont=dict(color="#6B7A8C"))
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
          <p>Campaign performance overview across clients, platforms, regions and dates.</p>
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
            st.caption("Date filtering uses **Report Date** from MasterData.")

    currency_options = available_currency_options(company, master, currency_df)
    selected_currencies = pills(
        "Currency",
        currency_options,
        default=currency_options.copy(),
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
        f'<span class="currency-note">Currency selection: <b>{currency_text}</b>. Financial amounts are never converted or added across different currencies.</span>',
        unsafe_allow_html=True,
    )
    st.write("")

    if filtered.empty:
        st.warning("No rows match the current filters. Try Lifetime or a different company/platform/region/date selection.")
        return

    # ---------- KPI calculations ----------
    impressions = pd.to_numeric(filtered["Impressions"], errors="coerce").fillna(0).sum()
    clicks = pd.to_numeric(filtered["Clicks"], errors="coerce").fillna(0).sum()
    leads = pd.to_numeric(filtered["Total Leads"], errors="coerce").fillna(0).sum()
    ctr = clicks / impressions if impressions else 0
    conversion = leads / clicks if clicks else 0
    spend_text = currency_breakdown(filtered, symbols, "spend")
    cpl_text = currency_breakdown(filtered, symbols, "cpl")
    cpc_text = currency_breakdown(filtered, symbols, "cpc")

    cards = [
        ("Total Spend", spend_text, "Native currency totals", True),
        ("Total Leads", human_number(leads), "Lead volume", False),
        ("Cost per Lead", cpl_text, "Spend ÷ leads", True),
        ("CTR", pct(ctr), "Clicks ÷ impressions", False),
        ("Impressions", human_number(impressions), "Ad impressions", False),
        ("Clicks", human_number(clicks), "Ad clicks", False),
        ("Avg. CPC", cpc_text, "Spend ÷ clicks", True),
        ("Lead Conversion", pct(conversion), "Leads ÷ clicks", False),
    ]
    for start in (0, 4):
        cols = st.columns(4, gap="medium")
        for col, item in zip(cols, cards[start : start + 4]):
            with col:
                metric_card(*item)
        st.write("")

    # ---------- Row 1: Leads by platform + Total spend by platform ----------
    r1a, r1b = st.columns(2, gap="medium")

    with r1a:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Leads by Platform</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-subtitle">Compare lead generation across advertising platforms.</div>', unsafe_allow_html=True)
            lead_platform = (
                filtered.assign(**{"Total Leads": pd.to_numeric(filtered["Total Leads"], errors="coerce").fillna(0)})
                .groupby("Platform", as_index=False)["Total Leads"].sum()
                .sort_values("Total Leads", ascending=True)
            )
            fig = px.bar(
                lead_platform,
                x="Total Leads",
                y="Platform",
                orientation="h",
                text_auto=True,
                color_discrete_sequence=[COLORS["orange"]],
            )
            fig.update_traces(marker_line_width=0, textposition="outside", cliponaxis=False)
            fig = style_plot(fig, height=355, legend=False)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with r1b:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Total Spend by Platform</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-subtitle">Platform-level spend. Currency is treated only as the unit of measure and is never converted or combined.</div>', unsafe_allow_html=True)
            spend_scope = filtered.copy()
            spend_scope["Total Spend"] = pd.to_numeric(spend_scope["Total Spend"], errors="coerce").fillna(0)
            spend_currencies = [str(c) for c in spend_scope["Currency"].dropna().unique().tolist()]

            def render_spend_platform_chart(frame: pd.DataFrame, currency_code: str):
                spend_platform = (
                    frame.groupby("Platform", as_index=False)["Total Spend"].sum()
                    .sort_values("Total Spend", ascending=False)
                )
                symbol = symbols.get(currency_code, "")
                fig = px.bar(
                    spend_platform,
                    x="Platform",
                    y="Total Spend",
                    text=spend_platform["Total Spend"].map(lambda x: f"{symbol}{x:,.0f}"),
                    color_discrete_sequence=[COLORS["orange_dark"]],
                    labels={"Total Spend": f"Spend ({currency_code})"},
                )
                fig.update_traces(marker_line_width=0, textposition="outside", cliponaxis=False)
                fig = style_plot(fig, height=320, legend=False)
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

            if len(spend_currencies) == 1:
                code = spend_currencies[0]
                render_spend_platform_chart(spend_scope, code)
            else:
                tabs = st.tabs(spend_currencies)
                for tab, code in zip(tabs, spend_currencies):
                    with tab:
                        render_spend_platform_chart(
                            spend_scope[spend_scope["Currency"].astype(str) == code],
                            code,
                        )

    # ---------- Row 2: Leads by date + campaign status ----------
    r2a, r2b = st.columns([1.65, 1], gap="medium")

    with r2a:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Leads by Date</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-subtitle">Lead volume over time using the report date.</div>', unsafe_allow_html=True)
            lead_date = filtered.dropna(subset=["Report Date"]).copy()
            lead_date["Report Date"] = pd.to_datetime(lead_date["Report Date"], errors="coerce")
            lead_date["Total Leads"] = pd.to_numeric(lead_date["Total Leads"], errors="coerce").fillna(0)
            lead_date = lead_date.dropna(subset=["Report Date"]).groupby("Report Date", as_index=False)["Total Leads"].sum()
            if lead_date.empty:
                st.info("No dated rows are available for this selection.")
            else:
                fig = px.line(
                    lead_date,
                    x="Report Date",
                    y="Total Leads",
                    markers=True,
                    color_discrete_sequence=[COLORS["orange"]],
                )
                fig.update_traces(line=dict(width=3), marker=dict(size=7))
                fig = style_plot(fig, height=355, legend=False)
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with r2b:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Campaign Status</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-subtitle">Current campaign mix in the selected view.</div>', unsafe_allow_html=True)
            status_df = (
                filtered["Campaign Status"].fillna("Unknown").astype(str)
                .value_counts().rename_axis("Campaign Status").reset_index(name="Campaigns")
            )
            fig = px.pie(
                status_df,
                names="Campaign Status",
                values="Campaigns",
                hole=0.55,
                color_discrete_sequence=PLOT_COLORS,
            )
            fig.update_traces(textposition="inside", textinfo="percent+label", marker=dict(line=dict(color="#FFFFFF", width=3)))
            fig.update_layout(
                height=355,
                margin=dict(l=12, r=12, t=20, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                font=dict(family="Inter, Arial", color="#263445"),
                showlegend=False,
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # ---------- Row 3: Top campaigns + campaign performance ----------
    r3a, r3b = st.columns([1, 2.05], gap="medium")

    with r3a:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Top Campaigns by Leads</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-subtitle">Highest lead-generating campaigns in the current selection.</div>', unsafe_allow_html=True)
            top_campaigns = filtered.copy()
            top_campaigns["Total Leads"] = pd.to_numeric(top_campaigns["Total Leads"], errors="coerce").fillna(0)
            top_campaigns = (
                top_campaigns.groupby("Campaign Name", as_index=False)["Total Leads"].sum()
                .nlargest(8, "Total Leads")
                .sort_values("Total Leads", ascending=True)
            )
            fig = px.bar(
                top_campaigns,
                x="Total Leads",
                y="Campaign Name",
                orientation="h",
                text_auto=True,
                color_discrete_sequence=[COLORS["orange_dark"]],
                labels={"Campaign Name": "Campaign"},
            )
            fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
            fig = style_plot(fig, height=390, legend=False)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with r3b:
        with st.container(border=True):
            st.markdown('<div class="panel-title">Campaign Performance</div>', unsafe_allow_html=True)
            st.markdown('<div class="panel-subtitle">Campaign, platform and region level results. Currency appears only inside monetary values so different currencies are never mixed.</div>', unsafe_allow_html=True)

            detail = filtered.copy()
            for col in ["Total Spend", "Total Leads", "Clicks"]:
                detail[col] = pd.to_numeric(detail[col], errors="coerce").fillna(0)
            detail["Cost Per Lead"] = detail.apply(
                lambda r: (r["Total Spend"] / r["Total Leads"]) if r["Total Leads"] else 0,
                axis=1,
            )
            detail["Lead Converted %"] = detail.apply(
                lambda r: (r["Total Leads"] / r["Clicks"] * 100) if r["Clicks"] else 0,
                axis=1,
            )
            detail["Total Spend"] = detail.apply(
                lambda r: f"{symbols.get(str(r['Currency']), '')}{r['Total Spend']:,.2f} {r['Currency']}",
                axis=1,
            )
            detail["Cost Per Lead"] = detail.apply(
                lambda r: f"{symbols.get(str(r['Currency']), '')}{r['Cost Per Lead']:,.2f} {r['Currency']}",
                axis=1,
            )
            detail["Lead Converted %"] = detail["Lead Converted %"].map(lambda x: f"{x:.2f}%")
            detail["Total Leads"] = detail["Total Leads"].astype(int)

            table_cols = [
                "Campaign Name",
                "Platform",
                "Region",
                "Total Spend",
                "Total Leads",
                "Cost Per Lead",
                "Lead Converted %",
            ]
            st.dataframe(
                detail[table_cols],
                hide_index=True,
                use_container_width=True,
                height=345,
                column_config={
                    "Campaign Name": st.column_config.TextColumn("Campaign"),
                    "Lead Converted %": st.column_config.TextColumn("Lead Converted"),
                },
            )

    # ---------- Row 4: Platform efficiency ----------
    with st.container(border=True):
        st.markdown('<div class="panel-title">Platform Efficiency</div>', unsafe_allow_html=True)
        st.markdown('<div class="panel-subtitle">CTR and lead conversion by platform, focused on performance rather than currency.</div>', unsafe_allow_html=True)
        perf = filtered.copy()
        for col in ["Clicks", "Impressions", "Total Leads"]:
            perf[col] = pd.to_numeric(perf[col], errors="coerce").fillna(0)
        perf = perf.groupby("Platform", as_index=False)[["Clicks", "Impressions", "Total Leads"]].sum()
        perf["CTR"] = perf.apply(lambda r: (r["Clicks"] / r["Impressions"] * 100) if r["Impressions"] else 0, axis=1)
        perf["Lead Conversion"] = perf.apply(lambda r: (r["Total Leads"] / r["Clicks"] * 100) if r["Clicks"] else 0, axis=1)
        perf_long = perf.melt(
            id_vars="Platform",
            value_vars=["CTR", "Lead Conversion"],
            var_name="Metric",
            value_name="Rate",
        )
        fig = px.bar(
            perf_long,
            x="Platform",
            y="Rate",
            color="Metric",
            barmode="group",
            text=perf_long["Rate"].map(lambda x: f"{x:.1f}%"),
            color_discrete_sequence=[COLORS["orange"], COLORS["orange_dark"]],
            labels={"Rate": "Rate (%)"},
        )
        fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
        fig = style_plot(fig, height=340, legend=True)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.caption(
        f"Source: {STORE.path.name} · {len(filtered):,} filtered campaign rows · "
        "Currency is used only to keep financial values mathematically valid; analysis is centered on campaign performance."
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
