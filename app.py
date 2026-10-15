from __future__ import annotations

import html
import os
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import urldefrag

import pandas as pd
import plotly.express as px
import streamlit as st

from data_store import COMPANY_CURRENCY_RULES, WorkbookStore


# -----------------------------------------------------------------------------
# App configuration
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Marketing Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = Path(
    os.getenv("DASHBOARD_WORKBOOK", str(BASE_DIR / "wolfable_master.xlsx"))
)
STORE = WorkbookStore(WORKBOOK_PATH)

# Palette sampled/matched to the supplied Power BI dashboard.
COLORS = {
    "orange": "#F18758",
    "orange_dark": "#E97849",
    "orange_soft": "#F8B08F",
    "page_grey": "#E1E1E1",
    "grid": "#ECECEC",
    "white": "#FFFFFF",
    "text": "#242424",
    "muted": "#666666",
    "header_grey": "#6C6C6C",
    "row_grey": "#F3F3F3",
    "row_grey_alt": "#E8E8E8",
    "border": "#E0E0E0",
}


# -----------------------------------------------------------------------------
# Power BI look-and-feel
# -----------------------------------------------------------------------------
st.markdown(
    f"""
    <style>
    :root {{
      --orange: {COLORS['orange']};
      --orange-dark: {COLORS['orange_dark']};
      --page-grey: {COLORS['page_grey']};
      --text: {COLORS['text']};
      --muted: {COLORS['muted']};
      --white: #FFFFFF;
      --border: {COLORS['border']};
    }}

    html, body, [class*="css"] {{
      font-family: "Segoe UI", Arial, sans-serif;
    }}

    .stApp {{
      background:
        linear-gradient(153deg, transparent 0 71%, rgba(255,255,255,.88) 71.1% 100%),
        linear-gradient(16deg, rgba(255,255,255,.78) 0 12%, transparent 12.1% 100%),
        var(--page-grey);
      color: var(--text);
    }}

    .block-container {{
      max-width: 1920px;
      padding: .25rem 4.6rem 1.1rem 4.6rem;
    }}

    /* Make Streamlit chrome disappear so the page feels like the PBIX canvas. */
    #MainMenu {{ visibility: hidden; }}
    footer {{ visibility: hidden; }}
    header[data-testid="stHeader"] {{ background: transparent; }}

    .brand-slot {{
      height: 54px;
      display: flex;
      align-items: center;
      justify-content: flex-start;
      overflow: hidden;
      margin-top: 2px;
    }}
    .brand-logo {{
      max-width: 190px;
      max-height: 50px;
      object-fit: contain;
      object-position: left center;
      display: block;
    }}
    .brand-fallback {{
      color: #000;
      font-size: 32px;
      line-height: 1;
      font-weight: 800;
      letter-spacing: -1.5px;
    }}
    .overview-title {{
      font-size: 30px;
      font-weight: 700;
      color: #828282;
      line-height: 1;
      margin: -2px 0 14px 4px;
    }}

    /* Slicers */
    [data-testid="stSelectbox"], [data-testid="stDateInput"] {{ margin-bottom: 0 !important; }}
    [data-testid="stSelectbox"] label,
    [data-testid="stDateInput"] label,
    [data-testid="stPills"] label,
    [data-testid="stButtonGroup"] label {{
      font-size: 14px !important;
      font-weight: 500 !important;
      color: #252525 !important;
      padding-bottom: 2px !important;
    }}
    div[data-baseweb="select"] > div {{
      min-height: 34px !important;
      height: 34px !important;
      border: 1px solid #C9C9C9 !important;
      border-radius: 3px !important;
      background: rgba(255,255,255,.18) !important;
      box-shadow: none !important;
    }}
    [data-testid="stDateInput"] input {{
      min-height: 34px !important;
      height: 34px !important;
      border-radius: 2px !important;
      border: 1px solid #D2D2D2 !important;
      background: rgba(255,255,255,.12) !important;
      color: #333 !important;
      font-size: 13px !important;
    }}


    .slicer-caption {{
      font-size: 14px;
      font-weight: 500;
      color: #252525;
      line-height: 1.15;
      margin: 0 0 5px 0;
    }}
    /* Currency/date tile controls: only the actual choice buttons are visible. */
    [data-testid="stButtonGroup"] {{
      margin-top: 0 !important;
      width: 100% !important;
    }}
    [data-testid="stButtonGroup"] > div {{
      display: flex !important;
      width: 100% !important;
      gap: 5px !important;
      flex-wrap: nowrap !important;
      overflow: visible !important;
    }}

    /* Currency + Date buttons: fully clickable, responsive Power BI tiles. */
    [data-testid="stPills"] button,
    [data-testid="stButtonGroup"] button,
    div[role="radiogroup"] button {{
      flex: 1 1 0 !important;
      min-width: 0 !important;
      width: 100% !important;
      min-height: 34px !important;
      height: 34px !important;
      padding: 2px 8px !important;
      border-radius: 4px !important;
      border: 1px solid #D4D4D4 !important;
      background: #FFFFFF !important;
      color: #333333 !important;
      font-size: 13px !important;
      line-height: 1 !important;
      white-space: nowrap !important;
      cursor: pointer !important;
      box-shadow: none !important;
      transition: background-color .16s ease, border-color .16s ease,
                  box-shadow .16s ease, transform .16s ease !important;
    }}

    [data-testid="stPills"] button:hover,
    [data-testid="stButtonGroup"] button:hover,
    div[role="radiogroup"] button:hover {{
      background: #EFEFEF !important;
      border-color: #AFAFAF !important;
      box-shadow: 0 2px 6px rgba(0,0,0,.10) !important;
      transform: translateY(-1px);
    }}

    /* Selected tile is grey, as in the Power BI-style slicer. */
    [data-testid="stPills"] button[aria-pressed="true"],
    [data-testid="stButtonGroup"] button[aria-pressed="true"],
    div[role="radiogroup"] button[aria-pressed="true"] {{
      background: #767676 !important;
      border-color: #767676 !important;
      color: #FFFFFF !important;
      font-weight: 600 !important;
      box-shadow: inset 0 0 0 1px rgba(255,255,255,.08) !important;
    }}
    [data-testid="stPills"] button[aria-pressed="true"]:hover,
    [data-testid="stButtonGroup"] button[aria-pressed="true"]:hover,
    div[role="radiogroup"] button[aria-pressed="true"]:hover {{
      background: #666666 !important;
      border-color: #666666 !important;
    }}

    /* Keep all date choices on one row. Custom is never clipped. */
    [data-testid="stPills"] > div {{
      width: 100% !important;
      display: flex !important;
      gap: 5px !important;
      flex-wrap: nowrap !important;
      overflow: visible !important;
    }}
    .kpi-row-divider {{ height: 9px; }}

    /* Native Streamlit button tiles used for Currency and Date. */
    div[data-testid="stButton"] > button {{
      width: 100% !important;
      min-width: 0 !important;
      height: 36px !important;
      min-height: 36px !important;
      padding: 0 10px !important;
      border-radius: 4px !important;
      border: 1px solid #D2D2D2 !important;
      background: #FFFFFF !important;
      color: #333333 !important;
      font-size: 13px !important;
      font-weight: 500 !important;
      cursor: pointer !important;
      box-shadow: none !important;
      transition: background-color .15s ease, border-color .15s ease, box-shadow .15s ease, transform .15s ease !important;
    }}
    div[data-testid="stButton"] > button p,
    div[data-testid="stButton"] > button span {{
      white-space: nowrap !important;
      overflow: visible !important;
      text-overflow: clip !important;
      max-width: none !important;
    }}
    div[data-testid="stButton"] > button:hover {{
      background: #ECECEC !important;
      border-color: #AFAFAF !important;
      box-shadow: 0 2px 6px rgba(0,0,0,.10) !important;
      transform: translateY(-1px) !important;
    }}
    div[data-testid="stButton"] > button[kind="primary"],
    div[data-testid="stButton"] > button[data-testid="baseButton-primary"] {{
      background: #777777 !important;
      border-color: #777777 !important;
      color: #FFFFFF !important;
      font-weight: 600 !important;
    }}
    div[data-testid="stButton"] > button[kind="primary"] p,
    div[data-testid="stButton"] > button[data-testid="baseButton-primary"] p {{
      color: #FFFFFF !important;
    }}
    div[data-testid="stButton"] > button[kind="primary"]:hover,
    div[data-testid="stButton"] > button[data-testid="baseButton-primary"]:hover {{
      background: #666666 !important;
      border-color: #666666 !important;
    }}

    /* KPI cards */
    .kpi-card {{
      height: 49px;
      min-height: 49px;
      background: rgba(255,255,255,.96);
      border: 0;
      border-radius: 7px;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 3px 8px 2px;
      overflow: hidden;
    }}
    .kpi-label {{
      color: #6B6B6B;
      font-size: 12px;
      font-weight: 400;
      line-height: 1;
      white-space: nowrap;
    }}
    .kpi-value {{
      color: #202020;
      font-size: 22px;
      font-weight: 600;
      line-height: 1.15;
      margin-top: 4px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 100%;
    }}
    .kpi-value.money {{ font-size: 20px; }}

    /* White visual canvases from the report. */
    [data-testid="stVerticalBlockBorderWrapper"] {{
      border: 0 !important;
      border-radius: 7px !important;
      background: rgba(255,255,255,.97) !important;
      box-shadow: none !important;
      overflow: hidden !important;
    }}
    .visual-title {{
      font-size: 14px;
      font-weight: 500;
      color: #242424;
      line-height: 1.1;
      margin: 2px 0 0 1px;
    }}
    .visual-subtitle {{
      font-size: 12px;
      font-weight: 400;
      color: #666666;
      line-height: 1.05;
      margin: 4px 0 0 1px;
    }}

    [data-testid="stDataFrame"] {{ border: 0 !important; border-radius: 0 !important; }}
    [data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 1px solid #DDDDDD; }}

    /* Tighten the vertical rhythm to match a 1920×1080 Power BI canvas. */
    div[data-testid="stHorizontalBlock"] {{ gap: .72rem; }}
    div[data-testid="stVerticalBlock"] > div:has(.kpi-card) {{ margin-bottom: 0 !important; }}

    @media (max-width: 1200px) {{
      .block-container {{ padding-left: 1rem; padding-right: 1rem; }}
      .brand-logo {{ max-width: 155px; }}
      .kpi-value {{ font-size: 18px; }}
      .kpi-value.money {{ font-size: 15px; }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False, ttl=20)
def load_data(path_str: str, modified_time: float):
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
        return f"{value / 1_000:.0f}K"
    return f"{value:,.0f}"


def ratio_text(value: float) -> str:
    return f"{float(value or 0):.2f}"


def currency_breakdown(df: pd.DataFrame, symbols: dict[str, str], metric: str = "spend") -> str:
    """Return each currency independently. Never add GBP, INR, USD, etc. together."""
    if df.empty:
        return "—"
    parts: list[str] = []
    currency_values = df.get("Currency", pd.Series(dtype=str)).dropna().astype(str).unique().tolist()
    currency_values = sorted(currency_values, key=lambda c: {"GBP": 0, "INR": 1, "USD": 2}.get(c, 99))
    for currency in currency_values:
        group = df[df["Currency"].astype(str) == currency]
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
    return " | ".join(parts)


def metric_card(label: str, value: str, money: bool = False):
    value_class = "kpi-value money" if money else "kpi-value"
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-label">{html.escape(label)}</div>'
        f'<div class="{value_class}">{html.escape(value)}</div></div>',
        unsafe_allow_html=True,
    )


def style_plot(fig, height=330, legend=False):
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=38, r=20, t=16, b=34),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(family="Segoe UI, Arial", color=COLORS["text"], size=11),
        showlegend=legend,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.01,
            xanchor="right",
            x=1,
            font=dict(size=10),
        ),
        hoverlabel=dict(bgcolor="white", font_color="#242424", bordercolor="#D5D5D5"),
        bargap=.58,
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor=COLORS["grid"],
        linecolor="rgba(0,0,0,0)",
        zeroline=False,
        tickfont=dict(color="#646464", size=10),
        title_font=dict(size=10, color="#444444"),
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor=COLORS["grid"],
        linecolor="rgba(0,0,0,0)",
        zeroline=False,
        tickfont=dict(color="#646464", size=10),
        title_font=dict(size=10, color="#444444"),
    )
    return fig


def available_currency_options(company: str, master: pd.DataFrame, currency_df: pd.DataFrame) -> list[str]:
    all_codes = [
        str(v).strip()
        for v in currency_df.get("Currency", pd.Series(dtype=str)).dropna().tolist()
    ]
    allowed = COMPANY_CURRENCY_RULES.get(company, all_codes)
    present = [
        str(v).strip()
        for v in master.loc[
            master["Company"].astype(str) == company, "Currency"
        ].dropna().unique().tolist()
    ]
    return [c for c in allowed if c in present] or allowed


def company_logo_url(companies: pd.DataFrame, company: str) -> str | None:
    if "Logo" not in companies.columns:
        return None
    row = companies[companies["Company Name"].astype(str) == company]
    if row.empty:
        return None
    value = row.iloc[0].get("Logo")
    if pd.isna(value) or not str(value).strip():
        return None
    return urldefrag(str(value).strip())[0]


def _tile_button_row(options: list[str], state_key: str, default: str | None = None, allow_clear: bool = False) -> str | None:
    """Render responsive, fully-clickable Streamlit buttons with selected state."""
    if state_key not in st.session_state:
        st.session_state[state_key] = default

    # Keep stale selections from surviving when the available options change.
    if st.session_state[state_key] not in options and st.session_state[state_key] is not None:
        st.session_state[state_key] = default

    cols = st.columns([1] * max(len(options), 1), gap="small")
    current = st.session_state[state_key]
    for col, option in zip(cols, options):
        with col:
            selected = current == option
            if st.button(
                option,
                key=f"{state_key}__{option}",
                type="primary" if selected else "secondary",
                use_container_width=True,
            ):
                if allow_clear and selected:
                    st.session_state[state_key] = None
                else:
                    st.session_state[state_key] = option
                current = st.session_state[state_key]
                st.rerun()
    return current


def currency_selector(options: list[str], key: str) -> str:
    """Power BI-style currency tiles. No selected tile means all company currencies."""
    st.markdown('<div class="slicer-caption">Currency</div>', unsafe_allow_html=True)
    selected = _tile_button_row(list(options), key, default=None, allow_clear=True)
    return selected or "All"

def _visual_heading(title: str, subtitle: str | None = None):
    st.markdown(f'<div class="visual-title">{html.escape(title)}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="visual-subtitle">{html.escape(subtitle)}</div>', unsafe_allow_html=True)


def _brand_html(logo_url: str | None, company: str) -> str:
    if logo_url:
        safe_url = html.escape(logo_url, quote=True)
        safe_company = html.escape(company, quote=True)
        return (
            '<div class="brand-slot">'
            f'<img class="brand-logo" src="{safe_url}" alt="{safe_company} logo" '
            'onerror="this.style.display=\'none\';this.nextElementSibling.style.display=\'block\';">'
            f'<div class="brand-fallback" style="display:none">{html.escape(company)}</div>'
            '</div>'
        )
    return f'<div class="brand-slot"><div class="brand-fallback">{html.escape(company)}</div></div>'


def _table_styler(df: pd.DataFrame):
    def row_style(row):
        if str(row.get("Campaign Name", "")) == "Total":
            return [
                f"background-color: {COLORS['orange']}; color: white; font-weight: 700;"
                for _ in row
            ]
        idx = row.name if isinstance(row.name, int) else 0
        bg = COLORS["row_grey"] if idx % 2 == 0 else COLORS["row_grey_alt"]
        return [f"background-color: {bg}; color: #242424;" for _ in row]

    return (
        df.style
        .apply(row_style, axis=1)
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("background-color", COLORS["header_grey"]),
                    ("color", "white"),
                    ("font-weight", "500"),
                    ("border", "1px solid #BFBFBF"),
                ],
            },
            {
                "selector": "td",
                "props": [("border", "1px solid #F1F1F1")],
            },
        ])
    )


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

    company_names = companies["Company Name"].dropna().astype(str).tolist()
    default_company = "Wolfable" if "Wolfable" in company_names else company_names[0]

    # Company must be resolved first so the logo, region/platform and currency controls
    # all react exactly like report slicers.
    logo_col, company_col, region_col, platform_col, currency_col, gap_col, date_col = st.columns(
        [2.05, 1.50, 1.40, 1.40, 1.80, .08, 4.25], gap="small"
    )

    with company_col:
        company = st.selectbox(
            "Company Name",
            company_names,
            index=company_names.index(default_company) if default_company in company_names else 0,
            key="dashboard_company",
        )

    logo_url = company_logo_url(companies, company)
    with logo_col:
        st.markdown(_brand_html(logo_url, company), unsafe_allow_html=True)

    company_scope = master[master["Company"].astype(str) == company].copy()
    region_options = sorted(company_scope["Region"].dropna().astype(str).unique().tolist())
    platform_options = sorted(company_scope["Platform"].dropna().astype(str).unique().tolist())

    with region_col:
        region = st.selectbox("Region", ["All"] + region_options, index=0, key="dashboard_region")
    with platform_col:
        platform = st.selectbox("Platform", ["All"] + platform_options, index=0, key="dashboard_platform")

    currency_options = available_currency_options(company, master, currency_df)
    with currency_col:
        currency_choice = currency_selector(currency_options, key=f"currency_choice_{company}")

    date_series = pd.to_datetime(company_scope.get("Report Date"), errors="coerce").dropna()
    if date_series.empty:
        min_date = date.today() - timedelta(days=30)
        max_date = date.today()
    else:
        min_date = date_series.min().date()
        max_date = date_series.max().date()

    with date_col:
        st.markdown('<div class="slicer-caption">Date</div>', unsafe_allow_html=True)
        date_choices = ["Yesterday", "Last Week", "Lifetime", "Custom"]
        date_period = _tile_button_row(
            date_choices,
            state_key="dashboard_date_period_value",
            default="Lifetime",
            allow_clear=False,
        ) or "Lifetime"

    # A date-range picker is visible only for the Custom period.
    # Yesterday and Last Week are based on today's date, rather than the last
    # report date, and Lifetime includes all historical records.
    start_date = None
    end_date = None
    if date_period == "Yesterday":
        start_date = end_date = date.today() - timedelta(days=1)
    elif date_period == "Last Week":
        this_monday = date.today() - timedelta(days=date.today().weekday())
        start_date = this_monday - timedelta(days=7)
        end_date = this_monday - timedelta(days=1)
    elif date_period == "Custom":
        with date_col:
            picked = st.date_input(
                "Custom date range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max(max_date, date.today()),
                format="MM/DD/YYYY",
                key=f"dashboard_custom_date_{company}",
            )
        if isinstance(picked, (tuple, list)) and len(picked) == 2:
            start_date, end_date = picked
        else:
            st.info("Choose both the start and end dates for the custom range.")
            st.stop()

    st.markdown('<div class="overview-title">Overview</div>', unsafe_allow_html=True)

    filtered = company_scope.copy()
    if region != "All":
        filtered = filtered[filtered["Region"].astype(str) == region]
    if platform != "All":
        filtered = filtered[filtered["Platform"].astype(str) == platform]
    if currency_choice != "All":
        filtered = filtered[filtered["Currency"].astype(str) == currency_choice]

    if start_date is not None and end_date is not None:
        report_dates = pd.to_datetime(filtered["Report Date"], errors="coerce").dt.date
        filtered = filtered[(report_dates >= start_date) & (report_dates <= end_date)]

    if filtered.empty:
        st.warning("No rows match the current selection.")
        return

    impressions = pd.to_numeric(filtered["Impressions"], errors="coerce").fillna(0).sum()
    clicks = pd.to_numeric(filtered["Clicks"], errors="coerce").fillna(0).sum()
    leads = pd.to_numeric(filtered["Total Leads"], errors="coerce").fillna(0).sum()
    ctr = clicks / impressions if impressions else 0
    conversion = leads / clicks if clicks else 0
    spend_text = currency_breakdown(filtered, symbols, "spend")
    cpl_text = currency_breakdown(filtered, symbols, "cpl")
    cpc_text = currency_breakdown(filtered, symbols, "cpc")

    row1 = st.columns(4, gap="small")
    for col, item in zip(
        row1,
        [
            ("Total Spend Card", spend_text, True),
            ("Leads", human_number(leads), False),
            ("CPL Card", cpl_text, True),
            ("CTR", ratio_text(ctr), False),
        ],
    ):
        with col:
            metric_card(*item)

    st.markdown('<div class="kpi-row-divider"></div>', unsafe_allow_html=True)
    row2 = st.columns(4, gap="small")
    for col, item in zip(
        row2,
        [
            ("Impressions", human_number(impressions), False),
            ("Clicks", human_number(clicks), False),
            ("CPC Card", cpc_text, True),
            ("Conversion Rate", ratio_text(conversion), False),
        ],
    ):
        with col:
            metric_card(*item)

    st.write("")

    # --- Main analytical row -------------------------------------------------
    c1, c2, c3 = st.columns([532, 554, 617], gap="small")

    with c1:
        with st.container(border=True):
            _visual_heading("Sum of Total Leads", "by Platform")
            lead_platform = filtered.copy()
            lead_platform["Total Leads"] = pd.to_numeric(lead_platform["Total Leads"], errors="coerce").fillna(0)
            lead_platform = (
                lead_platform.groupby("Platform", as_index=False)["Total Leads"].sum()
                .sort_values("Total Leads", ascending=False)
            )
            fig = px.bar(
                lead_platform,
                x="Platform",
                y="Total Leads",
                text="Total Leads",
                color_discrete_sequence=[COLORS["orange"]],
            )
            fig.update_traces(marker_line_width=0, textposition="outside", cliponaxis=False)
            fig = style_plot(fig, height=310, legend=False)
            fig.update_xaxes(title_text="Platform", showgrid=False)
            fig.update_yaxes(title_text="Sum of Total Leads")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with c2:
        with st.container(border=True):
            _visual_heading("Spend Amount", "by Platform")
            spend_scope = filtered.copy()
            spend_scope["Total Spend"] = pd.to_numeric(spend_scope["Total Spend"], errors="coerce").fillna(0)
            distinct_currencies = spend_scope["Currency"].dropna().astype(str).unique().tolist()

            if currency_choice != "All" or len(distinct_currencies) <= 1:
                code = currency_choice if currency_choice != "All" else (distinct_currencies[0] if distinct_currencies else "")
                spend_platform = (
                    spend_scope.groupby("Platform", as_index=False)["Total Spend"].sum()
                    .sort_values("Total Spend", ascending=True)
                )
                symbol = symbols.get(code, "")
                fig = px.bar(
                    spend_platform,
                    x="Total Spend",
                    y="Platform",
                    orientation="h",
                    text=spend_platform["Total Spend"].map(lambda x: f"{symbol}{x:,.0f}"),
                    color_discrete_sequence=[COLORS["orange"]],
                )
                fig.update_traces(marker_line_width=0, textposition="outside", cliponaxis=False)
                fig = style_plot(fig, height=310, legend=False)
            else:
                # All currencies: keep currencies separate so GBP + INR is never summed.
                spend_platform = (
                    spend_scope.groupby(["Platform", "Currency"], as_index=False)["Total Spend"].sum()
                    .sort_values("Total Spend", ascending=True)
                )
                fig = px.bar(
                    spend_platform,
                    x="Total Spend",
                    y="Platform",
                    color="Currency",
                    barmode="group",
                    orientation="h",
                    text="Total Spend",
                    color_discrete_sequence=[COLORS["orange"], COLORS["header_grey"], COLORS["orange_soft"]],
                )
                fig.update_traces(
                    marker_line_width=0,
                    texttemplate="%{text:,.0f}",
                    textposition="outside",
                    cliponaxis=False,
                )
                fig = style_plot(fig, height=310, legend=True)

            fig.update_xaxes(title_text="")
            fig.update_yaxes(title_text="", showgrid=False)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with c3:
        with st.container(border=True):
            _visual_heading("Sum of Total Leads", "by Date")
            lead_date = filtered.copy()
            lead_date["Report Date"] = pd.to_datetime(lead_date["Report Date"], errors="coerce")
            lead_date["Total Leads"] = pd.to_numeric(lead_date["Total Leads"], errors="coerce").fillna(0)
            lead_date = (
                lead_date.dropna(subset=["Report Date"])
                .groupby("Report Date", as_index=False)["Total Leads"].sum()
                .sort_values("Report Date")
            )
            fig = px.line(
                lead_date,
                x="Report Date",
                y="Total Leads",
                markers=False,
                color_discrete_sequence=[COLORS["orange"]],
            )
            fig.update_traces(line=dict(width=2.5, shape="spline", smoothing=.65))
            fig = style_plot(fig, height=310, legend=False)
            fig.update_xaxes(title_text="", showgrid=False, tickformat="%b %Y")
            fig.update_yaxes(title_text="")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # --- Bottom row ----------------------------------------------------------
    b1, b2 = st.columns([1104, 618], gap="small")

    with b1:
        with st.container(border=True):
            detail = filtered.copy()
            for col in ["Total Leads", "Total Spend", "Clicks"]:
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
                lambda r: f"{symbols.get(str(r['Currency']), '')}{r['Cost Per Lead']:,.2f}",
                axis=1,
            )
            detail["Lead Converted %"] = detail["Lead Converted %"].map(lambda x: f"{x:.2f}%")
            detail["Total Leads"] = detail["Total Leads"].astype(int)

            table_cols = [
                "Campaign Name",
                "Platform",
                "Region",
                "Campaign Status",
                "Total Spend",
                "Total Leads",
                "Cost Per Lead",
                "Lead Converted %",
            ]
            detail = detail[table_cols]

            total_row = {
                "Campaign Name": "Total",
                "Platform": "",
                "Region": "",
                "Campaign Status": "",
                "Total Spend": currency_breakdown(filtered, symbols, "spend"),
                "Total Leads": int(leads),
                "Cost Per Lead": currency_breakdown(filtered, symbols, "cpl"),
                "Lead Converted %": ratio_text(conversion),
            }
            table_view = pd.concat([detail, pd.DataFrame([total_row])], ignore_index=True)
            st.dataframe(
                _table_styler(table_view),
                hide_index=True,
                use_container_width=True,
                height=303,
                column_config={
                    "Campaign Name": st.column_config.TextColumn("Campaign Name", width="large"),
                    "Platform": st.column_config.TextColumn("Platform", width="small"),
                    "Region": st.column_config.TextColumn("Region", width="small"),
                    "Campaign Status": st.column_config.TextColumn("Campaign Status", width="small"),
                    "Total Spend": st.column_config.TextColumn("Total Spend", width="medium"),
                    "Total Leads": st.column_config.NumberColumn("Sum of Total Leads", format="%d", width="small"),
                    "Cost Per Lead": st.column_config.TextColumn("Cost Per Lead", width="medium"),
                    "Lead Converted %": st.column_config.TextColumn("Lead Converted", width="small"),
                },
            )

    with b2:
        with st.container(border=True):
            _visual_heading("Sum of Total Leads", "by Campaign Name")
            campaigns = filtered.copy()
            campaigns["Total Leads"] = pd.to_numeric(campaigns["Total Leads"], errors="coerce").fillna(0)
            campaigns = (
                campaigns.groupby("Campaign Name", as_index=False)["Total Leads"].sum()
                .sort_values("Total Leads", ascending=False)
                .head(5)
                .sort_values("Total Leads", ascending=True)
            )
            fig = px.bar(
                campaigns,
                x="Total Leads",
                y="Campaign Name",
                orientation="h",
                color_discrete_sequence=[COLORS["orange"]],
            )
            fig.update_traces(marker_line_width=0)
            fig = style_plot(fig, height=318, legend=False)
            fig.update_xaxes(title_text="", showgrid=True)
            fig.update_yaxes(title_text="", showgrid=False)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


# -----------------------------------------------------------------------------
# Data Manager post-action state helpers
# -----------------------------------------------------------------------------
def _queue_manager_success(action: str, company: str, message: str, backup_name: str | None = None):
    """Persist a success message across rerun and request a clean form reset."""
    st.session_state["_manager_flash"] = {
        "message": message,
        "backup": backup_name or "",
    }
    st.session_state["_manager_reset"] = {
        "action": action,
        "company": company,
    }


def _consume_manager_post_action():
    """Clear the completed operation's widgets before they are instantiated again."""
    reset = st.session_state.pop("_manager_reset", None)
    if reset:
        action = reset.get("action")
        company = reset.get("company", "")

        if action == "add":
            prefixes = [f"add_{company}_"]
            exact_keys = []
        elif action == "edit":
            prefixes = [f"edit_{company}_"]
            exact_keys = [f"edit_choice_{company}"]
        elif action == "delete":
            prefixes = []
            exact_keys = [f"delete_choice_{company}", f"delete_confirm_{company}"]
        else:
            prefixes, exact_keys = [], []

        for key in list(st.session_state.keys()):
            if key in exact_keys or any(str(key).startswith(prefix) for prefix in prefixes):
                del st.session_state[key]

    return st.session_state.pop("_manager_flash", None)


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
    # Process the previous successful operation before creating any widgets.
    # This lets us safely reset the completed form and show a persistent message.
    flash = _consume_manager_post_action()

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

    if flash:
        st.success(flash.get("message", "Operation completed successfully."))
        if flash.get("backup"):
            st.caption(f"Automatic backup: {flash['backup']}")

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
                    _queue_manager_success(
                        "add",
                        company,
                        f"Created {company} S.No {result.s_no} successfully. The company sheet and MasterData were both updated.",
                        result.backup_path.name,
                    )
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
                        _queue_manager_success(
                            "edit",
                            company,
                            f"Updated {company} S.No {result.s_no} successfully in both the company sheet and MasterData.",
                            result.backup_path.name,
                        )
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
            confirm = st.checkbox(
                "I understand this deletes the row from MasterData and the individual company sheet.",
                key=f"delete_confirm_{company}",
            )
            if st.button("Delete selected record", type="primary", disabled=not confirm, use_container_width=True):
                try:
                    result = STORE.delete_record(company, s_no)
                    clear_data_cache()
                    _queue_manager_success(
                        "delete",
                        company,
                        f"Deleted {company} S.No {s_no} successfully from both MasterData and the company sheet.",
                        result.backup_path.name,
                    )
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
    

if page == "Dashboard":
    render_dashboard()
else:
    render_data_manager()
