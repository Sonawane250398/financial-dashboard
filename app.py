"""
Financial analytics dashboard — Dash + Plotly.
Run: python3 app.py  →  http://127.0.0.1:8050
"""

import re

import dash
from dash import dcc, html, Input, Output, State
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from flask_caching import Cache
import base64
import io

# --- Theme & palette (Power BI–style light) ---
BG = "#F4F6F9"
CARD_BG = "#FFFFFF"
TEXT = "#1a1a2e"
TEXT_MUTED = "#6b7280"
BORDER = "#E5E7EB"
GRID = "#F3F4F6"
BLUE = "#2563EB"
GREEN = "#059669"
PURPLE = "#7C3AED"
AMBER = "#D97706"
RED = "#DC2626"
SHADOW = "0 1px 4px rgba(0,0,0,0.08)"
PALETTE = [BLUE, GREEN, PURPLE, AMBER]
SEGMENT_COLORS = ["#2563EB", "#059669", "#7C3AED", "#D97706", "#DC2626"]


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [
        re.sub(r"\s+", "_", str(c).strip().lower().replace(" ", "_"))
        for c in out.columns
    ]
    return out


def load_and_prepare() -> pd.DataFrame:
    df = pd.read_excel("financial_sample.xlsx", engine="openpyxl")
    df = clean_column_names(df)
    date_col = None
    for name in ("date", "order_date", "sale_date"):
        if name in df.columns:
            date_col = name
            break
    if date_col is None:
        raise ValueError("Expected a Date column in financial_sample.xlsx")
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=[date_col])
    df["year"] = df[date_col].dt.year.astype(int)
    df["month"] = df[date_col].dt.month.astype(int)
    df["year_month"] = df[date_col].dt.to_period("M").astype(str)

    gs = df["gross_sales"] if "gross_sales" in df.columns else df.get("sales", pd.Series(0, index=df.index))
    pr = df["profit"] if "profit" in df.columns else pd.Series(0, index=df.index)
    df["profit_margin_pct"] = np.where(gs != 0, (pr / gs) * 100.0, 0.0)

    if "discounts" not in df.columns:
        df["discounts"] = 0.0
    if "sales" not in df.columns:
        df["sales"] = gs
    if "units_sold" not in df.columns:
        df["units_sold"] = 0

    return df

def apply_filters(df: pd.DataFrame, segment: str, country: str, year_val) -> pd.DataFrame:
    d = df.copy()
    if segment and segment != "all":
        d = d[d["segment"] == segment]
    if country and country != "all":
        d = d[d["country"] == country]
    if year_val is not None and year_val != "all":
        d = d[d["year"] == int(year_val)]
    return d


def plot_layout(fig: go.Figure) -> None:
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=TEXT, size=12, family="system-ui, sans-serif"),
        margin=dict(l=48, r=24, t=44, b=48),
        showlegend=True,
        title_font=dict(color=TEXT, size=13, family="system-ui, sans-serif", weight=600),
        legend=dict(
            bgcolor="rgba(0,0,0,0)",
            borderwidth=0,
            font=dict(color=TEXT_MUTED, size=11),
        ),
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor=GRID,
        zeroline=False,
        showline=False,
        tickfont=dict(color=TEXT_MUTED, size=11),
        title_font=dict(color=TEXT_MUTED, size=11),
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor=GRID,
        zeroline=False,
        showline=False,
        tickfont=dict(color=TEXT_MUTED, size=11),
        title_font=dict(color=TEXT_MUTED, size=11),
    )


PLOT_CONFIG = {"displayModeBar": False, "staticPlot": False}

app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "Financial Performance Report"

cache = Cache(app.server, config={"CACHE_TYPE": "SimpleCache", "CACHE_DEFAULT_TIMEOUT": 300})

_DROPDOWN_CSS = """
<style>
#filter-segment .Select-control, #filter-country .Select-control, #filter-year .Select-control,
#filter-segment .Select__control, #filter-country .Select__control, #filter-year .Select__control {
    border: 1px solid #E5E7EB !important;
    border-radius: 8px !important;
    min-height: 38px !important;
    box-shadow: none !important;
}
#filter-segment .Select-value-label, #filter-country .Select-value-label, #filter-year .Select-value-label,
#filter-segment .Select__single-value, #filter-country .Select__single-value, #filter-year .Select__single-value {
    color: #1a1a2e !important;
}
#filter-segment .Select-placeholder, #filter-country .Select-placeholder, #filter-year .Select-placeholder,
#filter-segment .Select__placeholder, #filter-country .Select__placeholder, #filter-year .Select__placeholder {
    color: #6b7280 !important;
}
</style>
"""

_EXTRA_CSS = """
<style>
/* ── Header: thin blue accent bar on top ─────────────────────────── */
.dash-header-bar {
  border-top: 3px solid #2563EB !important;
}

/* ── Chart cards: hover lift ─────────────────────────────────────── */
.dash-chart-card {
  transition: box-shadow 0.18s ease, transform 0.18s ease;
}
.dash-chart-card:hover {
  transform: translateY(-2px);
  box-shadow: 0 6px 18px rgba(0,0,0,0.10) !important;
}

/* ── Filter active badge pulse glow ──────────────────────────────── */
.filter-active-pill {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 3px 10px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 600;
  background: rgba(37,99,235,0.10);
  color: #2563EB;
  border: 1px solid rgba(37,99,235,0.25);
}

/* ── Insights list items ─────────────────────────────────────────── */
.insight-item {
  display: flex;
  align-items: flex-start;
  gap: 10px;
  padding: 8px 12px;
  border-radius: 8px;
  margin-bottom: 8px;
  background: #F8FAFF;
  border-left: 3px solid #2563EB;
  font-size: 13px;
}
.insight-item.green  { border-left-color: #059669; background: #F0FDF4; }
.insight-item.purple { border-left-color: #7C3AED; background: #F5F3FF; }
.insight-item.amber  { border-left-color: #D97706; background: #FFFBEB; }

/* ── KPI pct change chip ─────────────────────────────────────────── */
.kpi-pct {
  font-size: 11px;
  font-weight: 600;
  padding: 2px 7px;
  border-radius: 999px;
  margin-left: 8px;
}
.kpi-pct.up   { background: #DCFCE7; color: #166534; }
.kpi-pct.down { background: #FEE2E2; color: #991B1B; }
.kpi-pct.flat { background: #F3F4F6; color: #6b7280; }

/* ── Download button hover ───────────────────────────────────────── */
#btn-download:hover {
  background-color: #EFF6FF !important;
}
</style>
"""

_RESPONSIVE_CSS = """
<style>
/* ── Mobile (≤768 px) ─────────────────────────────────────────────── */
@media (max-width: 768px) {
  .dash-main-layout {
    font-size: 80%;
  }
  /* Header */
  .dash-header-bar {
    padding: 10px 12px !important;
  }
  .dash-header-top {
    flex-direction: column !important;
    align-items: flex-start !important;
    gap: 10px !important;
  }
  .dash-header-badges {
    width: 100%;
    flex-wrap: wrap !important;
    justify-content: flex-start !important;
    gap: 6px !important;
  }
  /* Upload button — full width on mobile */
  #upload-data {
    width: 100% !important;
    justify-content: center !important;
  }
  /* Content area */
  .dash-content {
    padding: 10px 10px 24px !important;
  }
  /* Summary stats row — stack vertically */
  .dash-content > div:first-child {
    flex-direction: column !important;
    gap: 6px !important;
    padding: 10px 14px !important;
  }
  /* Filters */
  .dash-filter-card {
    padding: 10px 12px !important;
    margin-bottom: 10px !important;
  }
  .dash-filter-row {
    flex-direction: column !important;
    gap: 10px !important;
  }
  /* KPI cards — 2-column grid */
  #kpi-container {
    grid-template-columns: repeat(2, 1fr) !important;
    gap: 8px !important;
    margin-bottom: 14px !important;
  }
  /* Charts — single column */
  .dash-chart-grid-2,
  .dash-chart-grid-3 {
    grid-template-columns: 1fr !important;
    gap: 10px !important;
    margin-bottom: 10px !important;
  }
  /* Variance table — horizontal scroll */
  #variance-table table {
    display: block;
    overflow-x: auto;
    white-space: nowrap;
  }
  /* Insights card + variance section padding */
  #insights-card,
  #variance-table {
    padding: 14px !important;
  }
  /* Segment progress bars — tighter */
  #segment-bars {
    padding: 4px 0 !important;
  }
}

/* ── Small mobile (≤480 px) ───────────────────────────────────────── */
@media (max-width: 480px) {
  #kpi-container {
    grid-template-columns: 1fr !important;
  }
  .dash-header-bar h2,
  .dash-header-bar span {
    font-size: 14px !important;
  }
}
</style>
"""

_VIEWPORT_META = '<meta name="viewport" content="width=device-width, initial-scale=1.0">'

if "</head>" in app.index_string:
    app.index_string = app.index_string.replace(
        "</head>",
        _VIEWPORT_META + "\n" + _DROPDOWN_CSS + "\n" + _EXTRA_CSS + "\n" + _RESPONSIVE_CSS + "\n</head>",
        1,
    )


@cache.memoize()
def load_and_prepare_cached() -> pd.DataFrame:
    return load_and_prepare()


try:
    RAW_DF = load_and_prepare_cached()
    DATA_ERROR = None
except Exception:
    RAW_DF = None
    DATA_ERROR = "Error loading data. Please check that financial_sample.xlsx exists."


segments = sorted(RAW_DF["segment"].dropna().unique().tolist()) if RAW_DF is not None else []
countries = sorted(RAW_DF["country"].dropna().unique().tolist()) if RAW_DF is not None else []
if RAW_DF is not None:
    year_counts = RAW_DF["year"].value_counts()
    years = sorted(year_counts[year_counts >= 100].index.tolist())

    _hdr_sales = RAW_DF["sales"].sum()
    _hdr_profit = RAW_DF["profit"].sum()
    _hdr_gs = RAW_DF["gross_sales"].sum()
    _hdr_margin = (_hdr_profit / _hdr_gs * 100.0) if _hdr_gs else 0.0
    _year_lo = int(RAW_DF["year"].min())
    _year_hi = int(RAW_DF["year"].max())
else:
    years = []
    _hdr_sales = 0.0
    _hdr_profit = 0.0
    _hdr_gs = 0.0
    _hdr_margin = 0.0
    _year_lo = 0
    _year_hi = 0

filter_style = {
    "minWidth": "200px",
    "flex": "1 1 200px",
}

LABEL_UPPER = {
    "display": "block",
    "marginBottom": "8px",
    "color": TEXT_MUTED,
    "fontSize": "11px",
    "fontWeight": "600",
    "textTransform": "uppercase",
    "letterSpacing": "0.06em",
}

CHART_CARD = {
    "backgroundColor": CARD_BG,
    "border": f"1px solid {BORDER}",
    "borderRadius": "12px",
    "boxShadow": SHADOW,
    "padding": "12px 16px 16px",
    "overflow": "hidden",
}


def chart_wrap(graph):
    return html.Div(className="dash-chart-card", style=CHART_CARD, children=[graph])


def _main_layout():
    return html.Div(
        className="dash-main-layout",
        style={
            "backgroundColor": BG,
            "minHeight": "100vh",
            "color": TEXT,
            "fontFamily": "system-ui, -apple-system, sans-serif",
            "margin": 0,
            "padding": 0,
            "boxSizing": "border-box",
        },
        children=[
            dcc.Store(id="stored-data"),
            html.Div(
                className="dash-header-bar",
                style={
                    "width": "100%",
                    "backgroundColor": CARD_BG,
                    "borderBottom": f"1px solid {BORDER}",
                    "padding": "20px 28px 14px",
                    "boxSizing": "border-box",
                },
                children=[
                    html.Div(
                        className="dash-header-top",
                        style={
                            "display": "flex",
                            "flexWrap": "wrap",
                            "justifyContent": "space-between",
                            "alignItems": "center",
                            "gap": "20px",
                        },
                        children=[
                            html.Div(
                                children=[
                                    html.Div(
                                        "Financial Performance Report",
                                        style={
                                            "fontSize": "20px",
                                            "fontWeight": "700",
                                            "color": TEXT,
                                            "marginBottom": "6px",
                                            "lineHeight": "1.25",
                                        },
                                    ),
                                    html.Div(
                                        f"Sales · Profit · Segments · Countries | {_year_lo}–{_year_hi}",
                                        style={"fontSize": "12px", "color": TEXT_MUTED},
                                    ),
                                ]
                            ),
                            html.Div(
                                className="dash-header-badges",
                                style={
                                    "display": "flex",
                                    "flexWrap": "wrap",
                                    "gap": "10px",
                                    "alignItems": "center",
                                    "justifyContent": "center",
                                    "flex": "1 1 auto",
                                },
                                children=[
                                    html.Span(
                                        f"${_hdr_sales / 1e6:.2f}M Total Sales",
                                        style={
                                            "padding": "8px 14px",
                                            "borderRadius": "999px",
                                            "backgroundColor": "rgba(37, 99, 235, 0.12)",
                                            "color": TEXT,
                                            "fontSize": "13px",
                                            "fontWeight": "600",
                                        },
                                    ),
                                    html.Span(
                                        f"${_hdr_profit / 1e6:.2f}M Profit",
                                        style={
                                            "padding": "8px 14px",
                                            "borderRadius": "999px",
                                            "backgroundColor": "rgba(5, 150, 105, 0.12)",
                                            "color": TEXT,
                                            "fontSize": "13px",
                                            "fontWeight": "600",
                                        },
                                    ),
                                    html.Span(
                                        f"{_hdr_margin:.2f}% Margin",
                                        style={
                                            "padding": "8px 14px",
                                            "borderRadius": "999px",
                                            "backgroundColor": "rgba(124, 58, 237, 0.12)",
                                            "color": TEXT,
                                            "fontSize": "13px",
                                            "fontWeight": "600",
                                        },
                                    ),
                                ],
                            ),
                            html.Div(
                                children=[
                                    html.Div(
                                        "Built by Yash Sonawane | Python · Dash · Plotly",
                                        style={
                                            "fontSize": "11px",
                                            "color": TEXT_MUTED,
                                            "textAlign": "right",
                                            "fontWeight": "500",
                                        },
                                    )
                                ]
                            ),
                            html.Div(
                                [
                                    dcc.Upload(
                                        id="upload-data",
                                        children=html.Div([
                                            html.Span("↑", style={"fontSize": "14px", "marginRight": "6px"}),
                                            html.Span("Upload Dataset", style={"fontSize": "12px", "fontWeight": "600"}),
                                        ], style={"display": "flex", "alignItems": "center"}),
                                        style={
                                            "border": "1.5px solid #2563EB",
                                            "borderRadius": "8px",
                                            "padding": "6px 14px",
                                            "cursor": "pointer",
                                            "backgroundColor": "#EFF6FF",
                                            "color": "#2563EB",
                                            "display": "inline-flex",
                                            "alignItems": "center",
                                            "boxShadow": "0 1px 3px rgba(37,99,235,0.15)",
                                            "transition": "all 0.2s",
                                        },
                                        accept=".xlsx,.csv",
                                    ),
                                    html.Div(
                                        id="upload-status",
                                        style={
                                            "fontSize": "11px",
                                            "color": "#059669",
                                            "marginTop": "2px",
                                            "textAlign": "right",
                                        },
                                    ),
                                ],
                                style={"textAlign": "right"},
                            ),
                        ],
                    ),
                    html.Div(
                        "Data period: 2013–2015 | Last analyzed: March 2026",
                        style={
                            "marginTop": "10px",
                            "fontSize": "11px",
                            "color": TEXT_MUTED,
                        },
                    ),
                ],
            ),
            dcc.Loading(
                type="circle",
                color="#2563EB",
                children=[
                    html.Div(
                        className="dash-content",
                        style={"padding": "24px 28px 32px", "maxWidth": "1600px", "margin": "0 auto"},
                        children=[
                            html.Hr(style={"border": "none", "borderTop": "1px solid #E5E7EB", "margin": "8px 0"}),
                            html.Div([
                                html.Div([
                                    html.Span("📊 ", style={"fontSize": "16px"}),
                                    html.Span(f"{len(RAW_DF):,} transactions" if RAW_DF is not None else "—", style={"fontSize": "13px", "fontWeight": "600", "color": TEXT}),
                                    html.Span(" in dataset", style={"fontSize": "12px", "color": TEXT_MUTED}),
                                ], style={"display": "inline-flex", "alignItems": "center", "marginRight": "24px"}),
                                html.Div([
                                    html.Span("🗓 ", style={"fontSize": "16px"}),
                                    html.Span("2013–2015", style={"fontSize": "13px", "fontWeight": "600", "color": TEXT}),
                                    html.Span(" date range", style={"fontSize": "12px", "color": TEXT_MUTED}),
                                ], style={"display": "inline-flex", "alignItems": "center", "marginRight": "24px"}),
                                html.Div([
                                    html.Span("🌍 ", style={"fontSize": "16px"}),
                                    html.Span(f"{len(countries)} countries" if countries else "5 countries", style={"fontSize": "13px", "fontWeight": "600", "color": TEXT}),
                                    html.Span(f" · {len(segments)} segments" if segments else " · 5 segments", style={"fontSize": "12px", "color": TEXT_MUTED}),
                                ], style={"display": "inline-flex", "alignItems": "center", "marginRight": "24px"}),
                                html.Div([
                                    html.Span("📦 ", style={"fontSize": "16px"}),
                                    html.Span(f"{RAW_DF['product'].nunique()} products" if RAW_DF is not None else "—", style={"fontSize": "13px", "fontWeight": "600", "color": TEXT}),
                                    html.Span(" tracked", style={"fontSize": "12px", "color": TEXT_MUTED}),
                                ], style={"display": "inline-flex", "alignItems": "center"}),
                            ], style={"backgroundColor": CARD_BG, "borderRadius": "12px", "padding": "12px 20px", "marginBottom": "16px", "boxShadow": SHADOW, "display": "flex", "flexWrap": "wrap", "gap": "8px"}),
                            html.Div(
                                className="dash-filter-card",
                                style={
                                    "backgroundColor": CARD_BG,
                                    "border": f"1px solid {BORDER}",
                                    "borderRadius": "12px",
                                    "boxShadow": SHADOW,
                                    "padding": "20px 22px",
                                    "marginBottom": "20px",
                                },
                                children=[
                                    html.Div(
                                        className="dash-filter-row",
                                        style={"display": "flex", "flexWrap": "wrap", "gap": "20px"},
                                        children=[
                                            html.Div(
                                                style=filter_style,
                                                children=[
                                                    html.Label("Segment", style=LABEL_UPPER),
                                                    dcc.Dropdown(
                                                        id="filter-segment",
                                                        options=[{"label": "All segments", "value": "all"}]
                                                        + [{"label": s, "value": s} for s in segments],
                                                        value="all",
                                                        clearable=False,
                                                    ),
                                                ],
                                            ),
                                            html.Div(
                                                style=filter_style,
                                                children=[
                                                    html.Label("Country", style=LABEL_UPPER),
                                                    dcc.Dropdown(
                                                        id="filter-country",
                                                        options=[{"label": "All countries", "value": "all"}]
                                                        + [{"label": c, "value": c} for c in countries],
                                                        value="all",
                                                        clearable=False,
                                                    ),
                                                ],
                                            ),
                                            html.Div(
                                                style=filter_style,
                                                children=[
                                                    html.Label("Year", style=LABEL_UPPER),
                                                    dcc.Dropdown(
                                                        id="filter-year",
                                                        options=[{"label": "All years", "value": "all"}]
                                                        + [{"label": str(y), "value": y} for y in years],
                                                        value="all",
                                                        clearable=False,
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),
                                    html.Div(
                                        style={
                                            "marginTop": "16px",
                                            "display": "flex",
                                            "justifyContent": "flex-start",
                                        },
                                        children=[
                                            html.Button(
                                                "Download filtered data as CSV",
                                                id="btn-download",
                                                n_clicks=0,
                                                style={
                                                    "border": f"1px solid {BLUE}",
                                                    "backgroundColor": "transparent",
                                                    "color": BLUE,
                                                    "borderRadius": "999px",
                                                    "padding": "6px 14px",
                                                    "fontSize": "12px",
                                                    "fontWeight": "500",
                                                    "cursor": "pointer",
                                                },
                                            ),
                                            dcc.Download(id="download-data"),
                                        ],
                                    ),
                                ],
                            ),
                            html.Div(id="active-filters-bar", style={"marginBottom": "12px"}),
                            html.Hr(style={"border": "none", "borderTop": "1px solid #E5E7EB", "margin": "8px 0"}),
                            html.Div(
                                id="insights-card",
                                style={
                                    "backgroundColor": CARD_BG,
                                    "borderRadius": "12px",
                                    "padding": "20px",
                                    "marginBottom": "20px",
                                    "boxShadow": SHADOW,
                                },
                            ),
                            html.Div(
                                id="kpi-container",
                                style={
                                    "display": "grid",
                                    "gridTemplateColumns": "repeat(auto-fit, minmax(160px, 1fr))",
                                    "gap": "16px",
                                    "marginBottom": "24px",
                                },
                            ),
                            html.Hr(style={"border": "none", "borderTop": "1px solid #E5E7EB", "margin": "8px 0"}),
                            html.Div(
                                className="dash-chart-grid-2",
                                style={
                                    "display": "grid",
                                    "gridTemplateColumns": "repeat(auto-fit, minmax(300px, 1fr))",
                                    "gap": "20px",
                                    "marginBottom": "20px",
                                },
                                children=[
                                    html.Div(className="dash-chart-card", style=CHART_CARD, title="Business question: Is our business growing profitably or just growing revenue?", children=[dcc.Graph(id="chart-sales-profit-year", config=PLOT_CONFIG, style={"height": "380px"})]),
                                    html.Div(className="dash-chart-card", style=CHART_CARD, title="Business question: Is our efficiency improving over time?", children=[dcc.Graph(id="chart-margin-trend", config=PLOT_CONFIG, style={"height": "380px"})]),
                                ],
                            ),
                            html.Div(
                                className="dash-chart-grid-3",
                                style={
                                    "display": "grid",
                                    "gridTemplateColumns": "repeat(auto-fit, minmax(300px, 1fr))",
                                    "gap": "20px",
                                    "marginBottom": "20px",
                                },
                                children=[
                                    html.Div(className="dash-chart-card", style={**CHART_CARD, "minHeight": "360px"}, title="Business question: Which customer segment drives the most revenue?", children=[
                                        html.P("Sales by Segment", style={"fontSize": "13px", "fontWeight": "600", "color": TEXT, "marginBottom": "16px"}),
                                        html.Div(id="segment-bars", style={"padding": "8px 0"}),
                                    ]),
                                    html.Div(className="dash-chart-card", style=CHART_CARD, title="Business question: Which market is most profitable?", children=[dcc.Graph(id="chart-profit-country", config=PLOT_CONFIG, style={"height": "360px"})]),
                                    html.Div(className="dash-chart-card", style=CHART_CARD, title="Business question: Which products should we invest more in?", children=[dcc.Graph(id="chart-top-products", config=PLOT_CONFIG, style={"height": "360px"})]),
                                ],
                            ),
                            html.Div(
                                className="dash-chart-grid-2",
                                style={
                                    "display": "grid",
                                    "gridTemplateColumns": "repeat(auto-fit, minmax(300px, 1fr))",
                                    "gap": "20px",
                                    "marginBottom": "8px",
                                },
                                children=[
                                    html.Div(className="dash-chart-card", style=CHART_CARD, title="Business question: When are our peak sales months?", children=[dcc.Graph(id="chart-monthly-sales", config=PLOT_CONFIG, style={"height": "380px"})]),
                                    html.Div(className="dash-chart-card", style=CHART_CARD, title="Business question: Does discounting hurt our profitability?", children=[dcc.Graph(id="chart-scatter-discount", config=PLOT_CONFIG, style={"height": "380px"})]),
                                ],
                            ),
                            html.Div([
                                html.P("RECONCILIATION VARIANCE — TOP 5 PRODUCTS", style={"fontWeight": "600", "fontSize": "11px", "color": TEXT_MUTED, "letterSpacing": "0.06em", "marginBottom": "12px"}),
                                html.P("Products with largest gap between gross sales and profit — high variance signals cost or discount pressure", style={"fontSize": "12px", "color": TEXT_MUTED, "marginBottom": "12px"}),
                                html.Div(id="variance-table"),
                            ], style={"backgroundColor": CARD_BG, "borderRadius": "12px", "padding": "20px", "marginTop": "20px", "boxShadow": SHADOW}),
                            html.Div(
                                "Financial Performance Dashboard · Built with Python, Dash & Plotly · github.com/Sonawane250398/financial-dashboard",
                                style={
                                    "textAlign": "center",
                                    "color": "#9CA3AF",
                                    "fontSize": "12px",
                                    "padding": "24px 8px 0",
                                },
                            ),
                        ],
                    )
                ],
            ),
        ],
    )


if DATA_ERROR is not None:
    app.layout = html.Div(
        DATA_ERROR,
        style={"color": "red", "padding": "20px"},
    )
else:
    app.layout = _main_layout()


def kpi_card(title: str, value: str, accent: str, trend=None, pct_change=None) -> html.Div:
    if trend and pct_change is not None:
        arrow = "↑" if trend == "+" else "↓"
        sign = "+" if trend == "+" else ""
        pct_class = "up" if trend == "+" else "down"
        trend_el = html.Span(
            f"{arrow} {sign}{pct_change:.1f}%",
            className=f"kpi-pct {pct_class}",
        )
    elif trend:
        arrow = "↑" if trend == "+" else "↓"
        pct_class = "up" if trend == "+" else "down"
        trend_el = html.Span(arrow, className=f"kpi-pct {pct_class}")
    else:
        trend_el = html.Span()
    return html.Div(
        style={
            "backgroundColor": CARD_BG,
            "border": f"1px solid {BORDER}",
            "borderRadius": "12px",
            "boxShadow": SHADOW,
            "borderLeft": f"4px solid {accent}",
            "padding": "20px 22px",
        },
        children=[
            html.Div(
                style={"display": "flex", "alignItems": "center", "gap": "8px", "marginBottom": "10px"},
                children=[
                    html.Div(
                        style={
                            "width": "10px",
                            "height": "10px",
                            "backgroundColor": accent,
                            "borderRadius": "2px",
                            "flexShrink": 0,
                        }
                    ),
                    html.Div(
                        title,
                        style={
                            "color": TEXT_MUTED,
                            "fontSize": "11px",
                            "fontWeight": "600",
                            "textTransform": "uppercase",
                            "letterSpacing": "0.06em",
                        },
                    ),
                ],
            ),
            html.Div(
                style={"display": "flex", "alignItems": "baseline"},
                children=[
                    html.Span(value, style={"fontSize": "28px", "fontWeight": "700", "color": TEXT, "lineHeight": "1.15"}),
                    trend_el,
                ],
            ),
        ],
    )


@app.callback(
    Output("stored-data", "data"),
    Output("upload-status", "children"),
    Output("upload-status", "style"),
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    prevent_initial_call=True,
)
def handle_upload(contents, filename):
    if contents is None:
        return (
            None,
            "Using default Financial Sample dataset",
            {
                "fontSize": "12px",
                "color": "#6b7280",
                "textAlign": "center",
                "marginBottom": "4px",
            },
        )
    try:
        content_type, content_string = contents.split(",")
        decoded = base64.b64decode(content_string)
        if filename.endswith(".csv"):
            df = pd.read_csv(io.StringIO(decoded.decode("utf-8")))
        else:
            df = pd.read_excel(io.BytesIO(decoded))
        df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
        return (
            df.to_json(date_format="iso", orient="split"),
            f"Using your dataset: {filename}",
            {
                "fontSize": "12px",
                "color": "#059669",
                "textAlign": "center",
                "marginBottom": "4px",
                "fontWeight": "500",
            },
        )
    except Exception:
        return (
            None,
            "Error reading file — check column names match: Date, Sales, Profit, Segment, Country, Product",
            {
                "fontSize": "12px",
                "color": "#DC2626",
                "textAlign": "center",
                "marginBottom": "4px",
            },
        )


@app.callback(
    Output("download-data", "data"),
    Output("kpi-container", "children"),
    Output("chart-sales-profit-year", "figure"),
    Output("chart-margin-trend", "figure"),
    Output("segment-bars", "children"),
    Output("chart-profit-country", "figure"),
    Output("chart-top-products", "figure"),
    Output("chart-monthly-sales", "figure"),
    Output("chart-scatter-discount", "figure"),
    Output("insights-card", "children"),
    Output("variance-table", "children"),
    Output("active-filters-bar", "children"),
    Input("stored-data", "data"),
    Input("filter-segment", "value"),
    Input("filter-country", "value"),
    Input("filter-year", "value"),
    Input("btn-download", "n_clicks"),
)
def update_dashboard(stored_data, segment, country, year_val, n_clicks):
    if stored_data is not None:
        d_base = pd.read_json(stored_data, orient="split")
    elif RAW_DF is not None:
        d_base = RAW_DF.copy()
    else:
        d_base = None

    if d_base is None:
        d = pd.DataFrame()
    else:
        d = apply_filters(d_base, segment, country, year_val)

    total_sales = d["sales"].sum() if "sales" in d.columns else 0
    total_profit = d["profit"].sum() if "profit" in d.columns else 0
    total_gs = d["gross_sales"].sum() if "gross_sales" in d.columns else 0
    margin_agg = (total_profit / total_gs * 100.0) if total_gs else 0.0
    units = d["units_sold"].sum() if "units_sold" in d.columns else 0

    if year_val is not None and year_val != "all" and d_base is not None:
        d_prev = apply_filters(d_base, segment, country, int(year_val) - 1)
        prev_sales = d_prev["sales"].sum() if "sales" in d_prev.columns else 0
        prev_profit = d_prev["profit"].sum() if "profit" in d_prev.columns else 0
        prev_gs = d_prev["gross_sales"].sum() if "gross_sales" in d_prev.columns else 0
        prev_margin = (prev_profit / prev_gs * 100.0) if prev_gs else 0.0
        prev_units = d_prev["units_sold"].sum() if "units_sold" in d_prev.columns else 0
        trend_sales = "+" if total_sales > prev_sales else ("-" if total_sales < prev_sales else None)
        trend_profit = "+" if total_profit > prev_profit else ("-" if total_profit < prev_profit else None)
        trend_margin = "+" if margin_agg > prev_margin else ("-" if margin_agg < prev_margin else None)
        trend_units = "+" if units > prev_units else ("-" if units < prev_units else None)
        pct_sales   = ((total_sales  - prev_sales)   / prev_sales   * 100) if prev_sales   else None
        pct_profit  = ((total_profit - prev_profit)  / prev_profit  * 100) if prev_profit  else None
        pct_margin  = (margin_agg - prev_margin)
        pct_units   = ((units - prev_units) / prev_units * 100) if prev_units else None
    else:
        trend_sales = trend_profit = trend_margin = trend_units = None
        pct_sales = pct_profit = pct_margin = pct_units = None

    kpis = [
        kpi_card("Total sales",     f"${total_sales / 1e6:.2f}M", BLUE,   trend_sales,   pct_sales),
        kpi_card("Total profit",    f"${total_profit / 1e6:.2f}M", GREEN,  trend_profit,  pct_profit),
        kpi_card("Profit margin %", f"{margin_agg:.2f}%",          PURPLE, trend_margin,  pct_margin),
        kpi_card("Units sold",      f"{units:,.0f}",                AMBER,  trend_units,   pct_units),
    ]

    def _empty_fig():
        fig = go.Figure()
        plot_layout(fig)
        fig.add_annotation(
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            text="No data matches the selected filters — try adjusting your selections",
            showarrow=False,
            font=dict(size=12, color=TEXT_MUTED),
        )
        return fig

    if d.empty:
        kpis_zero = [
            kpi_card("Total sales", "$0", BLUE),
            kpi_card("Total profit", "$0", GREEN),
            kpi_card("Profit margin %", "0", PURPLE),
            kpi_card("Units sold", "0", AMBER),
        ]
        empty_fig = _empty_fig()
        return (
            None,
            kpis_zero,
            empty_fig,
            empty_fig,
            [],
            empty_fig,
            empty_fig,
            empty_fig,
            empty_fig,
            [],
            [],
            [],
        )
    # --- Sales vs Profit by Year (grouped bars) ---
    sp_data = d[d["year"] != 2016] if 2016 in d["year"].values and d[d["year"] == 2016].shape[0] < 100 else d
    by_year = sp_data.groupby("year", as_index=False).agg(sales=("sales", "sum"), profit=("profit", "sum"))
    by_year = by_year.sort_values("year")
    fig_sp = go.Figure(
        data=[
            go.Bar(name="Sales", x=by_year["year"], y=by_year["sales"], marker_color=BLUE, hovertemplate="<b>%{x}</b><br>$%{y:,.2f}<extra></extra>"),
            go.Bar(name="Profit", x=by_year["year"], y=by_year["profit"], marker_color=GREEN, hovertemplate="<b>%{x}</b><br>$%{y:,.2f}<extra></extra>"),
        ]
    )
    fig_sp.update_layout(barmode="group", title="Sales vs Profit by Year (2013–2015)")
    fig_sp.update_xaxes(tickmode="array", tickvals=sorted(by_year["year"].unique().tolist()), ticktext=[str(y) for y in sorted(by_year["year"].unique().tolist())])
    peak_year_mask = by_year["year"] == 2015
    if peak_year_mask.any():
        peak_sales = float(by_year.loc[peak_year_mask, "sales"].iloc[0])
        fig_sp.add_annotation(
            x=2015,
            y=peak_sales,
            text="Peak sales year",
            showarrow=True,
            arrowhead=2,
            arrowsize=1,
            arrowwidth=1,
            arrowcolor=TEXT_MUTED,
            ax=0,
            ay=-40,
            bgcolor="rgba(255,255,255,0.9)",
            bordercolor=BORDER,
            borderwidth=1,
            font=dict(size=11, color=TEXT),
        )
    plot_layout(fig_sp)

    # --- Profit margin % trend by year (all years including 2016) ---
    margin_by_year = (
        d.groupby("year", as_index=False)
        .agg(profit=("profit", "sum"), gross_sales=("gross_sales", "sum"))
    )
    margin_by_year["margin"] = (margin_by_year["profit"] / margin_by_year["gross_sales"] * 100).round(2)
    margin_by_year = margin_by_year[["year", "margin"]].sort_values("year")
    fig_margin = go.Figure(
        data=[
            go.Scatter(
                x=margin_by_year["year"],
                y=margin_by_year["margin"],
                mode="lines+markers",
                line=dict(color=PURPLE, width=2.5),
                marker=dict(size=8, color=AMBER, line=dict(width=0)),
                name="Margin %",
                hovertemplate="<b>%{x}</b><br>Margin: %{y:.1f}%<extra></extra>",
            )
        ]
    )
    fig_margin.update_layout(title="Profit margin % trend by year", yaxis_title="Margin %")
    fig_margin.update_xaxes(tickmode="array", tickvals=sorted(margin_by_year["year"].unique().tolist()), ticktext=[str(y) for y in sorted(margin_by_year["year"].unique().tolist())])
    if not margin_by_year.empty:
        idx_max = margin_by_year["margin"].idxmax()
        year_max = int(margin_by_year.loc[idx_max, "year"])
        margin_max = float(margin_by_year.loc[idx_max, "margin"])
        fig_margin.add_annotation(
            x=year_max,
            y=margin_max,
            text=f"{margin_max:.0f}% — highest margin",
            showarrow=True,
            arrowhead=2,
            arrowsize=1,
            arrowwidth=1,
            arrowcolor=TEXT_MUTED,
            ax=0,
            ay=-40,
            bgcolor="rgba(255,255,255,0.9)",
            bordercolor=BORDER,
            borderwidth=1,
            font=dict(size=11, color=TEXT),
        )
    plot_layout(fig_margin)

    # --- Sales by Segment (progress bars — always show all segments, respects country/year but not segment filter) ---
    _seg_base = d_base if d_base is not None else pd.DataFrame()
    d_all_segs = apply_filters(_seg_base, "all", country, year_val) if not _seg_base.empty else pd.DataFrame()
    seg_data = d_all_segs.groupby("segment")["sales"].sum().sort_values(ascending=False).reset_index()
    max_sales = seg_data["sales"].max() if not seg_data.empty else 1
    seg_colors = ["#2563EB", "#059669", "#7C3AED", "#D97706", "#DC2626"]
    segment_bars = [
        html.Div([
            html.Div([
                html.Span(
                    row["segment"],
                    style={
                        "fontSize": "13px",
                        "fontWeight": "700" if (segment and segment != "all" and row["segment"] == segment) else "500",
                        "color": BLUE if (segment and segment != "all" and row["segment"] == segment) else TEXT,
                    },
                ),
                html.Span(f"${row['sales'] / 1e6:.1f}M", style={"fontSize": "13px", "color": TEXT_MUTED, "float": "right"}),
            ], style={"marginBottom": "4px", "overflow": "hidden"}),
            html.Div(html.Div(style={
                "width": f"{(row['sales'] / max_sales) * 100:.0f}%",
                "height": "8px",
                "borderRadius": "4px",
                "backgroundColor": BLUE if (segment and segment != "all" and row["segment"] == segment) else seg_colors[i % len(seg_colors)],
                "transition": "width 0.6s ease",
            }), style={"backgroundColor": "#F3F4F6", "borderRadius": "4px", "marginBottom": "12px"}),
        ])
        for i, row in seg_data.iterrows()
    ]

    # --- Profit by Country (horizontal) ---
    by_ct = d.groupby("country", as_index=False)["profit"].sum().sort_values("profit", ascending=True)
    by_ct["profit_m"] = by_ct["profit"] / 1_000_000
    fig_ct = go.Figure(data=[go.Bar(x=by_ct["profit_m"], y=by_ct["country"], orientation="h", marker_color=GREEN, hovertemplate="<b>%{y}</b><br>Profit: $%{x:.1f}M<extra></extra>")])
    fig_ct.update_layout(title="Profit by country", xaxis_title="Profit ($M)", yaxis_title="")
    plot_layout(fig_ct)
    fig_ct.update_layout(showlegend=False)

    # --- Top 6 products by profit ---
    by_prod = d.groupby("product", as_index=False)["profit"].sum().nlargest(6, "profit").sort_values("profit", ascending=True)
    fig_prod = go.Figure(data=[go.Bar(x=by_prod["profit"], y=by_prod["product"], orientation="h", marker_color=PURPLE)])
    fig_prod.update_layout(title="Top 6 products by profit", xaxis_title="Profit ($)", yaxis_title="")
    plot_layout(fig_prod)
    fig_prod.update_layout(showlegend=False)

    # --- Monthly sales trend (area) ---
    monthly = d.groupby("year_month", as_index=False)["sales"].sum()
    monthly["_sort"] = pd.PeriodIndex(monthly["year_month"], freq="M")
    monthly = monthly.sort_values("_sort")
    fig_mo = go.Figure(
        data=[
            go.Scatter(
                x=monthly["year_month"],
                y=monthly["sales"],
                mode="lines",
                fill="tozeroy",
                fillcolor="rgba(37, 99, 235, 0.1)",
                line=dict(color=BLUE, width=2),
                name="Sales",
                hovertemplate="<b>%{x}</b><br>Sales: $%{y:,.0f}<extra></extra>",
            )
        ]
    )
    fig_mo.update_layout(title="Monthly sales trend", xaxis_title="Month", yaxis_title="Sales ($)")
    plot_layout(fig_mo)
    fig_mo.update_xaxes(tickangle=-35)

    # --- Scatter: Discount vs Profit margin, bubble size = sales ---
    sub = d.sample(min(500, len(d)), random_state=1) if len(d) > 500 else d
    sizes = sub["sales"].clip(lower=1)
    s_norm = 4 + (sizes / sizes.max()) * 18
    fig_sc = go.Figure(
        data=[
            go.Scatter(
                x=sub["discounts"],
                y=sub["profit_margin_pct"],
                mode="markers",
                showlegend=False,
                marker=dict(
                    size=s_norm,
                    sizemode="diameter",
                    sizemin=4,
                    color=GREEN,
                    opacity=0.5,
                    line=dict(width=0),
                ),
                text=sub["product"],
                customdata=sub["sales"],
                hovertemplate="<b>%{text}</b><br>Discount: $%{x:,.0f}<br>Margin: %{y:.2f}%<br>Sales: $%{customdata:,.0f}<extra></extra>",
            )
        ]
    )
    fig_sc.update_layout(title="Discount vs profit margin", xaxis_title="Discount ($)", yaxis_title="Profit margin %")
    plot_layout(fig_sc)

    download_data = None
    ctx = dash.callback_context
    if ctx.triggered and ctx.triggered[0]["prop_id"].startswith("btn-download") and not d.empty:
        download_data = dcc.send_data_frame(d.to_csv, "financial_export.csv", index=False)

    # --- Active filter pills ---
    active_pills = []
    if segment and segment != "all":
        active_pills.append(html.Span(f"Segment: {segment} ×", className="filter-active-pill"))
    if country and country != "all":
        active_pills.append(html.Span(f"Country: {country} ×", className="filter-active-pill"))
    if year_val and year_val != "all":
        active_pills.append(html.Span(f"Year: {year_val} ×", className="filter-active-pill"))
    if active_pills:
        active_filters_bar = html.Div(
            [html.Span("Active filters: ", style={"fontSize": "12px", "color": TEXT_MUTED, "marginRight": "8px", "fontWeight": "500"})] + active_pills,
            style={"display": "flex", "flexWrap": "wrap", "alignItems": "center", "gap": "6px",
                   "backgroundColor": "#EFF6FF", "borderRadius": "8px", "padding": "8px 14px",
                   "border": "1px solid rgba(37,99,235,0.2)"},
        )
    else:
        active_filters_bar = html.Div(
            html.Span("Showing all data — use filters above to drill down", style={"fontSize": "12px", "color": TEXT_MUTED}),
            style={"padding": "6px 0"},
        )

    # --- Dynamic insights ---
    top_segment = d.groupby("segment")["sales"].sum().idxmax() if not d.empty else "N/A"
    top_country = d.groupby("country")["profit"].sum().idxmax() if not d.empty else "N/A"
    avg_margin = (d["profit"].sum() / d["gross_sales"].sum() * 100) if not d.empty and d["gross_sales"].sum() > 0 else 0
    best_year = d.groupby("year").apply(lambda x: x["profit"].sum() / x["gross_sales"].sum() * 100).idxmax() if not d.empty and len(d["year"].unique()) > 1 else "N/A"

    insight_items = [
        ("blue",   f"Top segment: {top_segment} leads in total sales revenue"),
        ("green",  f"Best market: {top_country} generates the highest absolute profit"),
        ("purple", f"Avg margin: {avg_margin:.1f}% across the selected period"),
        ("amber",  f"Best year: {best_year} showed the strongest profitability ratio"),
    ]
    color_map = {"blue": BLUE, "green": GREEN, "purple": PURPLE, "amber": AMBER}
    insights_children = [
        html.P("KEY FINDINGS", style={"fontWeight": "600", "fontSize": "11px", "color": TEXT_MUTED, "letterSpacing": "0.06em", "marginBottom": "14px"}),
        html.Div([
            html.Div([
                html.Span(style={"width": "8px", "height": "8px", "borderRadius": "50%",
                                 "backgroundColor": color_map[c], "flexShrink": 0, "marginTop": "3px"}),
                html.Span(text, style={"fontSize": "13px", "color": TEXT, "lineHeight": "1.45"}),
            ], className=f"insight-item {'green' if c == 'green' else 'purple' if c == 'purple' else 'amber' if c == 'amber' else ''}")
            for c, text in insight_items
        ]),
    ]

    # --- Reconciliation variance table ---
    variance_df = d.groupby("product").agg(gross_sales=("gross_sales", "sum"), profit=("profit", "sum")).reset_index()
    variance_df["variance"] = variance_df["gross_sales"] - variance_df["profit"]
    variance_df["margin_pct"] = (variance_df["profit"] / variance_df["gross_sales"] * 100).round(1)
    variance_df = variance_df.nlargest(5, "variance")[["product", "gross_sales", "profit", "variance", "margin_pct"]]
    variance_df.columns = ["Product", "Gross Sales", "Profit", "Variance", "Margin %"]
    variance_df["Gross Sales"] = variance_df["Gross Sales"].apply(lambda x: f"${x:,.0f}")
    variance_df["Profit"] = variance_df["Profit"].apply(lambda x: f"${x:,.0f}")
    variance_df["Variance"] = variance_df["Variance"].apply(lambda x: f"${x:,.0f}")
    variance_df["Margin %"] = variance_df["Margin %"].apply(lambda x: f"{x}%")

    def variance_color(val_str):
        val = float(val_str.replace("$", "").replace(",", ""))
        if val > 5_000_000:
            return {"backgroundColor": "#FEF2F2", "color": "#DC2626", "fontWeight": "600"}
        elif val > 1_000_000:
            return {"backgroundColor": "#FFFBEB", "color": "#D97706", "fontWeight": "600"}
        return {"backgroundColor": "#F0FDF4", "color": "#059669", "fontWeight": "600"}

    def row_td(i, col):
        base = {"padding": "8px 12px", "fontSize": "13px", "borderBottom": f"1px solid {BORDER}"}
        if col == "Variance":
            return html.Td(variance_df.iloc[i][col], style={**base, **variance_color(variance_df.iloc[i][col])})
        return html.Td(variance_df.iloc[i][col], style={**base, "color": TEXT})

    table = html.Table([
        html.Thead(html.Tr([
            html.Th(col, style={"padding": "8px 12px", "textAlign": "left", "fontSize": "11px", "color": TEXT_MUTED, "borderBottom": f"2px solid {BORDER}", "fontWeight": "600", "letterSpacing": "0.04em"})
            for col in variance_df.columns
        ])),
        html.Tbody([
            html.Tr([
                row_td(i, col) for col in variance_df.columns
            ], style={"backgroundColor": CARD_BG if i % 2 == 0 else "#F9FAFB"})
            for i in range(len(variance_df))
        ]),
    ], style={"width": "100%", "borderCollapse": "collapse"})

    return download_data, kpis, fig_sp, fig_margin, segment_bars, fig_ct, fig_prod, fig_mo, fig_sc, insights_children, table, active_filters_bar


server = app.server


import threading
import urllib.request


def keep_alive():
    import time
    while True:
        time.sleep(840)  # ping every 14 minutes
        try:
            urllib.request.urlopen("https://financial-dashboard-bu2r.onrender.com")
            print("Keep-alive ping sent")
        except Exception as e:
            print(f"Keep-alive ping failed: {e}")


t = threading.Thread(target=keep_alive, daemon=True)
t.start()

if __name__ == "__main__":
    app.run(debug=False)
