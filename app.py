"""
Financial analytics dashboard — Dash + Plotly.
Run: python3 app.py  →  http://127.0.0.1:8050
"""

import re

import dash
from dash import dcc, html, Input, Output
import numpy as np
import pandas as pd
import plotly.graph_objects as go

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
SHADOW = "0 1px 4px rgba(0,0,0,0.08)"
PALETTE = [BLUE, GREEN, PURPLE, AMBER]


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


RAW_DF = load_and_prepare()


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

app = dash.Dash(__name__)
app.title = "Financial Performance Report"

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
if "</head>" in app.index_string:
    app.index_string = app.index_string.replace("</head>", _DROPDOWN_CSS + "\n</head>", 1)

segments = sorted(RAW_DF["segment"].dropna().unique().tolist())
countries = sorted(RAW_DF["country"].dropna().unique().tolist())
year_counts = RAW_DF["year"].value_counts()
years = sorted(year_counts[year_counts >= 100].index.tolist())

_hdr_sales = RAW_DF["sales"].sum()
_hdr_profit = RAW_DF["profit"].sum()
_hdr_gs = RAW_DF["gross_sales"].sum()
_hdr_margin = (_hdr_profit / _hdr_gs * 100.0) if _hdr_gs else 0.0
_year_lo = int(RAW_DF["year"].min())
_year_hi = int(RAW_DF["year"].max())

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
    return html.Div(style=CHART_CARD, children=[graph])


app.layout = html.Div(
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
        html.Div(
            style={
                "width": "100%",
                "backgroundColor": CARD_BG,
                "borderBottom": f"1px solid {BORDER}",
                "padding": "20px 28px",
                "boxSizing": "border-box",
            },
            children=[
                html.Div(
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
                            style={"display": "flex", "flexWrap": "wrap", "gap": "10px", "alignItems": "center"},
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
                    ],
                ),
            ],
        ),
        html.Div(
            style={"padding": "24px 28px 32px", "maxWidth": "1600px", "margin": "0 auto"},
            children=[
                html.Div(
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
                            style={"display": "flex", "flexWrap": "wrap", "gap": "20px"},
                            children=[
                                html.Div(
                                    style=filter_style,
                                    children=[
                                        html.Label("Segment", style=LABEL_UPPER),
                                        dcc.Dropdown(
                                            id="filter-segment",
                                            options=[{"label": "All segments", "value": "all"}] + [{"label": s, "value": s} for s in segments],
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
                                            options=[{"label": "All countries", "value": "all"}] + [{"label": c, "value": c} for c in countries],
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
                                            options=[{"label": "All years", "value": "all"}] + [{"label": str(y), "value": y} for y in years],
                                            value="all",
                                            clearable=False,
                                        ),
                                    ],
                                ),
                            ],
                        ),
                    ],
                ),
                html.Div(
                    id="kpi-row",
                    style={"display": "grid", "gridTemplateColumns": "repeat(4, 1fr)", "gap": "16px", "marginBottom": "24px"},
                ),
                html.Div(
                    style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "20px", "marginBottom": "20px"},
                    children=[
                        chart_wrap(dcc.Graph(id="chart-sales-profit-year", config=PLOT_CONFIG, style={"height": "380px"})),
                        chart_wrap(dcc.Graph(id="chart-margin-trend", config=PLOT_CONFIG, style={"height": "380px"})),
                    ],
                ),
                html.Div(
                    style={"display": "grid", "gridTemplateColumns": "1fr 1fr 1fr", "gap": "20px", "marginBottom": "20px"},
                    children=[
                        chart_wrap(dcc.Graph(id="chart-sales-segment", config=PLOT_CONFIG, style={"height": "360px"})),
                        chart_wrap(dcc.Graph(id="chart-profit-country", config=PLOT_CONFIG, style={"height": "360px"})),
                        chart_wrap(dcc.Graph(id="chart-top-products", config=PLOT_CONFIG, style={"height": "360px"})),
                    ],
                ),
                html.Div(
                    style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "20px", "marginBottom": "8px"},
                    children=[
                        chart_wrap(dcc.Graph(id="chart-monthly-sales", config=PLOT_CONFIG, style={"height": "380px"})),
                        chart_wrap(dcc.Graph(id="chart-scatter-discount", config=PLOT_CONFIG, style={"height": "380px"})),
                    ],
                ),
                html.Div(
                    "Built by Yash Sonawane · Financial Analytics Dashboard · Python · Dash · Plotly",
                    style={
                        "textAlign": "center",
                        "color": "#9CA3AF",
                        "fontSize": "12px",
                        "padding": "24px 8px 0",
                    },
                ),
            ],
        ),
    ],
)


def kpi_card(title: str, value: str, accent: str) -> html.Div:
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
            html.Div(value, style={"fontSize": "28px", "fontWeight": "700", "color": TEXT, "lineHeight": "1.15"}),
        ],
    )


@app.callback(
    Output("kpi-row", "children"),
    Output("chart-sales-profit-year", "figure"),
    Output("chart-margin-trend", "figure"),
    Output("chart-sales-segment", "figure"),
    Output("chart-profit-country", "figure"),
    Output("chart-top-products", "figure"),
    Output("chart-monthly-sales", "figure"),
    Output("chart-scatter-discount", "figure"),
    Input("filter-segment", "value"),
    Input("filter-country", "value"),
    Input("filter-year", "value"),
)
def update_dashboard(segment, country, year_val):
    d = apply_filters(RAW_DF, segment, country, year_val)

    total_sales = d["sales"].sum()
    total_profit = d["profit"].sum()
    total_gs = d["gross_sales"].sum()
    margin_agg = (total_profit / total_gs * 100.0) if total_gs else 0.0
    units = d["units_sold"].sum()

    kpis = [
        kpi_card("Total sales", f"${total_sales / 1e6:.2f}M", BLUE),
        kpi_card("Total profit", f"${total_profit / 1e6:.2f}M", GREEN),
        kpi_card("Profit margin %", f"{margin_agg:.2f}%", PURPLE),
        kpi_card("Units sold", f"{units:,.0f}", AMBER),
    ]

    if d.empty:
        empty = go.Figure()
        plot_layout(empty)
        empty.update_layout(title=dict(text="No data for current filters"))
        return kpis, empty, empty, empty, empty, empty, empty, empty

    # --- Sales vs Profit by Year (grouped bars) ---
    by_year = d.groupby("year", as_index=False).agg(sales=("sales", "sum"), profit=("profit", "sum"))
    by_year = by_year.sort_values("year")
    fig_sp = go.Figure(
        data=[
            go.Bar(name="Sales", x=by_year["year"], y=by_year["sales"], marker_color=BLUE),
            go.Bar(name="Profit", x=by_year["year"], y=by_year["profit"], marker_color=GREEN),
        ]
    )
    fig_sp.update_layout(barmode="group", title="Sales vs profit by year")
    plot_layout(fig_sp)

    # --- Profit margin % trend by year ---
    gy = d.groupby("year").agg(p=("profit", "sum"), g=("gross_sales", "sum")).reset_index()
    gy["margin_pct"] = np.where(gy["g"] != 0, gy["p"] / gy["g"] * 100.0, 0.0)
    gy = gy.sort_values("year")
    fig_margin = go.Figure(
        data=[
            go.Scatter(
                x=gy["year"],
                y=gy["margin_pct"],
                mode="lines+markers",
                line=dict(color=PURPLE, width=2.5),
                marker=dict(size=8, color=AMBER, line=dict(width=0)),
                name="Margin %",
            )
        ]
    )
    fig_margin.update_layout(title="Profit margin % trend by year", yaxis_title="Margin %")
    plot_layout(fig_margin)

    # --- Sales by Segment (vertical bars) ---
    by_seg = d.groupby("segment", as_index=False)["sales"].sum().sort_values("sales", ascending=False)
    fig_seg = go.Figure(data=[go.Bar(x=by_seg["segment"], y=by_seg["sales"], marker_color=BLUE)])
    fig_seg.update_layout(title="Sales by segment", xaxis_title="", yaxis_title="Sales ($)")
    plot_layout(fig_seg)
    fig_seg.update_layout(showlegend=False)
    fig_seg.update_xaxes(tickangle=-25)

    # --- Profit by Country (horizontal) ---
    by_ct = d.groupby("country", as_index=False)["profit"].sum().sort_values("profit", ascending=True)
    fig_ct = go.Figure(data=[go.Bar(x=by_ct["profit"], y=by_ct["country"], orientation="h", marker_color=GREEN)])
    fig_ct.update_layout(title="Profit by country", xaxis_title="Profit ($)", yaxis_title="")
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

    return kpis, fig_sp, fig_margin, fig_seg, fig_ct, fig_prod, fig_mo, fig_sc


if __name__ == "__main__":
    app.run(debug=True)
