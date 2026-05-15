"""Trend chart generation for the FinOps Report Generator.

Builds 6-month line charts using Plotly and exports them as PNG images
via kaleido for embedding into PowerPoint slides.
"""

from __future__ import annotations

import logging
from typing import Any

import plotly.graph_objects as go

from src.pptx_utils import fmt_abbreviated as _fmt_abbreviated

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Color constants
# ---------------------------------------------------------------------------

ACTUAL_COLOR = "#0099FF"   # Azure (AWN brand)
FORECAST_COLOR = "#FF730F"  # Sunset (AWN brand)

# Stacked bar palette — Arctic Wolf brand colors ordered for contrast
STACKED_BAR_PALETTE = [
    "#0099FF",  # Azure
    "#FF730F",  # Sunset
    "#83D6FF",  # Mist
    "#F83300",  # Blood Moon
    "#0066FF",  # Aurora
    "#D7DADE",  # Steel 15%
    "#0059B2",  # Sea
    "#8D949C",  # Steel 45%
    "#66CCFF",  # Light Azure
    "#FF9944",  # Light Sunset
    "#3A4958",  # Steel 85% (for "Other")
]

# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------




# ---------------------------------------------------------------------------
# Chart builders
# ---------------------------------------------------------------------------


def build_trend_chart(
    bucket_name: str,
    history: list[dict[str, Any]],
    *,
    dark_mode: bool = True,
) -> go.Figure:
    """Build a Plotly line chart showing 6-month actuals and forecast.

    Args:
        bucket_name: The bucket name used as chart title.
        history: Chronological list of dicts with keys ``month_label`` (str),
            ``actual`` (float), and ``forecast`` (float | None).
        dark_mode: If True, uses white text on transparent background
            (for dark PPTX slides). If False, uses dark text on white
            background (for light-themed output).

    Returns:
        A Plotly Figure object.
    """
    months = [h["month_label"] for h in history]
    actuals = [h["actual"] for h in history]
    forecasts = [h.get("forecast") for h in history]

    fig = go.Figure()

    # Compute abbreviated y-axis tick values so they match the data labels
    all_values = list(actuals) + [f for f in forecasts if f is not None]
    if all_values:
        y_min, y_max = min(all_values), max(all_values)
        y_range = y_max - y_min if y_max != y_min else y_max
        # Pick ~5 evenly spaced ticks spanning the data range (with padding)
        import math
        pad = y_range * 0.15 if y_range else y_max * 0.15
        tick_lo = max(0, y_min - pad)
        tick_hi = y_max + pad
        raw_step = (tick_hi - tick_lo) / 5
        # Round step to a clean number
        magnitude = 10 ** math.floor(math.log10(raw_step)) if raw_step > 0 else 1
        step = math.ceil(raw_step / magnitude) * magnitude
        tick_start = math.floor(tick_lo / step) * step
        y_tickvals = []
        v = tick_start
        while v <= tick_hi + step:
            if v >= 0:
                y_tickvals.append(v)
            v += step
    else:
        y_tickvals = []

    y_ticktext = [_fmt_abbreviated(v) for v in y_tickvals]

    has_any_forecast = any(f is not None for f in forecasts)

    # Determine text positions to avoid overlap: when actual and forecast
    # are both present for the same month, place the higher value's label
    # on top and the lower one on the bottom.
    actual_positions: list[str] = []
    for a, f in zip(actuals, forecasts):
        if f is not None and a <= f:
            actual_positions.append("bottom center")
        else:
            actual_positions.append("top center")

    # Actuals line
    fig.add_trace(
        go.Scatter(
            x=months,
            y=actuals,
            name="Actual",
            mode="lines+markers+text",
            line=dict(color=ACTUAL_COLOR, width=3),
            marker=dict(symbol="circle", size=12, color=ACTUAL_COLOR),
            text=[_fmt_abbreviated(v) for v in actuals],
            textposition=actual_positions,
        )
    )

    # Forecast line (only include months where forecast exists)
    if has_any_forecast:
        forecast_x = [m for m, f in zip(months, forecasts) if f is not None]
        forecast_y = [f for f in forecasts if f is not None]
        # Mirror: place forecast label opposite of actual
        forecast_positions: list[str] = []
        for a, f in zip(actuals, forecasts):
            if f is None:
                continue
            if f > a:
                forecast_positions.append("top center")
            else:
                forecast_positions.append("bottom center")

        fig.add_trace(
            go.Scatter(
                x=forecast_x,
                y=forecast_y,
                name="Forecast",
                mode="lines+markers+text",
                line=dict(color=FORECAST_COLOR, width=3, dash="dash"),
                marker=dict(symbol="circle", size=12, color=FORECAST_COLOR),
                text=[_fmt_abbreviated(v) for v in forecast_y],
                textposition=forecast_positions,
            )
        )

    # Styling
    text_color = "white" if dark_mode else "#333333"
    bg_color = "rgba(0,0,0,0)" if dark_mode else "white"
    grid_color = "rgba(255,255,255,0.2)" if dark_mode else "rgba(0,0,0,0.1)"

    fig.update_layout(
        title=None,
        xaxis=dict(
            title="",
            tickfont=dict(size=18, color=text_color),
            showgrid=False,
        ),
        yaxis=dict(
            title=dict(text="Spend", font=dict(size=16, color=text_color)),
            tickmode="array",
            tickvals=y_tickvals,
            ticktext=y_ticktext,
            tickfont=dict(size=14, color=text_color),
            gridcolor=grid_color,
        ),
        legend=dict(
            font=dict(size=16, color=text_color),
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
        ),
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        margin=dict(t=50, b=40, l=100, r=40),
        width=1200,
        height=560,
    )

    # Data label font
    fig.update_traces(textfont=dict(size=16, color=text_color))

    return fig


def build_cogs_stacked_bar(
    app_history: list[dict[str, Any]],
    *,
    top_n: int = 10,
    dark_mode: bool = True,
) -> go.Figure:
    """Build a stacked bar chart of COGS by application over time.

    Args:
        app_history: Raw output from ``fetch_app_cogs_history()``. Each dict
            has keys ``app_name`` (str), ``month_label`` (str like
            "Aug 2025"), and ``monthly_cost`` (float).
        top_n: Number of top applications to show individually. All
            remaining apps are collapsed into an "Other" bucket.
        dark_mode: If True, uses white text on transparent background
            (for dark PPTX slides). If False, uses dark text on white
            background (for light-themed output).

    Returns:
        A Plotly Figure object with a stacked bar chart.
    """
    from collections import defaultdict

    # ------------------------------------------------------------------
    # 1. Compute total cost per app and determine top N
    # ------------------------------------------------------------------
    app_totals: dict[str, float] = defaultdict(float)
    for row in app_history:
        app_totals[row["app_name"]] += row["monthly_cost"]

    sorted_apps = sorted(app_totals, key=lambda a: app_totals[a], reverse=True)
    top_apps = sorted_apps[:top_n]
    top_apps_set = set(top_apps)

    # ------------------------------------------------------------------
    # 2. Pivot data: month -> app -> cost (with "Other" bucket)
    # ------------------------------------------------------------------
    # Preserve chronological month order from the input data
    month_order: list[str] = []
    seen_months: set[str] = set()
    for row in app_history:
        if row["month_label"] not in seen_months:
            month_order.append(row["month_label"])
            seen_months.add(row["month_label"])

    month_app_cost: dict[str, dict[str, float]] = {
        m: defaultdict(float) for m in month_order
    }
    for row in app_history:
        month = row["month_label"]
        app = row["app_name"] if row["app_name"] in top_apps_set else "Other"
        month_app_cost[month][app] += row["monthly_cost"]

    # ------------------------------------------------------------------
    # 3. Compute monthly totals (for percentage threshold & total labels)
    # ------------------------------------------------------------------
    month_totals: dict[str, float] = {
        m: sum(costs.values()) for m, costs in month_app_cost.items()
    }

    # ------------------------------------------------------------------
    # 4. Build traces — "Other" first, then apps ascending so the
    #    largest app ends up at the top of each stack.
    # ------------------------------------------------------------------
    # Trace order: Other (bottom), then top apps in ascending total cost
    trace_order: list[str] = []
    has_other = any("Other" in costs for costs in month_app_cost.values())
    if has_other:
        trace_order.append("Other")
    trace_order.extend(reversed(top_apps))  # ascending -> largest last (top)

    text_color = "white" if dark_mode else "#333333"

    fig = go.Figure()

    for idx, app in enumerate(trace_order):
        y_vals = [month_app_cost[m].get(app, 0.0) for m in month_order]

        # Data labels: only show when segment >= 3% of that month's total
        segment_texts: list[str] = []
        for m, val in zip(month_order, y_vals):
            total = month_totals[m]
            if total > 0 and val / total >= 0.03:
                segment_texts.append(_fmt_abbreviated(val))
            else:
                segment_texts.append("")

        # Pick color from palette
        if app == "Other":
            bar_color = STACKED_BAR_PALETTE[-1]
        else:
            # Map app position in top_apps (0-based) to palette index
            palette_idx = top_apps.index(app) % (len(STACKED_BAR_PALETTE) - 1)
            bar_color = STACKED_BAR_PALETTE[palette_idx]

        fig.add_trace(
            go.Bar(
                x=month_order,
                y=y_vals,
                name=app,
                marker_color=bar_color,
                text=segment_texts,
                textposition="inside",
                insidetextanchor="middle",
                textfont=dict(size=14),
            )
        )

    # ------------------------------------------------------------------
    # 5. Total labels above each bar (scatter trace)
    # ------------------------------------------------------------------
    totals_y = [month_totals[m] for m in month_order]
    totals_text = [_fmt_abbreviated(v) for v in totals_y]

    fig.add_trace(
        go.Scatter(
            x=month_order,
            y=totals_y,
            mode="text",
            text=totals_text,
            textposition="top center",
            textfont=dict(size=16, color=text_color),
            showlegend=False,
            hoverinfo="skip",
        )
    )

    # ------------------------------------------------------------------
    # 6. Compute y-axis ticks (abbreviated currency labels)
    # ------------------------------------------------------------------
    import math

    all_totals = [month_totals[m] for m in month_order]
    y_max = max(all_totals) if all_totals else 0
    pad = y_max * 0.15 if y_max else 1
    tick_hi = y_max + pad
    raw_step = tick_hi / 5
    magnitude = 10 ** math.floor(math.log10(raw_step)) if raw_step > 0 else 1
    step = math.ceil(raw_step / magnitude) * magnitude
    y_tickvals = []
    v = 0
    while v <= tick_hi:
        y_tickvals.append(v)
        v += step
    y_ticktext = [_fmt_abbreviated(v) for v in y_tickvals]

    # ------------------------------------------------------------------
    # 7. Layout styling
    # ------------------------------------------------------------------
    bg_color = "rgba(0,0,0,0)" if dark_mode else "white"
    grid_color = "rgba(255,255,255,0.12)" if dark_mode else "rgba(0,0,0,0.08)"

    fig.update_layout(
        barmode="stack",
        bargap=0.25,
        title=None,
        xaxis=dict(
            title="",
            tickfont=dict(size=18, color=text_color),
            showgrid=False,
        ),
        yaxis=dict(
            title=dict(text="Spend", font=dict(size=16, color=text_color)),
            tickmode="array",
            tickvals=y_tickvals,
            ticktext=y_ticktext,
            tickfont=dict(size=14, color=text_color),
            showgrid=True,
            gridcolor=grid_color,
            zeroline=False,
        ),
        legend=dict(
            font=dict(size=14, color=text_color),
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
        ),
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        margin=dict(t=30, b=40, l=100, r=40),
        width=1200,
        height=560,
    )

    return fig


def build_unit_cost_dual_axis(
    org_history: list[dict[str, Any]],
    *,
    dark_mode: bool = True,
) -> go.Figure:
    """Build dual-axis chart: COGS bars (left) + Analyzed Observations line (right).

    Shows whether COGS and observation volume move together over time.

    Args:
        org_history: Chronological list of org-level monthly dicts from
            ``fetch_unit_cost_data()``.
        dark_mode: Dark text/background styling for PPTX slides.

    Returns:
        A Plotly Figure with dual y-axes.
    """
    import math
    from plotly.subplots import make_subplots

    months = [h["month_label"] for h in org_history]
    cogs = [h["total_cogs"] for h in org_history]
    obs_trillions = [h["total_analyzed_obs"] / 1e12 for h in org_history]

    text_color = "white" if dark_mode else "#333333"
    bg_color = "rgba(0,0,0,0)" if dark_mode else "white"
    grid_color = "rgba(255,255,255,0.12)" if dark_mode else "rgba(0,0,0,0.08)"

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # COGS bars on primary y-axis
    fig.add_trace(
        go.Bar(
            x=months,
            y=cogs,
            name="COGS ($)",
            marker_color=ACTUAL_COLOR,
            text=[_fmt_abbreviated(v) for v in cogs],
            textposition="outside",
            textfont=dict(size=14, color=text_color),
        ),
        secondary_y=False,
    )

    # Analyzed observations line on secondary y-axis
    fig.add_trace(
        go.Scatter(
            x=months,
            y=obs_trillions,
            name="Analyzed Obs (T)",
            mode="lines+markers+text",
            line=dict(color=FORECAST_COLOR, width=3),
            marker=dict(symbol="circle", size=10, color=FORECAST_COLOR),
            text=[f"{v:.1f}T" for v in obs_trillions],
            textposition="top center",
            textfont=dict(size=14, color=text_color),
        ),
        secondary_y=True,
    )

    # COGS y-axis ticks
    y_max = max(cogs) if cogs else 0
    pad = y_max * 0.25
    tick_hi = y_max + pad
    raw_step = tick_hi / 5
    magnitude = 10 ** math.floor(math.log10(raw_step)) if raw_step > 0 else 1
    step = math.ceil(raw_step / magnitude) * magnitude
    y_tickvals: list[float] = []
    v = 0.0
    while v <= tick_hi:
        y_tickvals.append(v)
        v += step
    y_ticktext = [_fmt_abbreviated(v) for v in y_tickvals]

    fig.update_layout(
        title=None,
        barmode="group",
        bargap=0.3,
        xaxis=dict(
            title="",
            tickfont=dict(size=18, color=text_color),
            showgrid=False,
        ),
        legend=dict(
            font=dict(size=14, color=text_color),
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
        ),
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        margin=dict(t=50, b=40, l=100, r=100),
        width=1200,
        height=560,
    )

    fig.update_yaxes(
        title=dict(text="COGS ($)", font=dict(size=16, color=text_color)),
        tickmode="array",
        tickvals=y_tickvals,
        ticktext=y_ticktext,
        tickfont=dict(size=14, color=text_color),
        showgrid=True,
        gridcolor=grid_color,
        secondary_y=False,
    )

    obs_max = max(obs_trillions) if obs_trillions else 0
    fig.update_yaxes(
        title=dict(text="Analyzed Observations (T)", font=dict(size=16, color=text_color)),
        tickfont=dict(size=14, color=text_color),
        showgrid=False,
        secondary_y=True,
        range=[0, obs_max * 1.3],
    )

    return fig


def build_cost_per_1m_trend(
    org_history: list[dict[str, Any]],
    *,
    dark_mode: bool = True,
) -> go.Figure:
    """Build a line chart of org-level cost per 1M analyzed observations.

    This is the key efficiency metric. A flat or declining line means
    COGS increases are volume-driven rather than unit-cost-driven.

    Args:
        org_history: Chronological list from ``fetch_unit_cost_data()``.
        dark_mode: Dark text/background styling.

    Returns:
        A Plotly Figure.
    """
    import math

    months = [h["month_label"] for h in org_history]
    unit_costs = [h["cogs_per_1m_analyzed"] for h in org_history]

    text_color = "white" if dark_mode else "#333333"
    bg_color = "rgba(0,0,0,0)" if dark_mode else "white"
    grid_color = "rgba(255,255,255,0.2)" if dark_mode else "rgba(0,0,0,0.1)"

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=months,
            y=unit_costs,
            name="Cost / 1M Analyzed Obs",
            mode="lines+markers+text",
            line=dict(color=ACTUAL_COLOR, width=3),
            marker=dict(symbol="circle", size=12, color=ACTUAL_COLOR),
            text=[f"${v:.2f}" for v in unit_costs],
            textposition="top center",
            textfont=dict(size=16, color=text_color),
        )
    )

    # Y-axis ticks
    if unit_costs:
        y_min, y_max = min(unit_costs), max(unit_costs)
        y_range = y_max - y_min if y_max != y_min else y_max
        pad = y_range * 0.2 if y_range else y_max * 0.2
        tick_lo = max(0, y_min - pad)
        tick_hi = y_max + pad
        raw_step = (tick_hi - tick_lo) / 5
        magnitude = 10 ** math.floor(math.log10(raw_step)) if raw_step > 0 else 0.01
        step = math.ceil(raw_step / magnitude) * magnitude
        tick_start = math.floor(tick_lo / step) * step if step > 0 else 0
        y_tickvals: list[float] = []
        v = tick_start
        while v <= tick_hi + step:
            if v >= 0:
                y_tickvals.append(round(v, 4))
            v += step
    else:
        y_tickvals = []
    y_ticktext = [f"${v:.2f}" for v in y_tickvals]

    fig.update_layout(
        title=None,
        xaxis=dict(
            title="",
            tickfont=dict(size=18, color=text_color),
            showgrid=False,
        ),
        yaxis=dict(
            title=dict(text="Cost per 1M Analyzed Obs ($)", font=dict(size=16, color=text_color)),
            tickmode="array",
            tickvals=y_tickvals,
            ticktext=y_ticktext,
            tickfont=dict(size=14, color=text_color),
            gridcolor=grid_color,
        ),
        legend=dict(
            font=dict(size=16, color=text_color),
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
        ),
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        margin=dict(t=50, b=40, l=100, r=40),
        width=1200,
        height=560,
    )

    return fig


def build_cost_per_1m_combo(
    org_history: list[dict[str, Any]],
    *,
    dark_mode: bool = True,
) -> go.Figure:
    """Combo chart: Cost/1M line (left) + COGS bars + Observed Obs line (right).

    Lets the reader see the unit-cost trend alongside absolute COGS and volume
    in one view, covering all available months (Nov 2025 onwards).

    Args:
        org_history: Chronological list from ``fetch_unit_cost_data()``.
        dark_mode: Dark text/background styling.

    Returns:
        A Plotly Figure with dual y-axes.
    """
    import math
    from plotly.subplots import make_subplots

    months = [h["month_label"] for h in org_history]
    unit_costs = [h["cogs_per_1m_analyzed"] for h in org_history]
    cogs = [h["total_cogs"] for h in org_history]
    obs_billions = [h["total_analyzed_obs"] / 1e9 for h in org_history]

    text_color = "white" if dark_mode else "#333333"
    bg_color = "rgba(0,0,0,0)" if dark_mode else "white"
    grid_color = "rgba(255,255,255,0.12)" if dark_mode else "rgba(0,0,0,0.08)"

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # COGS bars on secondary y-axis (background context)
    fig.add_trace(
        go.Bar(
            x=months,
            y=cogs,
            name="Total COGS ($)",
            marker_color="rgba(0,153,255,0.35)",
            text=[_fmt_abbreviated(v) for v in cogs],
            textposition="outside",
            textfont=dict(size=13, color=text_color),
        ),
        secondary_y=True,
    )

    # Analyzed Observations line on secondary y-axis
    fig.add_trace(
        go.Scatter(
            x=months,
            y=obs_billions,
            name="Analyzed Obs (B)",
            mode="lines+markers",
            line=dict(color="#83D6FF", width=2, dash="dot"),
            marker=dict(symbol="circle", size=8, color="#83D6FF"),
        ),
        secondary_y=True,
    )

    # Cost/1M line on primary y-axis (the headline metric)
    fig.add_trace(
        go.Scatter(
            x=months,
            y=unit_costs,
            name="Cost / 1M Analyzed Obs ($)",
            mode="lines+markers+text",
            line=dict(color=ACTUAL_COLOR, width=3),
            marker=dict(symbol="circle", size=12, color=ACTUAL_COLOR),
            text=[f"${v:.2f}" for v in unit_costs],
            textposition="top center",
            textfont=dict(size=16, color=text_color),
        ),
        secondary_y=False,
    )

    # Primary y-axis ticks (Cost/1M)
    if unit_costs:
        y_min, y_max = min(unit_costs), max(unit_costs)
        y_range = y_max - y_min if y_max != y_min else y_max
        pad = y_range * 0.25 if y_range else y_max * 0.25
        tick_lo = max(0, y_min - pad)
        tick_hi = y_max + pad
        raw_step = (tick_hi - tick_lo) / 5
        magnitude = 10 ** math.floor(math.log10(raw_step)) if raw_step > 0 else 0.01
        step = math.ceil(raw_step / magnitude) * magnitude
        tick_start = math.floor(tick_lo / step) * step if step > 0 else 0
        y_tickvals: list[float] = []
        v = tick_start
        while v <= tick_hi + step:
            if v >= 0:
                y_tickvals.append(round(v, 4))
            v += step
    else:
        y_tickvals = []
    y_ticktext = [f"${v:.2f}" for v in y_tickvals]

    fig.update_layout(
        title=None,
        barmode="group",
        bargap=0.25,
        xaxis=dict(
            title="",
            tickfont=dict(size=18, color=text_color),
            showgrid=False,
        ),
        legend=dict(
            font=dict(size=14, color=text_color),
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
        ),
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        margin=dict(t=50, b=40, l=110, r=110),
        width=1200,
        height=560,
    )

    fig.update_yaxes(
        title=dict(text="Cost per 1M Analyzed Obs ($)", font=dict(size=15, color=text_color)),
        tickmode="array",
        tickvals=y_tickvals,
        ticktext=y_ticktext,
        tickfont=dict(size=13, color=text_color),
        gridcolor=grid_color,
        secondary_y=False,
    )

    cogs_max = max(cogs) if cogs else 0
    fig.update_yaxes(
        title=dict(text="COGS ($) / Analyzed Obs (B)", font=dict(size=15, color=text_color)),
        tickformat="$,.0f",
        tickfont=dict(size=13, color=text_color),
        showgrid=False,
        secondary_y=True,
        range=[0, cogs_max * 1.5],
    )

    return fig


def build_pod_unit_cost_trend(
    pod_history: list[dict[str, Any]],
    *,
    dark_mode: bool = True,
) -> go.Figure:
    """Build a multi-line chart of cost per 1M analyzed by pod over time.

    Args:
        pod_history: Chronological list of per-pod dicts from
            ``fetch_unit_cost_data()``.
        dark_mode: Dark text/background styling.

    Returns:
        A Plotly Figure with one line per pod.
    """
    import math
    from collections import defaultdict

    # Pivot: pod -> ordered list of (month_label, cogs_per_1m)
    pod_data: dict[str, list[tuple[str, float]]] = defaultdict(list)
    month_order: list[str] = []
    seen: set[str] = set()
    for row in pod_history:
        pod = row["pod"].upper()
        ml = row["month_label"]
        pod_data[pod].append((ml, row["cogs_per_1m_analyzed"]))
        if ml not in seen:
            month_order.append(ml)
            seen.add(ml)

    text_color = "white" if dark_mode else "#333333"
    bg_color = "rgba(0,0,0,0)" if dark_mode else "white"
    grid_color = "rgba(255,255,255,0.2)" if dark_mode else "rgba(0,0,0,0.1)"

    fig = go.Figure()

    # Sort pods for consistent ordering; use the palette
    sorted_pods = sorted(pod_data.keys())
    for idx, pod in enumerate(sorted_pods):
        points = pod_data[pod]
        x_vals = [p[0] for p in points]
        y_vals = [p[1] for p in points]
        color = STACKED_BAR_PALETTE[idx % len(STACKED_BAR_PALETTE)]

        fig.add_trace(
            go.Scatter(
                x=x_vals,
                y=y_vals,
                name=pod,
                mode="lines+markers",
                line=dict(color=color, width=2.5),
                marker=dict(symbol="circle", size=8, color=color),
            )
        )

    # Y-axis ticks
    all_vals = [row["cogs_per_1m_analyzed"] for row in pod_history]
    if all_vals:
        y_min, y_max = min(all_vals), max(all_vals)
        y_range = y_max - y_min if y_max != y_min else y_max
        pad = y_range * 0.15
        tick_lo = max(0, y_min - pad)
        tick_hi = y_max + pad
        raw_step = (tick_hi - tick_lo) / 5
        magnitude = 10 ** math.floor(math.log10(raw_step)) if raw_step > 0 else 0.01
        step = math.ceil(raw_step / magnitude) * magnitude
        tick_start = math.floor(tick_lo / step) * step if step > 0 else 0
        y_tickvals: list[float] = []
        v = tick_start
        while v <= tick_hi + step:
            if v >= 0:
                y_tickvals.append(round(v, 4))
            v += step
    else:
        y_tickvals = []
    y_ticktext = [f"${v:.2f}" for v in y_tickvals]

    fig.update_layout(
        title=None,
        xaxis=dict(
            title="",
            tickfont=dict(size=18, color=text_color),
            showgrid=False,
        ),
        yaxis=dict(
            title=dict(text="Cost per 1M Analyzed Obs ($)", font=dict(size=16, color=text_color)),
            tickmode="array",
            tickvals=y_tickvals,
            ticktext=y_ticktext,
            tickfont=dict(size=14, color=text_color),
            gridcolor=grid_color,
        ),
        legend=dict(
            font=dict(size=14, color=text_color),
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
        ),
        paper_bgcolor=bg_color,
        plot_bgcolor=bg_color,
        margin=dict(t=50, b=40, l=100, r=40),
        width=1200,
        height=560,
    )

    return fig


def render_chart_png(fig: go.Figure) -> bytes:
    """Export a Plotly figure to PNG bytes via kaleido.

    Uses ``scale=2`` for crisp rendering on high-DPI slides.

    Args:
        fig: A Plotly Figure object.

    Returns:
        PNG image as raw bytes.
    """
    return fig.to_image(format="png", scale=2, engine="kaleido")


def build_all_trend_charts(
    chart_data: dict[str, list[dict[str, Any]]],
    sections: list[dict[str, Any]],
) -> dict[str, bytes]:
    """Render trend chart PNGs for all sections.

    Args:
        chart_data: Keyed by bucket name. Each value is a chronological list
            of dicts with ``month_label``, ``actual``, and ``forecast`` keys.
        sections: The ``pptx.sections`` list from config.yaml, each with
            ``bucket`` and ``title`` keys.

    Returns:
        A dict mapping bucket name to PNG bytes.
    """
    result: dict[str, bytes] = {}

    for section in sections:
        bucket = section["bucket"]
        history = chart_data.get(bucket)
        if history is None:
            logger.warning("No chart data for bucket '%s', skipping.", bucket)
            continue

        fig = build_trend_chart(bucket, history, dark_mode=True)
        png_bytes = render_chart_png(fig)
        result[bucket] = png_bytes
        logger.info(
            "Rendered trend chart for '%s': %d bytes", bucket, len(png_bytes)
        )

    return result
