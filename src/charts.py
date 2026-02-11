"""Trend chart generation for the FinOps Report Generator.

Builds 6-month line charts using Plotly and exports them as PNG images
via kaleido for embedding into PowerPoint slides.
"""

from __future__ import annotations

import logging
from typing import Any

import plotly.graph_objects as go

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Color constants
# ---------------------------------------------------------------------------

ACTUAL_COLOR = "#636EFA"   # solid blue
FORECAST_COLOR = "#00CC96"  # dashed green

# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------


def _fmt_abbreviated(value: float) -> str:
    """Format a dollar value with abbreviated suffix (K, M, B).

    Examples:
        4894000  -> "$4.89M"
        123400   -> "$123.4K"
        1200000000 -> "$1.20B"
        850      -> "$850"
    """
    abs_val = abs(value)
    sign = "-" if value < 0 else ""
    if abs_val >= 1_000_000_000:
        return f"{sign}${abs_val / 1_000_000_000:.2f}B"
    if abs_val >= 1_000_000:
        return f"{sign}${abs_val / 1_000_000:.2f}M"
    if abs_val >= 1_000:
        return f"{sign}${abs_val / 1_000:.1f}K"
    return f"{sign}${abs_val:,.0f}"


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
            background (for Streamlit preview).

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
