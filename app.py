"""FinOps Report Generator -- Streamlit UI ('The Cockpit').

Provides a browser-based interface for reviewing monthly cloud cost reports,
editing generated narratives, visualising actuals vs. forecasts, and exporting
the final deliverable as a PowerPoint deck.
"""

from __future__ import annotations

import atexit
import calendar
import io
import time
from typing import Any

import plotly.graph_objects as go
import streamlit as st

from src.calculations import calculate_all_buckets
from src.forecast import load_forecast
from src.ingestion import fetch_bucket_actuals, load_config, close_shared_connection
from src.narrative import generate_all_narratives

atexit.register(close_shared_connection)

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="FinOps Report Generator",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

MONTH_NAMES: list[str] = list(calendar.month_name)[1:]  # January .. December
BUCKET_ORDER: list[str] = ["Total", "COGS", "OpEx"]

# ---------------------------------------------------------------------------
# Sidebar controls
# ---------------------------------------------------------------------------

st.sidebar.title("Report Parameters")

selected_month_name: str = st.sidebar.selectbox(
    "Month",
    options=MONTH_NAMES,
    index=0,
)

selected_year: int = int(
    st.sidebar.number_input(
        "Year",
        min_value=2020,
        max_value=2099,
        value=2026,
        step=1,
    )
)

generate_clicked: bool = st.sidebar.button("Generate Report", type="primary")

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.title("FinOps Report Generator")

# ---------------------------------------------------------------------------
# Redshift connectivity check (runs once per session)
# ---------------------------------------------------------------------------

if "redshift_ok" not in st.session_state:
    with st.spinner("Checking Redshift connectivity..."):
        from src.ingestion import check_redshift_connection
        ok, msg = check_redshift_connection()
        st.session_state["redshift_ok"] = ok
        st.session_state["redshift_msg"] = msg

if not st.session_state.get("redshift_ok"):
    st.error(
        f"Cannot connect to Redshift. {st.session_state.get('redshift_msg', '')} "
        "Please check your .env configuration and network access."
    )
    st.stop()

# ---------------------------------------------------------------------------
# Report generation logic
# ---------------------------------------------------------------------------


def _month_name_to_padded_number(name: str) -> str:
    """Convert a full month name (e.g. 'January') to a zero-padded string ('01')."""
    month_num: int = list(calendar.month_name).index(name)
    return f"{month_num:02d}"


def _build_bucket_chart(
    bucket_name: str,
    metrics: dict[str, Any],
) -> go.Figure:
    """Build a Plotly bar chart comparing current, previous, and (optionally) forecast."""
    categories: list[str] = ["Previous Month", "Current Month"]
    values: list[float] = [metrics["previous_month"], metrics["actual"]]
    colors: list[str] = ["#636EFA", "#EF553B"]

    if metrics.get("has_forecast") and metrics.get("forecast") is not None:
        categories.append("Forecast")
        values.append(metrics["forecast"])
        colors.append("#00CC96")

    fig = go.Figure(
        data=[
            go.Bar(
                x=categories,
                y=values,
                marker_color=colors,
                text=[f"${v:,.0f}" for v in values],
                textposition="outside",
            )
        ]
    )

    fig.update_layout(
        title=dict(text=bucket_name, x=0.5),
        yaxis_title="Spend ($)",
        xaxis_title="",
        template="plotly_white",
        height=400,
        margin=dict(t=60, b=40),
    )

    return fig


@st.cache_data(ttl=3600, show_spinner=False)
def _fetch_actuals_cached(month: str, year: int) -> dict[str, dict[str, float]]:
    """Fetch Redshift actuals with a 1-hour cache."""
    return fetch_bucket_actuals(month, year)


@st.cache_data(ttl=3600, show_spinner=False)
def _load_forecast_cached(month_name: str, year: int) -> dict[str, float | None]:
    """Load forecast data with a 1-hour cache."""
    config = load_config("config.yaml")
    return load_forecast(config, month_name, year)


if generate_clicked:
    month_number_str: str = _month_name_to_padded_number(selected_month_name)

    with st.status("Generating report...", expanded=True) as status:
        # -- Load configuration ------------------------------------------------
        st.write("Loading configuration...")
        config: dict[str, Any] = load_config("config.yaml")
        bucket_names = [b["name"] for b in config.get("buckets", [])]
        st.write(f"Found {len(bucket_names)} buckets: {', '.join(bucket_names)}")

        # -- Fetch actuals from Redshift ---------------------------------------
        st.write("Fetching actuals from Redshift...")
        t0 = time.perf_counter()
        actuals: dict[str, dict[str, float]] = _fetch_actuals_cached(
            month_number_str, selected_year
        )
        elapsed = time.perf_counter() - t0
        cached_note = " *(cached)*" if elapsed < 0.5 else ""
        for bname, vals in actuals.items():
            st.write(
                f"&ensp;&ensp;{bname}: "
                f"**${vals['current_month']:,.2f}** current / "
                f"**${vals['previous_month']:,.2f}** previous"
            )
        st.write(f"Actuals loaded{cached_note}")

        # -- Load forecast (graceful degradation) ------------------------------
        forecasts: dict[str, float | None]
        st.write("Loading forecast data...")
        t0 = time.perf_counter()
        try:
            forecasts = _load_forecast_cached(selected_month_name, selected_year)
            elapsed = time.perf_counter() - t0
            cached_note = " *(cached)*" if elapsed < 0.5 else ""
            for key, val in forecasts.items():
                display = f"${val:,.2f}" if val is not None else "N/A"
                st.write(f"&ensp;&ensp;{key}: **{display}**")
            st.write(f"Forecasts loaded{cached_note}")
        except FileNotFoundError:
            st.write("Forecast file not found -- continuing without forecast data.")
            forecasts = {
                bucket["forecast_mapping_key"]: None
                for bucket in config["buckets"]
            }

        # -- Calculate metrics -------------------------------------------------
        st.write("Calculating metrics...")
        all_metrics: list[dict[str, Any]] = calculate_all_buckets(
            actuals, forecasts, config["buckets"]
        )
        for m in all_metrics:
            pct = abs(m["mom_pct"])
            st.write(
                f"&ensp;&ensp;{m['metric']}: {m['mom_dir']} of {pct:.1f}% MoM"
            )

        # -- Generate narratives -----------------------------------------------
        month_label: str = f"{selected_month_name} {selected_year}"
        st.write("Generating narratives...")
        narratives: dict[str, str] = generate_all_narratives(
            all_metrics, month_label, config
        )
        st.write(f"Generated {len(narratives)} narrative(s)")

        status.update(
            label="Report generated!", state="complete", expanded=False
        )

    # -- Persist results in session state -----------------------------------
    st.session_state["all_metrics"] = all_metrics
    st.session_state["narratives"] = narratives
    st.session_state["config"] = config
    st.session_state["report_generated"] = True
    st.session_state["month_label"] = month_label

    # Initialise editable narrative keys (only on fresh generation)
    for bucket_name, text in narratives.items():
        st.session_state[f"narrative_{bucket_name}"] = text

# ---------------------------------------------------------------------------
# Main content area -- displayed after generation
# ---------------------------------------------------------------------------

if st.session_state.get("report_generated"):
    all_metrics: list[dict[str, Any]] = st.session_state["all_metrics"]
    config: dict[str, Any] = st.session_state["config"]

    # Build a lookup from bucket name to its metrics dict
    metrics_by_name: dict[str, dict[str, Any]] = {
        m["metric"]: m for m in all_metrics
    }

    st.divider()

    for bucket_name in BUCKET_ORDER:
        metrics = metrics_by_name.get(bucket_name)
        if metrics is None:
            continue

        st.subheader(bucket_name)

        left_col, right_col = st.columns(2)

        with left_col:
            st.text_area(
                label=f"{bucket_name} Narrative",
                value=st.session_state.get(
                    f"narrative_{bucket_name}", ""
                ),
                height=200,
                key=f"narrative_{bucket_name}",
            )

        with right_col:
            fig = _build_bucket_chart(bucket_name, metrics)
            st.plotly_chart(fig, width="stretch")

    # -------------------------------------------------------------------
    # MCP Investigate widget
    # -------------------------------------------------------------------
    st.divider()

    with st.expander("Investigate with AI"):
        st.text_input(
            "Ask a question about the data",
            placeholder="e.g. Which service drove the biggest MoM increase?",
            key="mcp_query",
        )
        st.info(
            "MCP integration coming soon. Use the redshift MCP server "
            "in Claude Code directly for ad-hoc queries."
        )

    # -------------------------------------------------------------------
    # Download PowerPoint
    # -------------------------------------------------------------------
    st.divider()

    def _export_pptx() -> bytes | None:
        """Generate the PowerPoint file using edited narratives.

        Returns:
            The .pptx file contents as bytes, or None on failure.
        """
        # Collect the *edited* narratives from session state
        edited_narratives: dict[str, str] = {}
        for bucket_name in BUCKET_ORDER:
            key = f"narrative_{bucket_name}"
            edited_narratives[bucket_name] = st.session_state.get(key, "")

        try:
            from src.pptx_gen import generate_pptx  # type: ignore[import-untyped]

            pptx_bytes: bytes = generate_pptx(
                edited_narratives,
                st.session_state["all_metrics"],
                st.session_state["config"],
                month_label=st.session_state.get("month_label", ""),
            )
            return pptx_bytes
        except Exception as exc:
            st.error(
                f"Failed to generate PowerPoint: {exc}. "
                "Ensure the template file exists and src/pptx_gen.py is implemented."
            )
            return None

    # Two-step download: generate first, then offer download button
    if st.button("Prepare PowerPoint Download"):
        pptx_data = _export_pptx()
        if pptx_data is not None:
            st.session_state["pptx_bytes"] = pptx_data

    if st.session_state.get("pptx_bytes"):
        month_label = st.session_state.get("month_label", "report")
        safe_filename = month_label.replace(" ", "_").lower()
        st.download_button(
            label="Download PowerPoint",
            data=st.session_state["pptx_bytes"],
            file_name=f"finops_report_{safe_filename}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
