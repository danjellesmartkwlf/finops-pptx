"""Template-based narrative generation for the FinOps Report Generator.

Selects the appropriate template (Mode A with forecast, Mode B without) and
fills it with the calculated metrics to produce executive-ready text.

Before rendering, the raw metrics dict is enriched with presentation-ready
fields: formatted dollar strings (M/K suffixes), month names, daily averages,
and directional words ("higher"/"lower").
"""

from __future__ import annotations

import calendar

from src.pptx_utils import fmt_abbreviated as _format_dollars


# ---------------------------------------------------------------------------
# Metrics enrichment
# ---------------------------------------------------------------------------

def _enrich_metrics(metrics: dict, month_label: str) -> dict:
    """Add presentation-ready fields to a metrics dict.

    New keys added (all safe to use in config.yaml templates):

    Month names:
        ``curr_month_name``, ``prev_month_name``

    Formatted dollar amounts:
        ``actual_fmt``, ``previous_fmt``, ``mom_delta_fmt``

    Daily averages:
        ``curr_days``, ``prev_days``,
        ``daily_avg_curr_fmt``, ``daily_avg_prev_fmt``,
        ``daily_delta_fmt``, ``daily_delta_dir``

    Direction words:
        ``mom_delta_dir`` ("higher" / "lower")

    Forecast (when present):
        ``forecast_fmt``, ``var_delta_fmt``
    """
    enriched = dict(metrics)

    # Parse "January 2026" into components
    parts = month_label.split()
    month_name = parts[0] if parts else ""
    year = int(parts[1]) if len(parts) > 1 else 2026
    month_num = list(calendar.month_name).index(month_name) if month_name else 1

    # Previous month
    if month_num == 1:
        prev_month_num, prev_year = 12, year - 1
    else:
        prev_month_num, prev_year = month_num - 1, year
    prev_month_name = calendar.month_name[prev_month_num]

    # Days in each month
    curr_days = calendar.monthrange(year, month_num)[1]
    prev_days = calendar.monthrange(prev_year, prev_month_num)[1]

    # Daily averages
    daily_avg_curr = metrics["actual"] / curr_days if curr_days else 0.0
    daily_avg_prev = metrics["previous_month"] / prev_days if prev_days else 0.0
    daily_delta = daily_avg_curr - daily_avg_prev

    enriched.update({
        # Month names
        "curr_month_name": month_name,
        "prev_month_name": prev_month_name,

        # Formatted totals
        "actual_fmt": _format_dollars(metrics["actual"]),
        "previous_fmt": _format_dollars(metrics["previous_month"]),
        "mom_delta_fmt": _format_dollars(abs(metrics["mom_delta"])),
        "mom_delta_dir": "higher" if metrics["mom_delta"] >= 0 else "lower",

        # Daily averages
        "curr_days": curr_days,
        "prev_days": prev_days,
        "daily_avg_curr_fmt": _format_dollars(daily_avg_curr),
        "daily_avg_prev_fmt": _format_dollars(daily_avg_prev),
        "daily_delta_fmt": _format_dollars(abs(daily_delta)),
        "daily_delta_dir": "higher" if daily_delta >= 0 else "lower",
    })

    # Forecast-specific formatted fields
    if metrics.get("has_forecast") and metrics.get("forecast") is not None:
        enriched["forecast_fmt"] = _format_dollars(metrics["forecast"])
        enriched["var_delta_fmt"] = _format_dollars(abs(metrics["var_delta"]))

    return enriched


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_narrative(
    metrics: dict,
    month: str,
    templates: dict,
) -> str:
    """Generate narrative text for a single bucket.

    Args:
        metrics: A single bucket's metrics dict as returned by
            :func:`src.calculations.calculate_metrics`.
        month: Human-readable reporting month (e.g. "January 2026").
        templates: The ``narrative_templates`` section from config.yaml,
            containing ``mode_a`` and ``mode_b`` sub-dicts each with a
            ``template`` string.

    Returns:
        The fully rendered narrative string ready for presentation.
    """
    enriched = _enrich_metrics(metrics, month)

    if enriched["has_forecast"]:
        template_str: str = templates["mode_a"]["template"]
    else:
        template_str = templates["mode_b"]["template"]

    return template_str.format(month=month, **enriched)


def generate_all_narratives(
    all_metrics: list[dict],
    month: str,
    config: dict,
) -> dict[str, str]:
    """Generate narratives for every bucket in the report.

    Args:
        all_metrics: List of metric dicts (one per bucket), as returned by
            :func:`src.calculations.calculate_all_buckets`.
        month: Human-readable reporting month (e.g. "January 2026").
        config: The full parsed config.yaml dict. The function reads the
            ``narrative_templates`` key from this dict.

    Returns:
        A dict mapping each bucket name to its rendered narrative string.
    """
    templates: dict = config["narrative_templates"]

    return {
        bucket_metrics["metric"]: generate_narrative(
            metrics=bucket_metrics,
            month=month,
            templates=templates,
        )
        for bucket_metrics in all_metrics
    }
