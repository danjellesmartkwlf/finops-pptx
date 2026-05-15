"""App-level COGS data processing for Phase 2 slides.

Loads the app-to-category mapping, merges with actuals and forecasts,
computes category rollups, and identifies top movers.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Mapping loader
# ---------------------------------------------------------------------------

def load_app_category_mapping(
    config: dict[str, Any],
) -> dict[str, str]:
    """Load the awn_app -> awn_category mapping from the Excel file.

    Args:
        config: Parsed config.yaml.

    Returns:
        A dict mapping each awn_app name to its awn_category.
    """
    mapping_cfg = (
        config.get("data_files", {})
        .get("mapping", {})
        .get("app_category", {})
    )
    file_path = Path(mapping_cfg.get("file_path", "data/mapping/app_category_mapping.xlsx"))
    sheet_name = mapping_cfg.get("sheet_name", "Sheet1")

    if not file_path.exists():
        logger.warning("App category mapping file not found: %s", file_path)
        return {}

    df = pd.read_excel(file_path, sheet_name=sheet_name)
    result: dict[str, str] = {}
    for _, row in df.iterrows():
        app = row.get("awn_app")
        cat = row.get("awn_category")
        if pd.notna(app) and pd.notna(cat):
            result[str(app).strip()] = str(cat).strip()

    logger.info("Loaded %d app-to-category mappings", len(result))
    return result


# ---------------------------------------------------------------------------
# App metrics builder
# ---------------------------------------------------------------------------

def build_app_metrics(
    app_actuals: dict[str, dict[str, float]],
    app_forecasts: dict[str, float],
    app_category_map: dict[str, str],
) -> list[dict[str, Any]]:
    """Build per-app metrics combining actuals, forecasts, and category.

    Args:
        app_actuals: Keyed by awn_app. Each value has ``current_month``
            and ``previous_month``.
        app_forecasts: Keyed by awn_app -> forecast value.
        app_category_map: Keyed by awn_app -> awn_category.

    Returns:
        A list of dicts sorted by current_month spend descending, each with:
            - ``app``: awn_app name
            - ``category``: awn_category (or "Other")
            - ``current_month``: current month spend
            - ``previous_month``: previous month spend
            - ``forecast``: forecast value or None
            - ``mom_change``: current - previous
            - ``mom_pct``: MoM percentage change
            - ``var_change``: actual - forecast (or None)
            - ``var_pct``: variance percentage (or None)
    """
    all_apps = set(app_actuals.keys()) | set(app_forecasts.keys())
    metrics: list[dict[str, Any]] = []

    for app in all_apps:
        actuals = app_actuals.get(app, {"current_month": 0.0, "previous_month": 0.0, "two_months_ago": 0.0})
        current = actuals["current_month"]
        previous = actuals["previous_month"]
        two_months_ago = actuals.get("two_months_ago", 0.0)
        forecast = app_forecasts.get(app)
        category = app_category_map.get(app, "Other")

        mom_change = current - previous
        mom_pct = (mom_change / previous * 100.0) if previous != 0 else 0.0

        # Previous month vs two months ago
        prev_vs_prior_change = previous - two_months_ago
        prev_vs_prior_pct = (prev_vs_prior_change / two_months_ago * 100.0) if two_months_ago != 0 else 0.0

        var_change: float | None = None
        var_pct: float | None = None
        if forecast is not None:
            var_change = current - forecast
            var_pct = (var_change / forecast * 100.0) if forecast != 0 else 0.0

        metrics.append({
            "app": app,
            "category": category,
            "current_month": current,
            "previous_month": previous,
            "two_months_ago": two_months_ago,
            "forecast": forecast,
            "mom_change": mom_change,
            "mom_pct": mom_pct,
            "prev_vs_prior_change": prev_vs_prior_change,
            "prev_vs_prior_pct": prev_vs_prior_pct,
            "var_change": var_change,
            "var_pct": var_pct,
        })

    metrics.sort(key=lambda m: m["current_month"], reverse=True)
    return metrics


# ---------------------------------------------------------------------------
# Category rollup
# ---------------------------------------------------------------------------

def build_category_rollup(
    app_metrics: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Aggregate app metrics into category-level totals.

    Args:
        app_metrics: Per-app metrics from :func:`build_app_metrics`.

    Returns:
        A list of dicts sorted by current_month spend descending, each with:
            - ``category``: awn_category name
            - ``current_month``: summed current month spend
            - ``previous_month``: summed previous month spend
            - ``forecast``: summed forecast (None if no apps have forecast)
            - ``mom_change``: current - previous
            - ``mom_pct``: MoM percentage change
            - ``var_change``: actual - forecast (or None)
            - ``var_pct``: variance percentage (or None)
    """
    categories: dict[str, dict[str, Any]] = {}

    for m in app_metrics:
        cat = m["category"]
        if cat not in categories:
            categories[cat] = {
                "category": cat,
                "current_month": 0.0,
                "previous_month": 0.0,
                "forecast_sum": 0.0,
                "has_forecast": False,
            }
        categories[cat]["current_month"] += m["current_month"]
        categories[cat]["previous_month"] += m["previous_month"]
        if m["forecast"] is not None:
            categories[cat]["forecast_sum"] += m["forecast"]
            categories[cat]["has_forecast"] = True

    result: list[dict[str, Any]] = []
    for cat_data in categories.values():
        current = cat_data["current_month"]
        previous = cat_data["previous_month"]
        forecast = cat_data["forecast_sum"] if cat_data["has_forecast"] else None

        mom_change = current - previous
        mom_pct = (mom_change / previous * 100.0) if previous != 0 else 0.0

        var_change: float | None = None
        var_pct: float | None = None
        if forecast is not None:
            var_change = current - forecast
            var_pct = (var_change / forecast * 100.0) if forecast != 0 else 0.0

        result.append({
            "category": cat_data["category"],
            "current_month": current,
            "previous_month": previous,
            "forecast": forecast,
            "mom_change": mom_change,
            "mom_pct": mom_pct,
            "var_change": var_change,
            "var_pct": var_pct,
        })

    result.sort(key=lambda r: r["current_month"], reverse=True)
    return result


# ---------------------------------------------------------------------------
# Top movers
# ---------------------------------------------------------------------------

def find_top_movers(
    app_metrics: list[dict[str, Any]],
    n: int = 3,
) -> dict[str, list[dict[str, Any]]]:
    """Identify the top N apps with largest spend changes vs forecast.

    Only considers apps that have a forecast value.

    Args:
        app_metrics: Per-app metrics from :func:`build_app_metrics`.
        n: Number of top movers in each direction (default 3).

    Returns:
        A dict with keys ``"increases"`` and ``"decreases"``, each a list
        of app metric dicts sorted by absolute variance change.
    """
    with_forecast = [m for m in app_metrics if m["var_change"] is not None]

    increases = sorted(
        [m for m in with_forecast if m["var_change"] > 0],
        key=lambda m: m["var_change"],
        reverse=True,
    )[:n]

    decreases = sorted(
        [m for m in with_forecast if m["var_change"] < 0],
        key=lambda m: m["var_change"],
    )[:n]

    return {
        "increases": increases,
        "decreases": decreases,
    }


def find_top_movers_mom(
    app_metrics: list[dict[str, Any]],
    n: int = 3,
) -> dict[str, list[dict[str, Any]]]:
    """Identify the top N apps with largest MoM spend changes.

    Args:
        app_metrics: Per-app metrics from :func:`build_app_metrics`.
        n: Number of top movers in each direction (default 3).

    Returns:
        A dict with keys ``"increases"`` and ``"decreases"``, each a list
        of app metric dicts sorted by absolute MoM dollar change.
    """
    increases = sorted(
        [m for m in app_metrics if m["mom_change"] > 0],
        key=lambda m: m["mom_change"],
        reverse=True,
    )[:n]

    decreases = sorted(
        [m for m in app_metrics if m["mom_change"] < 0],
        key=lambda m: m["mom_change"],
    )[:n]

    return {
        "increases": increases,
        "decreases": decreases,
    }


# ---------------------------------------------------------------------------
# Totals row builder
# ---------------------------------------------------------------------------

def compute_totals(
    rows: list[dict[str, Any]],
    label_key: str = "app",
) -> dict[str, Any]:
    """Compute a totals row from a list of metric dicts.

    Args:
        rows: List of metric dicts (app-level or category-level).
        label_key: The key used for the row label ("app" or "category").

    Returns:
        A single dict with summed values and "Total" as the label.
    """
    total_current = sum(r["current_month"] for r in rows)
    total_previous = sum(r["previous_month"] for r in rows)
    total_two_months_ago = sum(r.get("two_months_ago", 0.0) for r in rows)

    forecasts = [r["forecast"] for r in rows if r["forecast"] is not None]
    total_forecast = sum(forecasts) if forecasts else None

    mom_change = total_current - total_previous
    mom_pct = (mom_change / total_previous * 100.0) if total_previous != 0 else 0.0

    prev_vs_prior_change = total_previous - total_two_months_ago
    prev_vs_prior_pct = (prev_vs_prior_change / total_two_months_ago * 100.0) if total_two_months_ago != 0 else 0.0

    var_change: float | None = None
    var_pct: float | None = None
    if total_forecast is not None:
        var_change = total_current - total_forecast
        var_pct = (var_change / total_forecast * 100.0) if total_forecast != 0 else 0.0

    return {
        label_key: "Total",
        "category": "Total",
        "current_month": total_current,
        "previous_month": total_previous,
        "two_months_ago": total_two_months_ago,
        "forecast": total_forecast,
        "mom_change": mom_change,
        "mom_pct": mom_pct,
        "prev_vs_prior_change": prev_vs_prior_change,
        "prev_vs_prior_pct": prev_vs_prior_pct,
        "var_change": var_change,
        "var_pct": var_pct,
    }


# ---------------------------------------------------------------------------
# COGS drill-down metrics (Phase 3)
# ---------------------------------------------------------------------------

_PRODUCT_NAME_ABBREVIATIONS: dict[str, str] = {
    "Elastic Compute Cloud": "EC2",
    "Simple Storage Service": "S3",
}


def build_drilldown_metrics(
    drilldown_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Enrich raw drilldown rows with labels and MoM percentage.

    Args:
        drilldown_rows: Raw rows from :func:`fetch_cogs_drilldown`.

    Returns:
        A list of enriched dicts sorted by ABS(delta_cost) DESC, each with
        additional keys: ``label``, ``mom_pct``, ``mom_label``.
    """
    metrics: list[dict[str, Any]] = []

    for row in drilldown_rows:
        parts = [row["awn_app"]]
        if row.get("product_name"):
            product_name = row["product_name"]
            parts.append(
                _PRODUCT_NAME_ABBREVIATIONS.get(product_name, product_name)
            )
        if row.get("operation"):
            parts.append(row["operation"])
        label = " | ".join(parts)

        current = row["current_month"]
        previous = row["previous_month"]
        delta = row["delta_cost"]

        mom_label: str | None = None
        if previous == 0 and current > 0:
            mom_label = "NEW"
            mom_pct = 0.0
        elif current == 0 and previous > 0:
            mom_label = "REMOVED"
            mom_pct = -100.0
        elif previous != 0:
            mom_pct = (delta / previous) * 100.0
        else:
            mom_pct = 0.0

        metrics.append({
            **row,
            "label": label,
            "mom_pct": mom_pct,
            "mom_label": mom_label,
        })

    metrics.sort(key=lambda m: abs(m["delta_cost"]), reverse=True)
    return metrics


def build_ec2_purchase_metrics(
    ec2_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Enrich raw EC2 purchase breakdown rows with labels and MoM percentage.

    Args:
        ec2_rows: Raw rows from :func:`fetch_ec2_purchase_breakdown`.

    Returns:
        A list of enriched dicts sorted by ABS(delta_cost) DESC, each with
        additional keys: ``label``, ``mom_pct``, ``mom_label``.
    """
    metrics: list[dict[str, Any]] = []

    for row in ec2_rows:
        parts = [row["purchase_option"]] if row.get("purchase_option") else []
        if row.get("region"):
            parts.append(row["region"])
        label = " | ".join(parts) if parts else "Other"

        current = row["current_month"]
        previous = row["previous_month"]
        delta = row["delta_cost"]

        mom_label: str | None = None
        if previous == 0 and current > 0:
            mom_label = "NEW"
            mom_pct = 0.0
        elif current == 0 and previous > 0:
            mom_label = "REMOVED"
            mom_pct = -100.0
        elif previous != 0:
            mom_pct = (delta / previous) * 100.0
        else:
            mom_pct = 0.0

        metrics.append({
            **row,
            "label": label,
            "mom_pct": mom_pct,
            "mom_label": mom_label,
        })

    metrics.sort(key=lambda m: abs(m["delta_cost"]), reverse=True)
    return metrics


# ---------------------------------------------------------------------------
# Data Platform drill-down
# ---------------------------------------------------------------------------

def build_data_platform_data(
    app_history: list[dict[str, Any]],
) -> dict[str, Any]:
    """Build Data Platform table and chart data from filtered app-level COGS history.

    Args:
        app_history: Rows from ``fetch_app_cogs_history()`` already filtered
            to Data Platform apps. Each dict has keys ``app_name``,
            ``month_label``, and ``monthly_cost``.

    Returns:
        A dict with keys:
            - ``table_rows``: per-app dicts with ``app``, ``monthly_values``,
              ``mom_change``, and ``mom_pct``.
            - ``month_labels``: list of formatted label strings (e.g. ``"SEP-2025"``).
            - ``table_totals``: totals row dict.
            - ``total_history``: chronological list of ``{month_label, actual,
              forecast}`` dicts suitable for ``build_trend_chart()``.
    """
    from collections import defaultdict

    if not app_history:
        return {}

    # Determine chronological month order from the data
    month_order: list[str] = []
    seen: set[str] = set()
    for row in app_history:
        if row["month_label"] not in seen:
            month_order.append(row["month_label"])
            seen.add(row["month_label"])

    def _format_month_label(label: str) -> str:
        """Convert "Sep 2025" -> "SEP-2025"."""
        parts = label.split()
        if len(parts) == 2:
            return f"{parts[0].upper()}-{parts[1]}"
        return label.upper()

    formatted_month_labels = [_format_month_label(m) for m in month_order]

    # Pivot: app_name -> month_label -> cost
    app_month_cost: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for row in app_history:
        app_month_cost[row["app_name"]][row["month_label"]] += row["monthly_cost"]

    # Build per-app rows
    table_rows: list[dict[str, Any]] = []
    for app_name in sorted(app_month_cost.keys()):
        monthly_values = [app_month_cost[app_name].get(m, 0.0) for m in month_order]
        current = monthly_values[-1] if monthly_values else 0.0
        previous = monthly_values[-2] if len(monthly_values) >= 2 else 0.0
        mom_change = current - previous
        mom_pct = (mom_change / previous * 100.0) if previous != 0 else 0.0

        table_rows.append({
            "app": app_name,
            "monthly_values": monthly_values,
            "mom_change": mom_change,
            "mom_pct": mom_pct,
        })

    # Sort by current-month spend descending
    table_rows.sort(
        key=lambda r: r["monthly_values"][-1] if r["monthly_values"] else 0.0,
        reverse=True,
    )

    # Compute column totals
    n_months = len(month_order)
    total_monthly = [
        sum(r["monthly_values"][i] for r in table_rows if i < len(r["monthly_values"]))
        for i in range(n_months)
    ]
    total_current = total_monthly[-1] if total_monthly else 0.0
    total_previous = total_monthly[-2] if len(total_monthly) >= 2 else 0.0
    total_mom_change = total_current - total_previous
    total_mom_pct = (total_mom_change / total_previous * 100.0) if total_previous != 0 else 0.0

    table_totals: dict[str, Any] = {
        "app": "Total",
        "monthly_values": total_monthly,
        "mom_change": total_mom_change,
        "mom_pct": total_mom_pct,
    }

    # Build chronological total history for the line chart
    total_history = [
        {"month_label": m, "actual": v, "forecast": None}
        for m, v in zip(month_order, total_monthly)
    ]

    return {
        "table_rows": table_rows,
        "month_labels": formatted_month_labels,
        "table_totals": table_totals,
        "total_history": total_history,
    }


def build_dbx_breakdown_metrics(
    rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Enrich raw Databricks breakdown rows with MoM percentage.

    Args:
        rows: Raw rows from :func:`fetch_dbx_awn_breakdown`.

    Returns:
        A list of enriched dicts sorted by current_month DESC, each with
        additional keys: ``mom_pct``, ``mom_label``.
    """
    metrics: list[dict[str, Any]] = []

    for row in rows:
        current = row["current_month"]
        previous = row["previous_month"]
        delta = row["delta_cost"]

        mom_label: str | None = None
        if previous == 0 and current > 0:
            mom_label = "NEW"
            mom_pct = 0.0
        elif current == 0 and previous > 0:
            mom_label = "REMOVED"
            mom_pct = -100.0
        elif previous != 0:
            mom_pct = (delta / previous) * 100.0
        else:
            mom_pct = 0.0

        metrics.append({
            **row,
            "mom_pct": mom_pct,
            "mom_label": mom_label,
        })

    metrics.sort(key=lambda m: m["current_month"], reverse=True)
    return metrics


def compute_drilldown_totals(
    rows: list[dict[str, Any]],
) -> dict[str, Any]:
    """Compute totals across all drilldown/EC2 breakdown rows.

    Args:
        rows: List of enriched drilldown or EC2 metric dicts.

    Returns:
        A dict with ``current_month``, ``previous_month``, ``delta_cost``,
        and ``mom_pct``.
    """
    total_current = sum(r["current_month"] for r in rows)
    total_previous = sum(r["previous_month"] for r in rows)
    total_delta = sum(r["delta_cost"] for r in rows)
    mom_pct = (total_delta / total_previous * 100.0) if total_previous != 0 else 0.0

    return {
        "current_month": total_current,
        "previous_month": total_previous,
        "delta_cost": total_delta,
        "mom_pct": mom_pct,
    }


# ---------------------------------------------------------------------------
# Data Lake drill-down
# ---------------------------------------------------------------------------

def build_data_lake_data(
    app_history: list[dict[str, Any]],
) -> dict[str, Any]:
    """Build Data Lake table data from app-level history.

    Accepts rows from ``fetch_data_lake_history()`` (AWS apps + AWN Databricks).
    Returns the same shape as ``build_data_platform_data``.
    """
    return build_data_platform_data(app_history)
