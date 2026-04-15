"""CLI script to generate a FinOps PowerPoint report.

Usage:
    uv run python generate_report.py                # defaults to previous month
    uv run python generate_report.py --month February --year 2026
    uv run python generate_report.py -m March -y 2026 -o my_report.pptx
"""

from __future__ import annotations

import argparse
import calendar
import sys
from datetime import date, timedelta
from pathlib import Path

from src.app_data import (
    build_app_metrics,
    build_category_rollup,
    build_data_platform_data,
    build_dbx_breakdown_metrics,
    build_drilldown_metrics,
    build_ec2_purchase_metrics,
    compute_drilldown_totals,
    compute_totals,
    find_top_movers,
    find_top_movers_mom,
    load_app_category_mapping,
)
from src.calculations import calculate_all_buckets
from src.charts import (
    build_all_trend_charts,
    build_cogs_stacked_bar,
    build_cost_per_1m_trend,
    build_pod_unit_cost_trend,
    build_trend_chart,
    build_unit_cost_dual_axis,
    render_chart_png,
)
from src.forecast import load_app_forecasts, load_forecast, load_forecast_history
from src.ingestion import (
    close_shared_connection,
    fetch_app_actuals,
    fetch_app_cogs_history,
    fetch_bucket_actuals,
    fetch_bucket_history,
    fetch_cogs_drilldown,
    fetch_cylance_actuals,
    fetch_cylance_dbx_summary,
    fetch_dbx_awn_breakdown,
    fetch_dbx_awn_summary,
    fetch_ec2_purchase_breakdown,
    fetch_unit_cost_data,
    load_config,
)
from src.narrative import generate_all_narratives
from src.pptx_gen import generate_pptx


def _month_name_to_number(name: str) -> str:
    idx = list(calendar.month_name).index(name)
    return f"{idx:02d}"


def _build_summary_table_data(
    actuals: dict,
    month_num: str,
    year: int,
) -> dict:
    """Build the data dict for the executive summary MoM comparison table.

    Always shows the two most recently completed months: the reporting month
    (``current_month``, highlighted yellow) vs the month before it
    (``previous_month``).  Daily rates are normalised by each month's day count.

    Args:
        actuals: Bucket actuals dict from fetch_bucket_actuals.
        month_num: Zero-padded current reporting month (e.g. ``"03"``).
        year: Four-digit reporting year.

    Returns:
        Dict with ``col_older``, ``col_newer``, and ``rows`` (list of row dicts).
    """
    m = int(month_num)

    prev_m = m - 1 if m > 1 else 12
    prev_y = year if m > 1 else year - 1

    curr_days = calendar.monthrange(year, m)[1]
    prev_days = calendar.monthrange(prev_y, prev_m)[1]

    curr_label = f"{calendar.month_abbr[m]} {year}"
    prev_label = f"{calendar.month_abbr[prev_m]} {prev_y}"

    cogs_newer = actuals["COGS"]["current_month"]
    cogs_older = actuals["COGS"]["previous_month"]
    opex_newer = actuals["OpEx"]["current_month"]
    opex_older = actuals["OpEx"]["previous_month"]

    total_newer = cogs_newer + opex_newer
    total_older = cogs_older + opex_older

    cogs_daily_newer = cogs_newer / curr_days
    cogs_daily_older = cogs_older / prev_days
    opex_daily_newer = opex_newer / curr_days
    opex_daily_older = opex_older / prev_days

    def _pct(new: float, old: float) -> float:
        return (new - old) / old * 100.0 if old != 0 else 0.0

    return {
        "col_older": prev_label,
        "col_newer": curr_label,
        "rows": [
            {
                "metric": "COGS",
                "older": cogs_older,
                "newer": cogs_newer,
                "change_dollar": cogs_newer - cogs_older,
                "change_pct": _pct(cogs_newer, cogs_older),
                "row_type": "currency",
                "is_bold": False,
            },
            {
                "metric": "OPEX",
                "older": opex_older,
                "newer": opex_newer,
                "change_dollar": opex_newer - opex_older,
                "change_pct": _pct(opex_newer, opex_older),
                "row_type": "currency",
                "is_bold": False,
            },
            {
                "metric": "Grand Total",
                "older": total_older,
                "newer": total_newer,
                "change_dollar": total_newer - total_older,
                "change_pct": _pct(total_newer, total_older),
                "row_type": "currency",
                "is_bold": True,
            },
            {
                "metric": "COGS Daily Rate",
                "older": cogs_daily_older,
                "newer": cogs_daily_newer,
                "change_dollar": cogs_daily_newer - cogs_daily_older,
                "change_pct": _pct(cogs_daily_newer, cogs_daily_older),
                "row_type": "daily_rate",
                "is_bold": False,
            },
            {
                "metric": "OPEX Daily Rate",
                "older": opex_daily_older,
                "newer": opex_daily_newer,
                "change_dollar": opex_daily_newer - opex_daily_older,
                "change_pct": _pct(opex_daily_newer, opex_daily_older),
                "row_type": "daily_rate",
                "is_bold": False,
            },
        ],
    }


def _build_kpi_grid_data(
    awn_actuals: dict,
    cylance_actuals: dict,
    dbx_summary: dict,
    cylance_dbx: dict,
) -> dict:
    """Build 3x3 KPI grid data for COGS, OpEx, and Total overview slides.

    Returns a dict keyed by bucket ("COGS", "OpEx", "Total").  Each value is
    a 3-row × 3-col list of cell dicts with keys:
        ``value``  — current-month cost
        ``prev``   — previous-month cost
        ``change`` — absolute MoM change
        ``pct``    — percentage MoM change
    Row order: AWN, Cylance, Total.
    Col order: AWS, Databricks, Total.
    """
    def _cell(curr: float, prev: float) -> dict:
        change = curr - prev
        pct = (change / prev * 100.0) if prev != 0 else 0.0
        return {"value": curr, "prev": prev, "change": change, "pct": pct}

    grid: dict = {}

    for bucket in ("COGS", "OpEx"):
        awn_aws   = _cell(awn_actuals[bucket]["current_month"],     awn_actuals[bucket]["previous_month"])
        awn_dbx   = _cell(dbx_summary[bucket]["current_month"],     dbx_summary[bucket]["previous_month"])
        cyl_aws   = _cell(cylance_actuals[bucket]["current_month"], cylance_actuals[bucket]["previous_month"])
        cyl_dbx   = _cell(cylance_dbx[bucket]["current_month"],     cylance_dbx[bucket]["previous_month"])

        awn_total = _cell(awn_aws["value"]   + awn_dbx["value"],   awn_aws["prev"]   + awn_dbx["prev"])
        cyl_total = _cell(cyl_aws["value"]   + cyl_dbx["value"],   cyl_aws["prev"]   + cyl_dbx["prev"])
        tot_aws   = _cell(awn_aws["value"]   + cyl_aws["value"],   awn_aws["prev"]   + cyl_aws["prev"])
        tot_dbx   = _cell(awn_dbx["value"]   + cyl_dbx["value"],   awn_dbx["prev"]   + cyl_dbx["prev"])
        grand     = _cell(awn_total["value"] + cyl_total["value"], awn_total["prev"] + cyl_total["prev"])

        grid[bucket] = [
            [awn_aws,  awn_dbx,  awn_total],   # AWN row
            [cyl_aws,  cyl_dbx,  cyl_total],   # Cylance row
            [tot_aws,  tot_dbx,  grand],        # Total row
        ]

    total_grid = []
    for r in range(3):
        row = []
        for c in range(3):
            v = grid["COGS"][r][c]["value"] + grid["OpEx"][r][c]["value"]
            p = grid["COGS"][r][c]["prev"]  + grid["OpEx"][r][c]["prev"]
            row.append(_cell(v, p))
        total_grid.append(row)
    grid["Total"] = total_grid

    return grid


def _merge_history(
    actuals_history: dict[str, list[dict]],
    forecast_history: dict[str, list[dict]],
    bucket_configs: list[dict],
) -> dict[str, list[dict]]:
    merged: dict[str, list[dict]] = {}
    for bucket in bucket_configs:
        bucket_name = bucket["name"]
        forecast_key = bucket["forecast_mapping_key"]
        actual_list = actuals_history.get(bucket_name, [])
        forecast_list = forecast_history.get(forecast_key, [])
        fc_lookup = {e["month_label"]: e.get("forecast") for e in forecast_list}
        merged[bucket_name] = [
            {
                "month_label": e["month_start"],
                "actual": e["actual"],
                "forecast": fc_lookup.get(e["month_start"]),
            }
            for e in actual_list
        ]
    return merged


def main(month_name: str, year: int, output_path: str | None) -> None:
    config = load_config("config.yaml")
    buckets = config["buckets"]
    lookback_months = config.get("lookback_months", 6)
    month_num = _month_name_to_number(month_name)
    month_label = f"{month_name} {year}"

    if output_path is None:
        safe = month_label.replace(" ", "_").lower()
        output_path = f"output/finops_report_{safe}.pptx"

    print(f"Generating report for {month_label} ...")

    # 1. Actuals
    print("  Fetching actuals from Redshift ...")
    actuals = fetch_bucket_actuals(month_num, year)
    for name, vals in actuals.items():
        print(f"    {name}: ${vals['current_month']:,.2f} current / ${vals['previous_month']:,.2f} previous")

    # 2. Forecasts
    print("  Loading forecasts ...")
    try:
        forecasts = load_forecast(config, month_name, year)
    except FileNotFoundError:
        print("    Forecast file not found -- continuing without forecast data.")
        forecasts = {b["forecast_mapping_key"]: None for b in buckets}

    # 3. Metrics
    print("  Calculating metrics ...")
    all_metrics = calculate_all_buckets(actuals, forecasts, buckets)

    # 4. Narratives
    print("  Generating narratives ...")
    narratives = generate_all_narratives(all_metrics, month_label, config)

    # 5. History + charts
    print(f"  Fetching {lookback_months}-month history ...")
    actuals_history = fetch_bucket_history(month_num, year, num_months=lookback_months)

    print("  Loading forecast history ...")
    try:
        forecast_history = load_forecast_history(config, month_name, year)
    except Exception:
        print("    Forecast history not available -- charts will show actuals only.")
        forecast_history = {}

    chart_data = _merge_history(actuals_history, forecast_history, buckets)

    print("  Rendering trend charts ...")
    sections = config.get("pptx", {}).get("sections", [])
    chart_images = build_all_trend_charts(chart_data, sections)

    # COGS stacked bar chart (per-app breakdown)
    print(f"  Fetching app-level COGS history ({lookback_months} months) ...")
    app_cogs_history = fetch_app_cogs_history(month_num, year, num_months=lookback_months)
    print(f"    Found {len(app_cogs_history)} app-month records")

    if app_cogs_history:
        print("  Rendering COGS stacked bar chart ...")
        cogs_bar_fig = build_cogs_stacked_bar(app_cogs_history)
        chart_images["COGS_app_breakdown"] = render_chart_png(cogs_bar_fig)

    # 6. Phase 2: App-level COGS breakdown
    # (app_category_map loaded below; Data Platform charts built after it)
    print("  Fetching app-level COGS actuals ...")
    app_actuals = fetch_app_actuals(month_num, year)
    print(f"    Found {len(app_actuals)} apps")

    print("  Loading app-level forecasts ...")
    app_forecasts = load_app_forecasts(config, month_name, year)
    print(f"    Loaded forecasts for {len(app_forecasts)} apps")

    print("  Loading app category mapping ...")
    app_category_map = load_app_category_mapping(config)
    print(f"    Mapped {len(app_category_map)} apps to categories")

    # Data Platform section
    print("  Building Data Platform data ...")
    data_platform_app_names = {
        app for app, cat in app_category_map.items() if cat == "Data Platform"
    }
    data_platform_history = [
        r for r in app_cogs_history if r["app_name"] in data_platform_app_names
    ]
    data_platform_data = build_data_platform_data(data_platform_history)
    if data_platform_data:
        print(f"    {len(data_platform_data['table_rows'])} Data Platform apps found")
        dp_trend_fig = build_trend_chart(
            "Data Platform", data_platform_data["total_history"], dark_mode=True
        )
        chart_images["data_platform_total"] = render_chart_png(dp_trend_fig)
        dp_bar_fig = build_cogs_stacked_bar(
            data_platform_history, top_n=len(data_platform_app_names)
        )
        chart_images["data_platform_stacked"] = render_chart_png(dp_bar_fig)
    else:
        data_platform_data = None

    print("  Building executive summary table data ...")
    summary_data = _build_summary_table_data(actuals, month_num, year)

    print("  Fetching Cylance actuals from Redshift ...")
    cylance_actuals = fetch_cylance_actuals(month_num, year)
    for name, vals in cylance_actuals.items():
        print(f"    Cylance {name}: ${vals['current_month']:,.2f} current / ${vals['previous_month']:,.2f} previous")
    cylance_summary_data = _build_summary_table_data(cylance_actuals, month_num, year)

    print("  Fetching Cylance Databricks spend ...")
    cylance_dbx = fetch_cylance_dbx_summary(month_num, year)
    for name, vals in cylance_dbx.items():
        print(f"    Cylance DBX {name}: ${vals['current_month']:,.2f} current / ${vals['previous_month']:,.2f} previous")

    cylance_dbx_summary_data = _build_summary_table_data(cylance_dbx, month_num, year)

    print("  Fetching AWN Databricks spend ...")
    dbx_summary = fetch_dbx_awn_summary(month_num, year)
    for name, vals in dbx_summary.items():
        print(f"    DBX {name}: ${vals['current_month']:,.2f} current / ${vals['previous_month']:,.2f} previous")
    dbx_summary_data = _build_summary_table_data(dbx_summary, month_num, year)

    print("  Building KPI grid data ...")
    kpi_grid_data = _build_kpi_grid_data(actuals, cylance_actuals, dbx_summary, cylance_dbx)

    print("  Building app metrics ...")
    app_metrics = build_app_metrics(app_actuals, app_forecasts, app_category_map)
    category_rollup = build_category_rollup(app_metrics)
    top_movers = find_top_movers(app_metrics)
    top_movers_mom = find_top_movers_mom(app_metrics)
    app_totals = compute_totals(app_metrics, label_key="app")
    category_totals = compute_totals(category_rollup, label_key="category")

    app_data = {
        "app_metrics": app_metrics,
        "category_rollup": category_rollup,
        "top_movers": top_movers,
        "top_movers_mom": top_movers_mom,
        "app_totals": app_totals,
        "category_totals": category_totals,
    }

    # 7. Phase 3: COGS drill-down
    print("  Fetching COGS drill-down data ...")
    drilldown_raw = fetch_cogs_drilldown(month_num, year)
    print(f"    Found {len(drilldown_raw)} drill-down rows")

    drilldown_metrics = build_drilldown_metrics(drilldown_raw)
    drilldown_totals = compute_drilldown_totals(drilldown_metrics)

    print("  Fetching EC2 purchase option breakdown ...")
    ec2_raw = fetch_ec2_purchase_breakdown(month_num, year)
    print(f"    Found {len(ec2_raw)} EC2 breakdown rows")

    ec2_metrics = build_ec2_purchase_metrics(ec2_raw)
    ec2_totals = compute_drilldown_totals(ec2_metrics)

    app_data["drilldown"] = {
        "drilldown_metrics": drilldown_metrics,
        "drilldown_totals": drilldown_totals,
        "ec2_metrics": ec2_metrics,
        "ec2_totals": ec2_totals,
    }

    # 8. Unit cost data (COGS vs analyzed observations)
    print("  Fetching unit cost data ...")
    unit_cost_data = fetch_unit_cost_data(month_num, year, num_months=lookback_months)
    print(f"    Org history: {len(unit_cost_data['org_history'])} months")
    print(f"    Pod MoM rows: {len(unit_cost_data['pod_mom'])}")

    if unit_cost_data["org_history"]:
        print("  Rendering unit cost charts ...")
        chart_images["unit_cost_overview"] = render_chart_png(
            build_unit_cost_dual_axis(unit_cost_data["org_history"])
        )
        chart_images["cost_per_1m_trend"] = render_chart_png(
            build_cost_per_1m_trend(unit_cost_data["org_history"])
        )
    if unit_cost_data["pod_history"]:
        chart_images["pod_unit_cost_trend"] = render_chart_png(
            build_pod_unit_cost_trend(unit_cost_data["pod_history"])
        )

    # 9. AWN Databricks breakdowns (summary already fetched above)
    print("  Fetching AWN Databricks breakdowns ...")
    dbx_cogs_by_workspace = build_dbx_breakdown_metrics(
        fetch_dbx_awn_breakdown(month_num, year, "workspace_name", "cogs")
    )
    dbx_cogs_by_sku = build_dbx_breakdown_metrics(
        fetch_dbx_awn_breakdown(month_num, year, "sku_name", "cogs")
    )
    dbx_opex_by_workspace = build_dbx_breakdown_metrics(
        fetch_dbx_awn_breakdown(month_num, year, "workspace_name", "opex")
    )
    dbx_opex_by_sku = build_dbx_breakdown_metrics(
        fetch_dbx_awn_breakdown(month_num, year, "sku_name", "opex")
    )

    dbx_data = {
        "cogs_by_workspace": dbx_cogs_by_workspace,
        "cogs_by_workspace_totals": compute_drilldown_totals(dbx_cogs_by_workspace),
        "cogs_by_sku": dbx_cogs_by_sku,
        "cogs_by_sku_totals": compute_drilldown_totals(dbx_cogs_by_sku),
        "opex_by_workspace": dbx_opex_by_workspace,
        "opex_by_workspace_totals": compute_drilldown_totals(dbx_opex_by_workspace),
        "opex_by_sku": dbx_opex_by_sku,
        "opex_by_sku_totals": compute_drilldown_totals(dbx_opex_by_sku),
    }
    print(f"    COGS workspaces: {len(dbx_cogs_by_workspace)}, SKUs: {len(dbx_cogs_by_sku)}")
    print(f"    OpEx workspaces: {len(dbx_opex_by_workspace)}, SKUs: {len(dbx_opex_by_sku)}")

    # 10. Build PPTX
    print("  Building PowerPoint ...")
    pptx_bytes = generate_pptx(
        narratives,
        all_metrics,
        config,
        month_label=month_label,
        chart_images=chart_images,
        app_data=app_data,
        summary_data=summary_data,
        cylance_summary_data=cylance_summary_data,
        cylance_dbx_summary_data=cylance_dbx_summary_data,
        unit_cost_data=unit_cost_data,
        data_platform_data=data_platform_data,
        dbx_data=dbx_data,
        dbx_summary_data=dbx_summary_data,
        kpi_grid_data=kpi_grid_data,
    )

    # 11. Write file
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(pptx_bytes)
    print(f"Done! Saved to {out}")

    close_shared_connection()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a FinOps PowerPoint report.")
    today = date.today()
    first_of_this_month = today.replace(day=1)
    prev_month_date = first_of_this_month - timedelta(days=1)
    default_month = calendar.month_name[prev_month_date.month]
    default_year = prev_month_date.year
    parser.add_argument("-m", "--month", default=default_month, help=f"Month name (default: {default_month})")
    parser.add_argument("-y", "--year", type=int, default=default_year, help=f"Year (default: {default_year})")
    parser.add_argument("-o", "--output", default=None, help="Output file path (default: output/finops_report_<month>_<year>.pptx)")
    args = parser.parse_args()

    if args.month not in list(calendar.month_name)[1:]:
        print(f"Error: '{args.month}' is not a valid month name. Use full name like 'January'.", file=sys.stderr)
        sys.exit(1)

    main(args.month, args.year, args.output)
