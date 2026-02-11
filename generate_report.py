"""CLI script to generate a FinOps PowerPoint report without Streamlit.

Usage:
    uv run python generate_report.py                # defaults to January 2026
    uv run python generate_report.py --month February --year 2026
    uv run python generate_report.py -m March -y 2026 -o my_report.pptx
"""

from __future__ import annotations

import argparse
import calendar
import sys
from pathlib import Path

from src.calculations import calculate_all_buckets
from src.charts import build_all_trend_charts
from src.forecast import load_forecast, load_forecast_history
from src.ingestion import (
    close_shared_connection,
    fetch_bucket_actuals,
    fetch_bucket_history,
    load_config,
)
from src.narrative import generate_all_narratives
from src.pptx_gen import generate_pptx


def _month_name_to_number(name: str) -> str:
    idx = list(calendar.month_name).index(name)
    return f"{idx:02d}"


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
    print("  Fetching 6-month history ...")
    actuals_history = fetch_bucket_history(month_num, year)

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

    # 6. Build PPTX
    print("  Building PowerPoint ...")
    pptx_bytes = generate_pptx(
        narratives,
        all_metrics,
        config,
        month_label=month_label,
        chart_images=chart_images,
    )

    # 7. Write file
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(pptx_bytes)
    print(f"Done! Saved to {out}")

    close_shared_connection()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a FinOps PowerPoint report.")
    parser.add_argument("-m", "--month", default="January", help="Month name (default: January)")
    parser.add_argument("-y", "--year", type=int, default=2026, help="Year (default: 2026)")
    parser.add_argument("-o", "--output", default=None, help="Output file path (default: output/finops_report_<month>_<year>.pptx)")
    args = parser.parse_args()

    if args.month not in list(calendar.month_name)[1:]:
        print(f"Error: '{args.month}' is not a valid month name. Use full name like 'January'.", file=sys.stderr)
        sys.exit(1)

    main(args.month, args.year, args.output)
