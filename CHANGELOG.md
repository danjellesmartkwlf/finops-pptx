# Changelog

## [0.2.0] — 2026-05-15

### Changed

- **Refactored `generate_report.py`:** Broke up the monolithic `main()` (~200 lines) into 8 named phase functions (~50-line orchestrator). No logic changes — purely structural improvement for debuggability and testability.
  - `_fetch_actuals` — Redshift fetches (AWN, Cylance, DBX)
  - `_load_forecasts` — bucket + app-level Excel forecasts
  - `_build_metrics_and_narratives` — variance calculations + narrative text
  - `_build_history_and_charts` — trend history + chart rendering
  - `_build_app_level_data` — app metrics, summaries, KPI grid, data platform/lake
  - `_build_drilldowns` — COGS drilldown, EC2, other app breakdowns
  - `_build_unit_cost_charts` — unit cost data + charts
  - `_build_dbx_breakdowns` — AWN Databricks workspace/SKU breakdowns

## [0.1.0] — Initial release

- End-to-end report generation pipeline: Redshift ingestion, forecast loading, variance calculations, narrative generation, chart rendering, and PowerPoint export.
