# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

FinOps Report Generator — automates the monthly executive cloud cost presentation. Ingests actuals from Redshift and forecasts from Excel, calculates MoM and Forecast Variance for Total/COGS/OpEx buckets, generates narrative text from templates, and exports a PowerPoint deck.

See `PRD-PHASE1.md` for full requirements and `DATA.md` for Redshift table schemas and business logic.

## Commands

- **Install dependencies:** `uv sync`
- **Add a dependency:** `uv add <package>`
- **Generate report (CLI):** `uv run python generate_report.py`
  - Options: `--month February --year 2026 -o my_report.pptx` (defaults to previous month)

## Development

- Use **UV** for all Python management (not pip/venv directly)
- Python 3.12 (pinned in `.python-version`)
- Use subagents in parallel where possible

## Architecture

```
config.yaml          # Bucket definitions, SQL sources, narrative templates, forecast file paths, drilldown config
slides_config.yaml   # Slide sequence and layout definitions (add/remove/reorder slides here)
generate_report.py   # Headless CLI entry point
src/
  ingestion.py       # Redshift connection, SQL builders, all fetch_* functions
  forecast.py        # Excel forecast loading (two paths: legacy row-label and new data_files)
  calculations.py    # MoM and Forecast Variance math
  narrative.py       # Template-based narrative generation (Mode A / Mode B)
  app_data.py        # App-level COGS processing: category rollups, top movers, drill-down enrichment
  charts.py          # Plotly chart builders; render_chart_png() converts to PNG for PPTX
  pptx_gen.py        # Config-driven slide builder; dispatches slide types from slides_config.yaml
  pptx_utils.py      # Formatting, layout helpers, and table utilities for pptx_gen.py
pptx_template/       # Corporate .pptx master template (template1.pptx) + index YAML
data/
  forecasts/         # COGS and OpEx forecast Excel files (separate files, columnar by month)
  mapping/           # Reference data (e.g., app_category_mapping.xlsx)
```

## Key Design Patterns

**Two-config system:** `config.yaml` owns data logic (buckets, SQL filters, narrative templates, file paths). `slides_config.yaml` owns presentation structure (slide order, types, layout names). Adding or reordering slides only requires editing `slides_config.yaml`.

**Forecast two-path system:** `src/forecast.py` supports a legacy single-file format (row labels in col A, `forecast.mapping` in config) and the current format (separate COGS/OpEx Excel files under `config["data_files"]["forecasts"]`). The current path is activated when `data_files.forecasts` exists in config.

**Slide renderer dispatch:** `pptx_gen.py` reads `slides_config.yaml` and dispatches each entry by `type` (`title`, `transition`, `content`, `chart`, `table`, `split_table`). The `requires` key on a slide entry makes it conditional on data being present.

## Data & MCP

A Redshift MCP server (`awn-redshift-mcp`) is configured in `.mcp.json` for read-only access. Connection credentials come from environment variables (`REDSHIFT_HOST`, `REDSHIFT_PORT`, `REDSHIFT_DATABASE`, `REDSHIFT_USER`, `REDSHIFT_PASSWORD`, `REDSHIFT_SCHEMA`).

Key Redshift tables (all in `public` schema):
- `daily_cur_summary` — Arctic Wolf AWS CUR
- `cylance_daily_cur_summary` — Cylance AWS CUR
- `dbx_cur` — Databricks CUR for both Arctic Wolf and Cylance (filter on `organization`)

See `DATA.md` for column details, COGS/OpEx logic, common filters, and Databricks cost attribution rules.

## Domain Concepts

- **Buckets:** Total, COGS, OpEx — the three cost categories tracked in every report
- **COGS vs OpEx:** Determined by account, tags (`tag_opex_in_prod`), and service rules. Some services are always OpEx regardless of account (see DATA.md)
- **Narrative modes:** Mode A (forecast exists) includes variance; Mode B (no forecast) shows MoM only
- **Forecast:** Comes from external Excel files mapped by config keys, not from Redshift
