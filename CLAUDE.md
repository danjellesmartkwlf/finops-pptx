# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

FinOps Report Generator — automates the monthly executive cloud cost presentation. Ingests actuals from Redshift and forecasts from Excel, calculates MoM and Forecast Variance for Total/COGS/OpEx buckets, generates narrative text from templates, and outputs a PowerPoint deck. The review UI is Streamlit-based ("The Cockpit").

See `PRD-PHASE1.md` for full requirements and `DATA.md` for Redshift table schemas and business logic.

## Commands

- **Install dependencies:** `uv sync`
- **Add a dependency:** `uv add <package>`
- **Run the app:** `uv run python main.py`
- **Run Streamlit UI (when built):** `uv run streamlit run app.py`

## Development

- Use **UV** for all Python management (not pip/venv directly)
- Python 3.12 (pinned in `.python-version`)
- Use subagents in parallel where possible

## Data & MCP

A Redshift MCP server (`awn-redshift-mcp`) is configured in `.mcp.json` for read-only access. Connection credentials come from environment variables (`REDSHIFT_HOST`, `REDSHIFT_PORT`, `REDSHIFT_DATABASE`, `REDSHIFT_USER`, `REDSHIFT_PASSWORD`, `REDSHIFT_SCHEMA`).

Key Redshift tables (all in `public` schema):
- `daily_cur_summary` — Arctic Wolf AWS CUR
- `cylance_daily_cur_summary` — Cylance AWS CUR
- `dbx_cur` — Databricks CUR for both Arctic Wolf and Cylance (filter on `organization`)

See `DATA.md` for column details, COGS/OpEx logic, common filters, and Databricks cost attribution rules.

## Target Architecture (from PRD)

```
config.yaml          # Bucket definitions, SQL logic, template strings, slide mappings
app.py               # Streamlit UI entry point
src/
  ingestion.py       # Redshift & Excel data loading
  narrative.py       # Template-based narrative generation
  pptx_gen.py        # PowerPoint slide builder
  mcp_client.py      # MCP sidecar for ad-hoc Redshift queries
pptx_template/       # Corporate .pptx master template + index
```

## Domain Concepts

- **Buckets:** Total, COGS, OpEx — the three cost categories tracked in every report
- **COGS vs OpEx:** Determined by account, tags (`tag_opex_in_prod`), and service rules. Some services are always OpEx regardless of account (see DATA.md)
- **Narrative modes:** Mode A (forecast exists) includes variance; Mode B (no forecast) shows MoM only
- **Forecast:** Comes from an external Excel file mapped by config keys, not from Redshift
