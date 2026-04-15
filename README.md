# FinOps Report Generator

Automates the monthly executive cloud cost presentation for Arctic Wolf. Replaces manual data gathering with a pipeline that ingests actuals from Redshift and forecasts from Excel, calculates variances, generates narrative text, and exports a PowerPoint deck.

## What It Does

1. **Ingests actuals** from Redshift -- Arctic Wolf AWS CUR, Cylance AWS CUR, and Databricks CUR (both orgs).
2. **Ingests forecasts** from Excel files (separate files for COGS and OpEx).
3. **Calculates** Month-over-Month and Forecast Variance for Total, COGS, and OpEx buckets.
4. **Generates narrative text** from pre-approved templates (Mode A with forecast variance, Mode B with MoM only).
5. **Exports a PowerPoint deck** using the corporate template.

## Tech Stack

- Python 3.12
- UV (package management)
- psycopg2-binary (Redshift)
- python-pptx (PowerPoint generation)
- openpyxl (Excel ingestion)
- Plotly (charts)
- PyYAML (configuration)
- python-dotenv (environment variables)

## Project Structure

```
config.yaml                # Bucket definitions, SQL logic, templates, file paths
slides_config.yaml         # Slide sequence and layout definitions
generate_report.py         # CLI entry point (defaults to previous month)
src/
  ingestion.py             # Redshift connection and data loading
  forecast.py              # Excel forecast ingestion and validation
  calculations.py          # MoM and Forecast Variance calculations
  narrative.py             # Template-based narrative generation
  pptx_gen.py              # PowerPoint slide builder
data/
  forecasts/               # COGS and OpEx forecast Excel files
  mapping/                 # Reference data (e.g., app_category_mapping.xlsx)
pptx_template/             # Corporate PowerPoint template
```

## Setup

1. Install [UV](https://docs.astral.sh/uv/).

2. Install dependencies:

   ```
   uv sync
   ```

3. Copy `.env.example` to `.env` and fill in your Redshift credentials:

   ```
   REDSHIFT_HOST=
   REDSHIFT_PORT=
   REDSHIFT_DATABASE=
   REDSHIFT_USER=
   REDSHIFT_PASSWORD=
   REDSHIFT_SCHEMA=
   ```

## Usage

Generate a report for the previous month (default):

```
uv run python generate_report.py
```

Specify a month and year:

```
uv run python generate_report.py --month February --year 2026
```

Custom output path:

```
uv run python generate_report.py -o my_report.pptx
```

## Configuration

All configuration lives in `config.yaml`:

- **Bucket definitions** -- Total, COGS, OpEx with their SQL logic and tag filters.
- **Narrative templates** -- Mode A (forecast exists) and Mode B (MoM only).
- **Forecast file paths** -- Under `data_files`, so forecast Excel files can be swapped without code changes.
- **PPTX mappings** -- Maps internal metric names to slide IDs and shape placeholder IDs in the corporate template.

## Data Sources

All tables are in the `public` schema in Redshift:

| Table | Description |
|---|---|
| `daily_cur_summary` | Arctic Wolf AWS Cost and Usage Report |
| `cylance_daily_cur_summary` | Cylance AWS Cost and Usage Report |
| `dbx_cur` | Databricks CUR for both Arctic Wolf and Cylance (filter on `organization`) |

See `DATA.md` for column details, COGS/OpEx classification logic, and Databricks cost attribution rules.
