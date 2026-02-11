# Data Contracts

This document defines the exact data shapes exchanged between modules in the FinOps Report Generator pipeline, the SQL extraction recommendation, and field-to-column mapping tables for every slide type.

---

## 1. SQL Extraction Recommendation

### Should SQL be extracted to `.sql` files?

**Recommendation: No.** Keep the SQL inline in `ingestion.py`.

### Rationale

1. **Dynamic interpolation everywhere.** Every query in `ingestion.py` uses f-string interpolation for values that change per invocation: `month`, `year`, `noise_floor`, `top_n`, `ec2_product`, `ec2_operation`, config-driven `table`, `cost_column`, `date_column`, and `filters[]`. Extracting to `.sql` template files would require a placeholder syntax (`{month}`, `$noise_floor`, or Jinja2 `{{ var }}`) that adds a layer of indirection without reducing complexity. The current f-strings are immediately readable next to the Python logic that assembles them.

2. **CTE patterns are well-structured already.** The two large queries (`fetch_cogs_drilldown` and `fetch_ec2_purchase_breakdown`) use a clean `grouped -> filtered -> ranked -> combined` CTE chain. This pattern is easy to follow inline and would not become materially clearer in a separate file.

3. **Config-driven filters make templates awkward.** The `filters` list from `config.yaml` is variable-length and joined with `AND`. A `.sql` template would need a loop construct (Jinja2 `{% for %}`) or a single `{filters_clause}` placeholder, neither of which improves readability over the current `" AND ".join(where_parts)` pattern.

4. **Small number of queries.** There are exactly five SQL-producing functions (`_build_simple_query`, `fetch_app_actuals`, `fetch_cogs_drilldown`, `fetch_ec2_purchase_breakdown`, `_build_history_query`). This is a manageable number to keep inline.

5. **Testing considerations.** With inline SQL, unit tests can mock `cursor.execute()` and assert on the generated SQL string directly. External `.sql` files would require loading the template, rendering it, then asserting -- adding a file I/O dependency to tests.

**When extraction _would_ make sense:** If the project grows to 15+ distinct query shapes, or if SQL review by non-Python-fluent DBAs becomes a regular workflow, then a `sql/` directory with Jinja2 templates would be justified. That threshold has not been reached.

---

## 2. Data Shapes Documentation

### 2a. `app_metrics` (list of dicts)

Produced by: `app_data.build_app_metrics()`
Consumed by: `pptx_gen._add_mom_table_slides()`, `pptx_gen._add_forecast_vs_spend_slides()`, `pptx_gen._add_what_changed_slides()` (via `find_top_movers`)

Sorted by: `current_month` descending.

| Key | Python Type | Description | pptx_gen Consumer | Formatter |
|-----|------------|-------------|-------------------|-----------|
| `app` | `str` | The `awn_app` name | `_add_mom_table_slides`, `_add_forecast_vs_spend_slides` | None (raw string) |
| `category` | `str` | The `awn_category` name (or `"Other"`) | Not directly consumed by pptx_gen | None |
| `current_month` | `float` | Current month COGS spend | `_add_mom_table_slides`, `_add_forecast_vs_spend_slides` | `_fmt_abbreviated` |
| `previous_month` | `float` | Previous month COGS spend | `_add_mom_table_slides` | `_fmt_abbreviated` |
| `two_months_ago` | `float` | Two months ago COGS spend | Not consumed by pptx_gen directly | N/A |
| `forecast` | `float \| None` | Forecast value for this app (None if no forecast) | `_add_forecast_vs_spend_slides`, `_add_what_changed_slides` | `_fmt_abbreviated` |
| `mom_change` | `float` | `current_month - previous_month` | `_add_mom_table_slides` | `_fmt_abbreviated` |
| `mom_pct` | `float` | MoM percentage change | `_add_mom_table_slides` | `_fmt_pct` |
| `prev_vs_prior_change` | `float` | `previous_month - two_months_ago` | Not consumed by pptx_gen | N/A |
| `prev_vs_prior_pct` | `float` | Previous vs two-months-ago percentage | Not consumed by pptx_gen | N/A |
| `var_change` | `float \| None` | `current_month - forecast` (None if no forecast) | `_add_forecast_vs_spend_slides`, `_add_what_changed_slides` | `_fmt_abbreviated` |
| `var_pct` | `float \| None` | Variance percentage (None if no forecast) | `_add_forecast_vs_spend_slides`, `_add_what_changed_slides` | `_fmt_pct` |

### 2b. `category_rollup` (list of dicts)

Produced by: `app_data.build_category_rollup()`
Consumed by: `pptx_gen._add_category_slides()`

Sorted by: `current_month` descending.

| Key | Python Type | Description | pptx_gen Consumer | Formatter |
|-----|------------|-------------|-------------------|-----------|
| `category` | `str` | The `awn_category` name | `_add_category_slides` | None (raw string) |
| `current_month` | `float` | Summed current month spend for category | `_add_category_slides` | `_fmt_abbreviated` |
| `previous_month` | `float` | Summed previous month spend for category | Not displayed in category slide | N/A |
| `forecast` | `float \| None` | Summed forecast (None if no apps have forecast) | `_add_category_slides` | `_fmt_abbreviated` |
| `mom_change` | `float` | `current_month - previous_month` | Not displayed in category slide | N/A |
| `mom_pct` | `float` | MoM percentage | Not displayed in category slide | N/A |
| `var_change` | `float \| None` | `current_month - forecast` (None if no forecast) | `_add_category_slides` | `_fmt_abbreviated` |
| `var_pct` | `float \| None` | Variance percentage (None if no forecast) | `_add_category_slides` | `_fmt_pct` |

### 2c. `top_movers` (dict with two lists)

Produced by: `app_data.find_top_movers()`
Consumed by: `pptx_gen._add_what_changed_slides()`

Top-level structure:

```python
{
    "increases": list[dict],  # Top N apps with largest positive var_change
    "decreases": list[dict],  # Top N apps with largest negative var_change
}
```

Each element in `increases` and `decreases` is an `app_metrics` dict (same shape as section 2a). Only apps where `var_change is not None` (i.e., they have a forecast) are eligible.

- `increases`: Sorted by `var_change` descending (largest overspend first), capped at `n` (default 3).
- `decreases`: Sorted by `var_change` ascending (largest underspend first), capped at `n` (default 3).

Keys consumed by `_add_what_changed_slides`:

| Key | Formatter |
|-----|-----------|
| `app` | None (raw string) |
| `forecast` | `_fmt_abbreviated` |
| `current_month` | `_fmt_abbreviated` |
| `var_change` | `_fmt_abbreviated` |
| `var_pct` | `_fmt_pct` |

### 2d. `app_totals` / `category_totals` (dicts)

Produced by: `app_data.compute_totals(rows, label_key="app")` and `app_data.compute_totals(rows, label_key="category")`
Consumed by: `pptx_gen._add_mom_table_slides()`, `pptx_gen._add_forecast_vs_spend_slides()`, `pptx_gen._add_category_slides()`

Both return a single dict with identical structure. The `label_key` param determines whether the label field is `"app"` or `"category"`.

| Key | Python Type | Description | pptx_gen Consumer | Formatter |
|-----|------------|-------------|-------------------|-----------|
| `{label_key}` | `str` | Always `"Total"` | Row label in total row | None |
| `category` | `str` | Always `"Total"` | N/A (redundant with label) | N/A |
| `current_month` | `float` | Summed current month across all rows | `_add_mom_table_slides`, `_add_forecast_vs_spend_slides`, `_add_category_slides` | `_fmt_abbreviated` |
| `previous_month` | `float` | Summed previous month across all rows | `_add_mom_table_slides` | `_fmt_abbreviated` |
| `two_months_ago` | `float` | Summed two-months-ago across all rows | Not consumed by pptx_gen | N/A |
| `forecast` | `float \| None` | Summed forecast (None if no rows have forecast) | `_add_forecast_vs_spend_slides`, `_add_category_slides` | `_fmt_abbreviated` |
| `mom_change` | `float` | `current_month - previous_month` | `_add_mom_table_slides` | `_fmt_abbreviated` |
| `mom_pct` | `float` | MoM percentage | `_add_mom_table_slides` | `_fmt_pct` |
| `prev_vs_prior_change` | `float` | `previous_month - two_months_ago` | Not consumed by pptx_gen | N/A |
| `prev_vs_prior_pct` | `float` | Previous vs two-months-ago percentage | Not consumed by pptx_gen | N/A |
| `var_change` | `float \| None` | `current_month - forecast` | `_add_forecast_vs_spend_slides`, `_add_category_slides` | `_fmt_abbreviated` |
| `var_pct` | `float \| None` | Variance percentage | `_add_forecast_vs_spend_slides`, `_add_category_slides` | `_fmt_pct` |

### 2e. `drilldown_metrics` (list of dicts)

Produced by: `app_data.build_drilldown_metrics()`
Consumed by: `pptx_gen._add_drilldown_slides()`

Input comes from `ingestion.fetch_cogs_drilldown()` raw rows. `build_drilldown_metrics()` spreads the raw row keys and adds enrichment keys.

Sorted by: `abs(delta_cost)` descending.

| Key | Python Type | Origin | pptx_gen Consumer | Formatter |
|-----|------------|--------|-------------------|-----------|
| `awn_app` | `str` | Raw from Redshift (`COALESCE(awn_app, 'Untagged')`) | Not directly displayed | N/A |
| `product_name` | `str` | Raw from Redshift | Not directly displayed | N/A |
| `operation` | `str` | Raw from Redshift | Not directly displayed | N/A |
| `current_month` | `float` | Raw from Redshift (conditional SUM for current month) | `_add_drilldown_slides` | `_fmt_abbreviated` |
| `previous_month` | `float` | Raw from Redshift (conditional SUM for previous month) | `_add_drilldown_slides` | `_fmt_abbreviated` |
| `delta_cost` | `float` | Raw from Redshift (`current_month - previous_month`) | `_add_drilldown_slides` | `_fmt_abbreviated` |
| `label` | `str` | Enriched: `"awn_app \| product_name \| operation"` (product_name abbreviated via `_PRODUCT_NAME_ABBREVIATIONS`) | `_add_drilldown_slides` | None (raw string) |
| `mom_pct` | `float` | Enriched: `(delta_cost / previous_month) * 100` | `_add_drilldown_slides` | `_fmt_pct` (unless `mom_label` is set) |
| `mom_label` | `str \| None` | Enriched: `"NEW"` if previous=0 and current>0; `"REMOVED"` if current=0 and previous>0; else `None` | `_add_drilldown_slides` | Displayed raw if non-None, else `_fmt_pct(mom_pct)` |

Product name abbreviation map (in `app_data.py`):
- `"Elastic Compute Cloud"` -> `"EC2"`
- `"Simple Storage Service"` -> `"S3"`

### 2f. `ec2_metrics` (list of dicts)

Produced by: `app_data.build_ec2_purchase_metrics()`
Consumed by: `pptx_gen._add_ec2_purchase_slides()`

Input comes from `ingestion.fetch_ec2_purchase_breakdown()` raw rows.

Sorted by: `abs(delta_cost)` descending.

| Key | Python Type | Origin | pptx_gen Consumer | Formatter |
|-----|------------|--------|-------------------|-----------|
| `purchase_option` | `str` | Raw from Redshift | Not directly displayed | N/A |
| `region` | `str` | Raw from Redshift | Not directly displayed | N/A |
| `current_month` | `float` | Raw from Redshift (conditional SUM for current month) | `_add_ec2_purchase_slides` | `_fmt_abbreviated` |
| `previous_month` | `float` | Raw from Redshift (conditional SUM for previous month) | `_add_ec2_purchase_slides` | `_fmt_abbreviated` |
| `delta_cost` | `float` | Raw from Redshift (`current_month - previous_month`) | `_add_ec2_purchase_slides` | `_fmt_abbreviated` |
| `label` | `str` | Enriched: `"purchase_option \| region"` | `_add_ec2_purchase_slides` | None (raw string) |
| `mom_pct` | `float` | Enriched: `(delta_cost / previous_month) * 100` | `_add_ec2_purchase_slides` | `_fmt_pct` (unless `mom_label` is set) |
| `mom_label` | `str \| None` | Enriched: `"NEW"` / `"REMOVED"` / `None` | `_add_ec2_purchase_slides` | Displayed raw if non-None, else `_fmt_pct(mom_pct)` |

### 2g. `drilldown_totals` / `ec2_totals` (dicts)

Produced by: `app_data.compute_drilldown_totals()`
Consumed by: `pptx_gen._add_drilldown_slides()`, `pptx_gen._add_ec2_purchase_slides()`

Both use the same function and return the same shape.

| Key | Python Type | Description | pptx_gen Consumer | Formatter |
|-----|------------|-------------|-------------------|-----------|
| `current_month` | `float` | Sum of `current_month` across all rows | `_add_drilldown_slides`, `_add_ec2_purchase_slides` | `_fmt_abbreviated` |
| `previous_month` | `float` | Sum of `previous_month` across all rows | `_add_drilldown_slides`, `_add_ec2_purchase_slides` | `_fmt_abbreviated` |
| `delta_cost` | `float` | Sum of `delta_cost` across all rows | `_add_drilldown_slides`, `_add_ec2_purchase_slides` | `_fmt_abbreviated` |
| `mom_pct` | `float` | `(total_delta / total_previous) * 100` | `_add_drilldown_slides`, `_add_ec2_purchase_slides` | `_fmt_pct` |

---

## 3. Field-to-Column Mapping Tables

### 3a. AWN COGS by Category (slide function: `_add_category_slides`)

Slide title: **"AWN COGS by Category"**

| Table Column Header | Data Field | Source Dict | Formatter |
|---------------------|-----------|-------------|-----------|
| `"Category"` | `category` | `category_rollup[i]` | None |
| `"Forecast"` | `forecast` | `category_rollup[i]` | `_fmt_abbreviated` |
| `"Spend"` | `current_month` | `category_rollup[i]` | `_fmt_abbreviated` |
| `"Change $"` | `var_change` | `category_rollup[i]` | `_fmt_abbreviated` |
| `"Change %"` | `var_pct` | `category_rollup[i]` | `_fmt_pct` |

Total row uses same fields from `category_totals`.

### 3b. AWN COGS Spend vs Previous Month (slide function: `_add_mom_table_slides`)

Slide title: **"AWN COGS Spend vs Previous Month"**

| Table Column Header | Data Field | Source Dict | Formatter |
|---------------------|-----------|-------------|-----------|
| `"App"` | `app` | `app_metrics[i]` | None |
| `"{prev_label}"` (e.g. "Dec 2025") | `previous_month` | `app_metrics[i]` | `_fmt_abbreviated` |
| `"{current_label}"` (e.g. "Jan 2026") | `current_month` | `app_metrics[i]` | `_fmt_abbreviated` |
| `"Change $"` | `mom_change` | `app_metrics[i]` | `_fmt_abbreviated` |
| `"Change %"` | `mom_pct` | `app_metrics[i]` | `_fmt_pct` |

Total row uses same fields from `app_totals`.

### 3c. AWN COGS Forecast vs Spend (slide function: `_add_forecast_vs_spend_slides`)

Slide title: **"AWN COGS Forecast vs Spend"**

| Table Column Header | Data Field | Source Dict | Formatter |
|---------------------|-----------|-------------|-----------|
| `"App"` | `app` | `app_metrics[i]` | None |
| `"Forecast"` | `forecast` | `app_metrics[i]` | `_fmt_abbreviated` |
| `"Spend"` | `current_month` | `app_metrics[i]` | `_fmt_abbreviated` |
| `"Change $"` | `var_change` | `app_metrics[i]` | `_fmt_abbreviated` |
| `"Change %"` | `var_pct` | `app_metrics[i]` | `_fmt_pct` |

Total row uses same fields from `app_totals`.

### 3d. What Changed? (slide function: `_add_what_changed_slides`)

Slide title: **"What Changed?"**

Two sub-tables on the same slide: "Top Increases vs Forecast" and "Top Decreases vs Forecast".

| Table Column Header | Data Field | Source Dict | Formatter |
|---------------------|-----------|-------------|-----------|
| `"Top Increases vs Forecast"` / `"Top Decreases vs Forecast"` | `app` | `top_movers["increases"][i]` / `top_movers["decreases"][i]` | None |
| `"Forecast"` | `forecast` | Same | `_fmt_abbreviated` |
| `"Spend"` | `current_month` | Same | `_fmt_abbreviated` |
| `"Change $"` | `var_change` | Same | `_fmt_abbreviated` |
| `"Change %"` | `var_pct` | Same | `_fmt_pct` |

No total row on this slide.

### 3e. AWN COGS Drill-Down: Top MoM Movers (slide function: `_add_drilldown_slides`)

Slide title: **"AWN COGS Drill-Down: Top MoM Movers"**

| Table Column Header | Data Field | Source Dict | Formatter |
|---------------------|-----------|-------------|-----------|
| `"App \| Service \| Operation"` | `label` | `drilldown_metrics[i]` | None |
| `"{prev_label}"` (e.g. "Dec 2025") | `previous_month` | `drilldown_metrics[i]` | `_fmt_abbreviated` |
| `"{current_label}"` (e.g. "Jan 2026") | `current_month` | `drilldown_metrics[i]` | `_fmt_abbreviated` |
| `"Change $"` | `delta_cost` | `drilldown_metrics[i]` | `_fmt_abbreviated` |
| `"Change %"` | `mom_label` if non-None, else `mom_pct` | `drilldown_metrics[i]` | Raw string if `mom_label`, else `_fmt_pct` |

Total row uses same fields from `drilldown_totals` (which uses `delta_cost` and `mom_pct`).

Column widths: `[4.5, 2.0, 2.0, 2.0, 1.8]` (inches). Font size: `11` (hardcoded).

### 3f. EC2 RunInstances: Purchase Option Breakdown (slide function: `_add_ec2_purchase_slides`)

Slide title: **"EC2 RunInstances: Purchase Option Breakdown"**

| Table Column Header | Data Field | Source Dict | Formatter |
|---------------------|-----------|-------------|-----------|
| `"Purchase Option \| Region"` | `label` | `ec2_metrics[i]` | None |
| `"{prev_label}"` (e.g. "Dec 2025") | `previous_month` | `ec2_metrics[i]` | `_fmt_abbreviated` |
| `"{current_label}"` (e.g. "Jan 2026") | `current_month` | `ec2_metrics[i]` | `_fmt_abbreviated` |
| `"Change $"` | `delta_cost` | `ec2_metrics[i]` | `_fmt_abbreviated` |
| `"Change %"` | `mom_label` if non-None, else `mom_pct` | `ec2_metrics[i]` | Raw string if `mom_label`, else `_fmt_pct` |

Total row uses same fields from `ec2_totals`.

Column widths: `[4.5, 2.0, 2.0, 2.0, 1.8]` (inches). Font size: `11` (hardcoded).

### 3g. Narrative Content Slides (slide function: `generate_pptx` main loop)

Slide title: Section title from `config["pptx"]["sections"][i]["title"]`

Each content slide renders a narrative string into placeholder idx 11. The narrative is built from `all_metrics` (the bucket-level metrics from `calculate_all_buckets`), which has this shape per bucket:

| Key | Python Type | Description |
|-----|------------|-------------|
| `metric` | `str` | Bucket name (e.g. "Total", "COGS", "OpEx") |
| `actual` | `float` | Current month actual spend |
| `previous_month` | `float` | Previous month actual spend |
| `mom_delta` | `float` | `actual - previous_month` |
| `mom_pct` | `float` | MoM percentage change |
| `mom_dir` | `str` | `"increase"` / `"decrease"` / `"no change"` |
| `has_forecast` | `bool` | Whether a forecast value exists |
| `forecast` | `float` | _(Only present if `has_forecast`)_ Forecast value |
| `var_delta` | `float` | _(Only present if `has_forecast`)_ `actual - forecast` |
| `var_pct` | `float` | _(Only present if `has_forecast`)_ Variance percentage |
| `var_dir` | `str` | _(Only present if `has_forecast`)_ `"over"` / `"under"` / `"on target"` |

The narrative module (`narrative.py`) enriches these with formatted strings before template rendering. Formatters used in narrative are `_format_dollars()` (not the same as `pptx_gen._fmt_abbreviated()` -- slightly different formatting for small values).

### 3h. Trend Chart Slides (slide function: `generate_pptx` main loop + `charts.py`)

Each chart slide embeds a PNG rendered from `chart_data[bucket_name]`, which is a list of:

| Key | Python Type | Description |
|-----|------------|-------------|
| `month_label` | `str` | E.g. `"Jan 2026"` |
| `actual` | `float` | Monthly actual spend |
| `forecast` | `float \| None` | Monthly forecast (None if unavailable) |

The chart module uses its own `_fmt_abbreviated()` function (defined in `charts.py`) for data point labels on the chart. This is functionally identical to `pptx_gen._fmt_abbreviated()`.

---

## 4. Complete `app_data` Dict Structure (as passed to `generate_pptx`)

The `app_data` parameter passed to `pptx_gen.generate_pptx()` has the following top-level shape:

```python
app_data = {
    "app_metrics": list[dict],           # Section 2a
    "category_rollup": list[dict],       # Section 2b
    "top_movers": {                      # Section 2c
        "increases": list[dict],
        "decreases": list[dict],
    },
    "app_totals": dict,                  # Section 2d
    "category_totals": dict,             # Section 2d
    "drilldown": {                       # Phase 3 (optional key)
        "drilldown_metrics": list[dict], # Section 2e
        "drilldown_totals": dict,        # Section 2g
        "ec2_metrics": list[dict],       # Section 2f
        "ec2_totals": dict,              # Section 2g
    },
}
```

The `drilldown` key is optional. When absent, Phase 3 slides are skipped. Within `drilldown`, the `ec2_metrics` key is also optional (checked via `dd.get("ec2_metrics")`).

---

## 5. Redshift Source Table Columns Referenced

All queries target `public.daily_cur_summary`. The following columns are referenced across the codebase:

| Column | Used In | Purpose |
|--------|---------|---------|
| `usage_date` | All queries (date filter) | Timestamp column for month filtering via `DATE_TRUNC` |
| `net_amortized_cost` | Total bucket | Cost column for Total spend |
| `cogs_adjusted_cost` | COGS bucket, drilldown, EC2 breakdown | Cost column for COGS spend |
| `opex_adjusted_cost` | OpEx bucket | Cost column for OpEx spend |
| `awn_app` | `fetch_app_actuals`, `fetch_cogs_drilldown` | Application tag for grouping |
| `product_name` | `fetch_cogs_drilldown`, `fetch_ec2_purchase_breakdown` | AWS service product name |
| `operation` | `fetch_cogs_drilldown`, `fetch_ec2_purchase_breakdown` | AWS API operation |
| `purchase_option` | `fetch_ec2_purchase_breakdown` | EC2 pricing model (On-Demand, Reserved, Savings Plan) |
| `region` | `fetch_ec2_purchase_breakdown` | AWS region |
| `charge_type` | All queries (filter) | Filtered to exclude `'Credit'` |
| `service` | All queries (filter) | Filtered to exclude Databricks and Tackle services |
