# pptx_gen.py Refactor: Risk Analysis

**Author:** Devil's Advocate / Senior Architect Review
**Scope:** Refactoring `src/pptx_gen.py` (936 lines) into a YAML config + `pptx_gen_v2.py` renderer

---

## 1. Edge Cases That Could Break Fidelity

### 1.1 "What Changed?" Slide: Two Tables on One Slide

**Risk: HIGH**

The `_add_what_changed_slides()` function (lines 546-633) is fundamentally different from every other slide builder in the file. It places **two independent tables on a single slide** with a dynamically-computed vertical gap:

```python
inc_table_bottom = inc_top + inc_height   # bottom edge of first table
dec_top = inc_table_bottom + Inches(0.8)  # second table starts 0.8" below
```

The vertical position of the second table depends on the *rendered height* of the first table, which itself depends on the number of rows in `top_movers["increases"]`. This is a runtime-computed layout, not a static template.

**Why this is hard to express in YAML:**
- Every other slide type uses `_build_table_on_slide` or `_add_paginated_table_slides`, which assume a single table at a fixed `top=Inches(1.4)` position.
- A generic "table slide" YAML schema cannot represent "place table B at (bottom-of-table-A + 0.8 inches)" without introducing a mini layout engine in the config.
- The row heights also differ: `0.40 * n_rows` for "What Changed?" vs `0.32 * n_rows` for standard tables.

**Recommendation:** Do NOT try to make this slide config-driven. Keep `_add_what_changed_slides` as a named **special-case handler** that the renderer dispatches to by a `handler: what_changed` key in the YAML. The YAML should declare *that* this slide exists and what data feeds it, but the *layout logic* stays in Python. Attempting to express dual-table dynamic positioning in YAML will produce something harder to read and debug than the Python it replaces.

### 1.2 Font Size Auto-Computation Based on Row Count

**Risk: MEDIUM**

`_font_size_for_row_count()` (lines 288-299) implements a three-tier step function:

| Data Rows | Font Size |
|-----------|-----------|
| < 10      | 18pt      |
| 10-14     | 16pt      |
| 15+       | 14pt      |

Some callers override this (e.g., drill-down slides hardcode `font_size=11`).

**Should this be config or code?**

It should be **code with config overrides**. The step-function thresholds could live in YAML as optional parameters per slide type:

```yaml
font_size_rules:
  default: {under_10: 18, under_15: 16, else: 14}
  drilldown: {fixed: 11}
```

But the *computation* (selecting which tier applies) must remain in Python. Putting the `if/elif` logic in YAML would require a rules engine, which is over-engineering.

### 1.3 Bullet Parsing in `_parse_bullet_lines` / `_set_placeholder_text`

**Risk: HIGH if moved to config; LOW if kept in code**

The bullet parser (lines 93-130) implements a mini-markup language:
- `- Text` at indent 0 = bold, orange-colored section header
- `  - Text` at indent 2+ = normal sub-item
- Empty lines = spacer paragraphs
- Plain text without `- ` prefixes = passthrough

This also applies formatting: orange color (`#FF730E`) for bold headers, preserved template font for everything else.

**Config cannot express this.** This is a text *renderer*, not a data mapping. The YAML should not attempt to describe "if line starts with dash-space at zero indent, make it bold and orange." This belongs entirely in code. The config's role is limited to specifying the narrative template strings (which already live in `config.yaml` under `narrative_templates`).

### 1.4 Dynamic Month-Based Column Headers

**Risk: MEDIUM**

`_compute_mom_headers()` (lines 456-475) takes `"January 2026"` and produces `("Dec 2025", "Jan 2026")`. These are used as table column headers in the MoM and drill-down slides.

The YAML will naturally want to declare column headers, but these headers contain runtime values. A config like:

```yaml
headers: ["App", "{prev_month}", "{curr_month}", "Change $", "Change %"]
```

...requires a template resolution pass. This is feasible but introduces a new concern: **the YAML is no longer pure data; it contains template expressions.** This means:
- Validation becomes harder (you cannot check header count without resolving templates first).
- The renderer must know which placeholders are valid in which context.

**Recommendation:** Allow `{prev_month}` and `{curr_month}` as the only supported template tokens in header definitions. Document them explicitly. Do not build a general-purpose template engine for YAML values.

### 1.5 The `mom_label` Override (Conditional Cell Formatting)

**Risk: MEDIUM**

In `_add_drilldown_slides` (line 656) and `_add_ec2_purchase_slides` (line 693):

```python
pct_display = m["mom_label"] if m.get("mom_label") else _fmt_pct(m["mom_pct"])
```

When an app is NEW or REMOVED, the percentage column shows the string `"NEW"` or `"REMOVED"` instead of a computed percentage. This is a **per-cell conditional override** that comes from the data, not the config.

This is actually fine for a config-driven approach -- it does not need to be in YAML at all. The data-preparation layer (`build_drilldown_metrics`) already sets `mom_label`, and the renderer just needs to check for it. The risk is only if someone tries to move this conditional into YAML rules, which would be unnecessary.

**Recommendation:** Keep this as a renderer convention: "if the data dict contains `mom_label`, use it instead of formatting `mom_pct`." Document it as a data contract, not a config concern.

---

## 2. What Should NOT Be in Config (Keep in Code)

### 2.1 `_fmt_abbreviated` (Dollar Formatting with K/M/B)

**Why code:** This is a pure formatting function with business-specific thresholds ($1K, $1M, $1B) and precise format strings (`$1.23M`, `$45.6K`). Putting thresholds in YAML adds no value -- nobody will change `1_000_000` to a different value. The format strings (`{sign}${abs_val / 1_000_000:.2f}M`) are not expressible in YAML without eval(). This is a utility function, not a configuration knob.

### 2.2 `_apply_table_style` (XML Manipulation)

**Why code:** This function manipulates `lxml` elements directly -- removing children by tag name, creating `SubElement` with namespace-qualified names. YAML cannot represent XML tree surgery. The style GUID (`{E929F9F4-...}`) *could* be a config value, but extracting just the GUID while leaving the XML plumbing in code is not worth the indirection for a single constant.

### 2.3 `_set_placeholder_text` / `_parse_bullet_lines` (Text Rendering)

**Why code:** As discussed in 1.3, this is a rendering engine. It captures existing font properties, clears the text frame, iterates parsed items, sets paragraph levels, applies conditional bold/color. This is 50+ lines of imperative UI code. No config format can express "copy the first run's font name, clear the frame, then for each parsed item create a paragraph at the appropriate level and apply orange color if bold."

### 2.4 Template Resolution Logic (`_resolve_template_path`)

**Why code:** This searches multiple candidate paths with fallback logic. The *paths* are already in config (`_TEMPLATE_CANDIDATES`, `config["pptx"]["template_path"]`). The *resolution algorithm* (try candidates in order, fall back to config path, raise with searched locations) is control flow, not data.

### 2.5 Slide Removal (`_remove_all_slides`)

**Why code:** This manipulates the internal `_sldIdLst` XML element and drops OPC relationships. It is a one-liner in terms of intent ("clear the deck") but requires python-pptx internals knowledge. There is no reason this would ever vary by config.

### 2.6 `_build_table_on_slide` (Table Construction Orchestrator)

**Why partially code:** The *dimensions* (left, top, width, height formula) could be config, but the *orchestration* -- creating the shape, iterating headers/data/totals, applying styles, setting tblPr flags -- is procedural code. Splitting "create table, populate cells, apply style" into config steps would produce a YAML-flavored imperative script, which is worse than Python.

**General principle:** If a function's logic involves branching, iteration over internal python-pptx objects, or XML manipulation, it belongs in code. Config should express **what** (which slides, what data, what titles) but not **how** (cell-by-cell rendering mechanics).

---

## 3. Import Strategy Recommendation

### Options Evaluated

| Strategy | Pros | Cons |
|----------|------|------|
| **(a) Import from pptx_gen.py** | DRY, no duplication | Creates runtime dependency on "old" module; confusing ownership; both files import from each other's domain |
| **(b) Copy helpers into pptx_gen_v2.py** | Standalone, no coupling | Duplication; divergence risk; two copies of `_fmt_abbreviated` etc. |
| **(c) Extract into pptx_utils.py** | Clean separation; both old and new can import; clear ownership | One more file; requires touching existing imports in pptx_gen.py |

### Recommendation: Option (c) -- Extract `pptx_utils.py`

**Rationale:**

1. **Option (a) is the worst choice.** The whole point of the refactor is to replace pptx_gen.py. Creating a dependency from v2 back to the old file means you can never delete the old file without another migration. It also creates confusing import chains during the transition period.

2. **Option (b) is acceptable short-term** but becomes a maintenance hazard. If a bug is found in `_fmt_abbreviated`, it must be fixed in two places. During the transition period where both generators coexist (for fidelity comparison), this is especially dangerous.

3. **Option (c) is the right call.** Extract these functions into `src/pptx_utils.py`:
   - `_fmt_abbreviated` -> `fmt_abbreviated`
   - `_fmt_pct` -> `fmt_pct`
   - `_apply_table_style` -> `apply_table_style`
   - `_set_cell` -> `set_cell`
   - `_font_size_for_row_count` -> `font_size_for_row_count`
   - `_set_placeholder_text` / `_parse_bullet_lines` -> `set_placeholder_text` / `parse_bullet_lines`
   - `_build_table_on_slide` -> `build_table_on_slide`
   - `_get_layout_by_name` -> `get_layout_by_name`
   - `_remove_all_slides` -> `remove_all_slides`
   - `_resolve_template_path` -> `resolve_template_path`
   - `_compute_mom_headers` -> `compute_mom_headers`

   Then update `pptx_gen.py` to import from `pptx_utils.py` (a single find-and-replace of the function calls). The new `pptx_gen_v2.py` also imports from `pptx_utils.py`. Both generators share the exact same helper implementations.

   **Migration path:** After fidelity is confirmed and the old generator is removed, `pptx_gen_v2.py` simply renames to `pptx_gen.py`, and `pptx_utils.py` continues as-is.

   The cost (one additional file) is negligible compared to the benefit of guaranteed helper consistency during the transition period.

---

## 4. YAML Complexity Risk

### 4.1 Can the YAML Become Harder to Understand Than the Python?

**Yes, absolutely.** This is the single biggest risk of the refactor.

The current `pptx_gen.py` is 936 lines, but it is *straightforward* Python. A developer can read `_add_mom_table_slides` and understand exactly what it does in 30 seconds. The function names are descriptive, the flow is linear, and there is no indirection.

A poorly designed YAML config could look like this (anti-pattern):

```yaml
slides:
  - type: paginated_table
    title: "AWN COGS Spend vs Previous Month"
    data_source: app_metrics
    headers:
      - {label: "App", field: "app", formatter: null, align: left}
      - {label: "{prev_month}", field: "previous_month", formatter: abbreviated, align: right}
      - {label: "{curr_month}", field: "current_month", formatter: abbreviated, align: right}
      - {label: "Change $", field: "mom_change", formatter: abbreviated, align: right}
      - {label: "Change %", field: "mom_pct", formatter: pct, align: right}
    totals_source: app_totals
    totals_mapping:
      - {field: null, value: "Total"}
      - {field: "previous_month", formatter: abbreviated}
      ...
    col_widths: [3.0, 2.3, 2.3, 2.3, 2.4]
    pagination: {max_rows: 16}
    font_size: auto
```

This is **harder** to read than the Python. It is longer, requires understanding the YAML schema, and any change requires cross-referencing the renderer to know which keys are valid.

### 4.2 Right Abstraction Level

**The sweet spot is: YAML declares the slide deck *outline*, not the cell-level rendering.**

Good abstraction level (RECOMMENDED):

```yaml
slides:
  - type: title
    heading: "Arctic Wolf AWS Cloud Cost Report"
    sub_title: "FinOps Team | {month} {year}"

  - type: transition
    title: "Arctic Wolf AWS Total Spend"

  - type: narrative
    title: "Arctic Wolf AWS Total Spend"
    bucket: "Total"

  - type: trend_chart
    bucket: "Total"

  - type: transition
    title: "Arctic Wolf AWS COGS"

  - type: narrative
    title: "Arctic Wolf AWS COGS"
    bucket: "COGS"

  - type: trend_chart
    bucket: "COGS"

  - type: transition
    title: "AWN COGS Breakdown"

  - type: category_table
    title: "AWN COGS by Category"

  - type: mom_table
    title: "AWN COGS Spend vs Previous Month"

  - type: forecast_vs_spend_table
    title: "AWN COGS Forecast vs Spend"

  - type: what_changed

  - type: drilldown_table
    title: "AWN COGS Drill-Down: Top MoM Movers"

  - type: ec2_purchase_table
    title: "EC2 RunInstances: Purchase Option Breakdown"
```

Each `type` maps to a Python handler. The YAML says "put a category table here with this title." The Python handler knows how to build that table. Column definitions, formatting, and layout are **not** in the YAML.

Bad abstraction level (AVOID): Trying to make every column, formatter, alignment, and row mapping configurable. This recreates python-pptx's API in YAML syntax.

### 4.3 Guard Rails

1. **Schema validation.** Define a JSON Schema (or Pydantic model) for the YAML. Reject unknown keys at load time. This prevents the YAML from silently drifting.

2. **Max nesting depth: 2 levels.** If any YAML structure requires 3+ levels of nesting, it is a sign the abstraction is leaking implementation details into config. Refactor the Python handler instead.

3. **No conditionals in YAML.** If you find yourself wanting `if:` or `when:` keys in the YAML, that logic belongs in Python. The YAML should be a flat declaration of intent.

4. **One source of truth per concern.** Column widths live in exactly one place (YAML or Python, not both). If column widths are in YAML, they must be in YAML for ALL table types. Inconsistency breeds confusion.

5. **Comments in the YAML.** Each slide entry should have a one-line comment explaining what it produces. The YAML file IS the documentation for "what is in the deck."

---

## 5. Migration Path

### 5.1 How Should `generate_report.py` and `app.py` Switch?

Both files currently import and call the same function:

```python
from src.pptx_gen import generate_pptx
```

(`generate_report.py` line 39, `app.py` imports it indirectly via session state / callback)

**Phase 1 (Development):** Keep both generators. Add a feature flag:

```python
# generate_report.py
import os
if os.environ.get("USE_PPTX_V2", "0") == "1":
    from src.pptx_gen_v2 import generate_pptx
else:
    from src.pptx_gen import generate_pptx
```

This allows A/B comparison during development. The env var is only for internal testing.

**Phase 2 (Fidelity Verified):** After binary-identical output is confirmed (Task #5), change the import to point to v2 and delete the env var branching:

```python
from src.pptx_gen_v2 import generate_pptx
```

**Phase 3 (Cleanup):** Rename `pptx_gen_v2.py` -> `pptx_gen.py`, delete the old file, update all imports. This is a single commit.

### 5.2 Should `pptx_gen_v2.py` Expose the Same Signature?

**Yes, absolutely.** The public API must be identical:

```python
def generate_pptx(
    narratives: dict[str, str],
    all_metrics: list[dict],
    config: dict[str, Any],
    month_label: str = "",
    project_root: Path | None = None,
    chart_images: dict[str, bytes] | None = None,
    app_data: dict[str, Any] | None = None,
) -> bytes:
```

Same function name, same parameter names, same types, same return type. This is non-negotiable for a drop-in replacement. Both `generate_report.py` and `app.py` should require **zero changes** to their calling code when switching generators.

The v2 implementation will *additionally* read the slide config YAML internally, but this is an implementation detail. The caller should not need to know or care whether the generator is config-driven or hardcoded.

**One exception to watch:** If the v2 renderer needs additional data that the current callers do not provide (e.g., a new data key in `app_data`), that is a sign the refactor is expanding scope. The refactor should reproduce existing behavior, not add features. Any new data requirements should be deferred to a separate PR.

---

## Summary of Key Risks

| Risk | Severity | Mitigation |
|------|----------|------------|
| "What Changed?" dual-table layout breaks in generic renderer | HIGH | Use named handler dispatch, not generic table config |
| YAML becomes more complex than the Python it replaces | HIGH | Keep YAML at outline level; column details stay in Python |
| Helper function divergence during transition | MEDIUM | Extract to `pptx_utils.py` (Option c) |
| Template tokens in YAML headers create validation blind spots | MEDIUM | Restrict to `{prev_month}` and `{curr_month}` only |
| Bullet-parsing logic accidentally moved to config | MEDIUM | Document that text rendering is always code |
| Callers require changes to switch generators | LOW | Match `generate_pptx()` signature exactly |
| Font size logic split across config and code | LOW | Config provides overrides; code owns the step function |
