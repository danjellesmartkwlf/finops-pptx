"""PowerPoint slide builder for the FinOps Report Generator.

Builds a deck from scratch using the corporate template's slide layouts.
Each run:
  1. Opens the .pptx template to access its slide layouts.
  2. Removes all pre-existing slides.
  3. Adds a Title slide, Transition slides, and Content slides
     based on the sections defined in config.yaml.
"""

from __future__ import annotations

import calendar
import logging
from io import BytesIO
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Template resolution
# ---------------------------------------------------------------------------

_TEMPLATE_CANDIDATES: list[str] = [
    "pptx_template/template1.pptx"
]


def _resolve_template_path(
    config: dict[str, Any],
    project_root: Path | None = None,
) -> Path:
    """Locate the .pptx template file on disk."""
    if project_root is None:
        project_root = Path(__file__).resolve().parent.parent

    for candidate in _TEMPLATE_CANDIDATES:
        candidate_path = project_root / candidate
        if candidate_path.is_file():
            logger.info("Using template at %s", candidate_path)
            return candidate_path

    config_path_str: str = config.get("pptx", {}).get(
        "template_path", "assets/template.pptx"
    )
    config_path = project_root / config_path_str
    if config_path.is_file():
        logger.info("Using config-defined template at %s", config_path)
        return config_path

    searched = [
        str(project_root / c) for c in _TEMPLATE_CANDIDATES
    ] + [str(config_path)]
    raise FileNotFoundError(
        f"PowerPoint template not found. Searched: {searched}"
    )


# ---------------------------------------------------------------------------
# Layout & slide helpers
# ---------------------------------------------------------------------------

def _get_layout_by_name(prs: Any, name: str) -> Any:
    """Return a slide layout by its name, or raise ValueError."""
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    available = [layout.name for layout in prs.slide_layouts]
    raise ValueError(f"Layout '{name}' not found. Available: {available}")


def _remove_all_slides(prs: Any) -> None:
    """Remove every slide from the presentation, keeping layouts intact."""
    sld_id_lst = prs.slides._sldIdLst
    while len(sld_id_lst):
        rId = sld_id_lst[0].rId
        prs.part.drop_rel(rId)
        del sld_id_lst[0]


# ---------------------------------------------------------------------------
# Placeholder helpers
# ---------------------------------------------------------------------------

def _parse_bullet_lines(text: str) -> list[dict[str, Any]]:
    """Parse narrative text that uses bullet markup.

    Convention (driven by config.yaml templates):
        ``- Text``       → level 0, bold (section header)
        ``  - Text``     → level 1, normal (detail item)
        empty line       → blank spacer paragraph

    Plain text without any ``- `` prefixes falls through as level-0,
    non-bold paragraphs so that title/subtitle text is unaffected.
    """
    lines = text.split("\n")
    has_bullets = any(
        line.lstrip().startswith("- ") for line in lines if line.strip()
    )
    if not has_bullets:
        # Plain text path -- no bullet parsing
        return [{"text": line, "level": 0, "bold": False} for line in lines]

    items: list[dict[str, Any]] = []
    for line in lines:
        stripped = line.rstrip()
        if not stripped:
            # Blank line → spacer paragraph
            items.append({"text": "", "level": 0, "bold": False})
            continue

        content = stripped.lstrip()
        indent = len(stripped) - len(content)

        if content.startswith("- "):
            content = content[2:]
            level = 1 if indent >= 2 else 0
            items.append({"text": content, "level": level, "bold": level == 0})
        else:
            items.append({"text": stripped, "level": 0, "bold": False})

    return items


def _set_placeholder_text(
    placeholder: Any,
    text: str,
    font_size_pt: float | None = None,
) -> None:
    """Replace placeholder text while preserving the template's formatting.

    If *text* contains bullet markup (``- `` prefixes), paragraphs are
    assigned the appropriate ``level`` (0 or 1) and section headers are
    bolded.  Plain text is rendered as simple paragraphs.
    """
    tf = placeholder.text_frame

    # Capture existing formatting from the first run (if any)
    existing_font_name: str | None = None
    existing_font_size: Any = None
    existing_font_color: Any = None

    if tf.paragraphs and tf.paragraphs[0].runs:
        first_run = tf.paragraphs[0].runs[0]
        existing_font_name = first_run.font.name
        existing_font_size = first_run.font.size
        try:
            existing_font_color = first_run.font.color.rgb
        except AttributeError:
            existing_font_color = None

    tf.clear()

    items = _parse_bullet_lines(text)
    for i, item in enumerate(items):
        if i == 0:
            paragraph = tf.paragraphs[0]
        else:
            paragraph = tf.add_paragraph()

        paragraph.level = item["level"]

        run = paragraph.add_run()
        run.text = item["text"]

        # Font formatting
        if existing_font_name:
            run.font.name = existing_font_name
        if font_size_pt is not None:
            run.font.size = Pt(font_size_pt)
        elif existing_font_size is not None:
            run.font.size = existing_font_size
        if item["bold"]:
            run.font.color.rgb = RGBColor(0xFF, 0x73, 0x0E)
        elif existing_font_color is not None:
            run.font.color.rgb = existing_font_color

        run.font.bold = item["bold"]


def _set_text_by_idx(slide: Any, idx: int, text: str) -> bool:
    """Find a placeholder by idx on *slide* and set its text.

    Returns True if the placeholder was found and set, False otherwise.
    """
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == idx:
            _set_placeholder_text(shape, text)
            return True
    logger.warning("Placeholder idx %d not found on slide.", idx)
    return False


# ---------------------------------------------------------------------------
# Layout constants (placeholder idx values per layout)
#   Title Slide A:    idx 0 = TITLE, idx 10 = BODY (subtitle)
#   Transition Slide B: idx 0 = TITLE
#   Title, Body:      idx 0 = TITLE, idx 11 = BODY
# ---------------------------------------------------------------------------

_LAYOUT_TITLE_SLIDE = "Title Slide A"
_LAYOUT_TRANSITION = "Transition Slide B"
_LAYOUT_CONTENT = "Title, Body"
_LAYOUT_TITLE_ONLY = "Title Only"


# ---------------------------------------------------------------------------
# Table helpers (Phase 2)
# ---------------------------------------------------------------------------

# Maximum data rows (excluding header/total) before splitting to a new slide
_MAX_TABLE_ROWS_PER_SLIDE = 16

# "Dark Style 1 - Accent 4" built-in table style GUID
_TABLE_STYLE_GUID = "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}"


def _fmt_abbreviated(value: float | None) -> str:
    """Format a dollar value with abbreviated suffix (K, M, B)."""
    if value is None:
        return "N/A"
    abs_val = abs(value)
    sign = "-" if value < 0 else ""
    if abs_val >= 1_000_000_000:
        return f"{sign}${abs_val / 1_000_000_000:.2f}B"
    if abs_val >= 1_000_000:
        return f"{sign}${abs_val / 1_000_000:.2f}M"
    if abs_val >= 1_000:
        return f"{sign}${abs_val / 1_000:.1f}K"
    return f"{sign}${abs_val:,.0f}"


def _fmt_pct(value: float | None) -> str:
    """Format a percentage value with sign."""
    if value is None:
        return "N/A"
    sign = "+" if value > 0 else ""
    return f"{sign}{value:.1f}%"


def _apply_table_style(table: Any) -> None:
    """Apply the 'Dark Style 1 - Accent 4' built-in style via GUID."""
    tbl_pr = table._tbl.tblPr

    # Remove any existing tableStyleId element (python-pptx default)
    for child in list(tbl_pr):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("tableStyleId", "tblStyle"):
            tbl_pr.remove(child)

    style_el = etree.SubElement(tbl_pr, qn("a:tableStyleId"))
    style_el.text = _TABLE_STYLE_GUID


def _set_cell(
    cell: Any,
    text: str,
    font_size: int = 18,
    bold: bool = False,
    alignment: Any = None,
) -> None:
    """Write text into a table cell with formatting."""
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    # Reduce internal margins for tighter cell spacing
    tf.margin_top = Emu(36000)
    tf.margin_bottom = Emu(36000)

    p = tf.paragraphs[0]
    if alignment is not None:
        p.alignment = alignment

    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold


def _font_size_for_row_count(data_row_count: int) -> int:
    """Return the appropriate font size based on data row count.

    - Under 10 data rows: 18pt
    - 10-14 data rows: 16pt
    - 15+ data rows: 14pt
    """
    if data_row_count >= 15:
        return 14
    if data_row_count >= 10:
        return 16
    return 18


def _build_table_on_slide(
    slide: Any,
    headers: list[str],
    data_rows: list[list[str]],
    total_row: list[str] | None = None,
    col_widths: list[float] | None = None,
    font_size: int | None = None,
) -> Any:
    """Create a styled table on a slide.

    Args:
        slide: The slide to add the table to.
        headers: Column header strings.
        data_rows: List of lists, each inner list is one row of cell strings.
        total_row: Optional totals row (same length as headers).
        col_widths: Optional column widths in inches.
        font_size: Optional font size override. If None, auto-computed
            from the data row count.

    Returns:
        The python-pptx Table object.
    """
    n_rows = 1 + len(data_rows) + (1 if total_row else 0)
    n_cols = len(headers)

    if font_size is None:
        font_size = _font_size_for_row_count(len(data_rows))

    left = Inches(0.5)
    top = Inches(1.4)
    width = Inches(12.3)
    height = Inches(min(0.32 * n_rows, 5.8))

    shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = shape.table

    # Set tblPr flags (no banded columns)
    tbl_pr = table._tbl.tblPr
    tbl_pr.set("firstRow", "1")
    tbl_pr.set("lastRow", "1" if total_row else "0")
    tbl_pr.set("bandRow", "1")
    tbl_pr.set("bandCol", "0")
    tbl_pr.set("firstCol", "1")

    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = Inches(w)

    # -- Populate cells --
    for col_idx, header in enumerate(headers):
        align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
        _set_cell(table.cell(0, col_idx), header, font_size=font_size, bold=True, alignment=align)

    for row_idx, row_data in enumerate(data_rows, start=1):
        for col_idx, value in enumerate(row_data):
            align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
            _set_cell(table.cell(row_idx, col_idx), value, font_size=font_size, alignment=align)

    if total_row:
        t_row_idx = 1 + len(data_rows)
        for col_idx, value in enumerate(total_row):
            align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
            _set_cell(
                table.cell(t_row_idx, col_idx), value, font_size=font_size, bold=True, alignment=align
            )

    # -- Apply Dark Style 1 - Accent 4 via built-in GUID --
    _apply_table_style(table)

    return table


def _add_paginated_table_slides(
    prs: Any,
    layout: Any,
    title: str,
    headers: list[str],
    data_rows: list[list[str]],
    total_row: list[str] | None = None,
    col_widths: list[float] | None = None,
    font_size: int | None = None,
) -> None:
    """Add one or more slides with a paginated table.

    If data_rows exceeds ``_MAX_TABLE_ROWS_PER_SLIDE``, splits across
    multiple slides.  The total row only appears on the last slide.

    Args:
        font_size: Optional font size override in points. If None,
            auto-computed from the per-page row count.
    """
    for page_start in range(0, max(len(data_rows), 1), _MAX_TABLE_ROWS_PER_SLIDE):
        page_rows = data_rows[page_start : page_start + _MAX_TABLE_ROWS_PER_SLIDE]
        is_last_page = (
            page_start + _MAX_TABLE_ROWS_PER_SLIDE >= len(data_rows)
        )

        slide = prs.slides.add_slide(layout)
        _set_text_by_idx(slide, 0, title)

        # Use explicit font_size if provided, otherwise auto-compute
        page_font_size = font_size if font_size is not None else _font_size_for_row_count(len(page_rows))

        _build_table_on_slide(
            slide,
            headers,
            page_rows,
            total_row=total_row if is_last_page else None,
            col_widths=col_widths,
            font_size=page_font_size,
        )


# ---------------------------------------------------------------------------
# Phase 2 slide builders
# ---------------------------------------------------------------------------

_TABLE_COL_WIDTHS = [3.0, 2.3, 2.3, 2.3, 2.4]
_TABLE_HEADERS = ["", "Forecast", "Spend", "Change $", "Change %"]


def _add_category_slides(
    prs: Any,
    layout: Any,
    category_rollup: list[dict],
    category_totals: dict,
) -> None:
    """Add the App Categories slide(s)."""
    headers = ["Category"] + _TABLE_HEADERS[1:]

    data_rows = []
    for cat in category_rollup:
        data_rows.append([
            cat["category"],
            _fmt_abbreviated(cat["forecast"]),
            _fmt_abbreviated(cat["current_month"]),
            _fmt_abbreviated(cat["var_change"]),
            _fmt_pct(cat["var_pct"]),
        ])

    total_row = [
        "Total",
        _fmt_abbreviated(category_totals["forecast"]),
        _fmt_abbreviated(category_totals["current_month"]),
        _fmt_abbreviated(category_totals["var_change"]),
        _fmt_pct(category_totals["var_pct"]),
    ]

    _add_paginated_table_slides(
        prs, layout, "AWN COGS by Category",
        headers, data_rows, total_row, _TABLE_COL_WIDTHS,
    )


def _compute_mom_headers(month_label: str) -> tuple[str, str]:
    """Compute column headers for the MoM table.

    Given month_label like "January 2026", returns abbreviated labels for
    the previous month and the current month: ("Dec 2025", "Jan 2026").
    """
    parts = month_label.split()
    month_name, year = parts[0], int(parts[1])
    month_num = list(calendar.month_name).index(month_name)

    # Current month label
    current_label = f"{calendar.month_abbr[month_num]} {year}"

    # Previous month label
    prev_m, prev_y = month_num - 1, year
    if prev_m == 0:
        prev_m, prev_y = 12, year - 1
    prev_label = f"{calendar.month_abbr[prev_m]} {prev_y}"

    return prev_label, current_label


def _add_mom_table_slides(
    prs: Any,
    layout: Any,
    app_metrics: list[dict],
    app_totals: dict,
    month_label: str = "",
) -> None:
    """Add the Spend vs Previous Month slide(s)."""
    prev_label, current_label = _compute_mom_headers(month_label)
    headers = ["App", prev_label, current_label, "Change $", "Change %"]

    data_rows = []
    for m in app_metrics:
        data_rows.append([
            m["app"],
            _fmt_abbreviated(m["previous_month"]),
            _fmt_abbreviated(m["current_month"]),
            _fmt_abbreviated(m["mom_change"]),
            _fmt_pct(m["mom_pct"]),
        ])

    total_row = [
        "Total",
        _fmt_abbreviated(app_totals["previous_month"]),
        _fmt_abbreviated(app_totals["current_month"]),
        _fmt_abbreviated(app_totals["mom_change"]),
        _fmt_pct(app_totals["mom_pct"]),
    ]

    _add_paginated_table_slides(
        prs, layout, "AWN COGS Spend vs Previous Month",
        headers, data_rows, total_row, _TABLE_COL_WIDTHS,
    )


def _add_forecast_vs_spend_slides(
    prs: Any,
    layout: Any,
    app_metrics: list[dict],
    app_totals: dict,
) -> None:
    """Add the Forecast vs Spend slide(s)."""
    headers = ["App"] + _TABLE_HEADERS[1:]

    data_rows = []
    for m in app_metrics:
        data_rows.append([
            m["app"],
            _fmt_abbreviated(m["forecast"]),
            _fmt_abbreviated(m["current_month"]),
            _fmt_abbreviated(m["var_change"]),
            _fmt_pct(m["var_pct"]),
        ])

    total_row = [
        "Total",
        _fmt_abbreviated(app_totals["forecast"]),
        _fmt_abbreviated(app_totals["current_month"]),
        _fmt_abbreviated(app_totals["var_change"]),
        _fmt_pct(app_totals["var_pct"]),
    ]

    _add_paginated_table_slides(
        prs, layout, "AWN COGS Forecast vs Spend",
        headers, data_rows, total_row, _TABLE_COL_WIDTHS,
    )


def _add_what_changed_slides(
    prs: Any,
    layout: Any,
    top_movers: dict[str, list[dict]],
) -> None:
    """Add the 'What Changed?' slide with top increases and decreases."""
    slide = prs.slides.add_slide(layout)
    _set_text_by_idx(slide, 0, "What Changed?")

    headers = ["App", "Forecast", "Spend", "Change $", "Change %"]

    # --- Top increases table (upper half) ---
    increases = top_movers.get("increases", [])
    inc_rows = []
    for m in increases:
        inc_rows.append([
            m["app"],
            _fmt_abbreviated(m["forecast"]),
            _fmt_abbreviated(m["current_month"]),
            _fmt_abbreviated(m["var_change"]),
            _fmt_pct(m["var_pct"]),
        ])

    inc_table_bottom = Inches(1.4)  # track bottom edge for positioning decreases
    if inc_rows:
        n_rows = 1 + len(inc_rows)
        inc_top = Inches(1.4)
        inc_height = Inches(0.40 * n_rows)
        shape_inc = slide.shapes.add_table(
            n_rows, 5, Inches(0.5), inc_top, Inches(12.3), inc_height,
        )
        tbl_inc = shape_inc.table
        tbl_inc._tbl.tblPr.set("firstRow", "1")
        tbl_inc._tbl.tblPr.set("lastRow", "0")
        tbl_inc._tbl.tblPr.set("bandRow", "1")
        tbl_inc._tbl.tblPr.set("bandCol", "0")
        tbl_inc._tbl.tblPr.set("firstCol", "1")
        for i, w in enumerate(_TABLE_COL_WIDTHS):
            tbl_inc.columns[i].width = Inches(w)

        inc_headers = ["Top Increases vs Forecast"] + headers[1:]
        for col_idx, h in enumerate(inc_headers):
            align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
            _set_cell(tbl_inc.cell(0, col_idx), h, bold=True, alignment=align)
        for row_idx, row_data in enumerate(inc_rows, start=1):
            for col_idx, val in enumerate(row_data):
                align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
                _set_cell(tbl_inc.cell(row_idx, col_idx), val, alignment=align)
        _apply_table_style(tbl_inc)
        inc_table_bottom = inc_top + inc_height

    # --- Top decreases table (lower half) ---
    decreases = top_movers.get("decreases", [])
    dec_rows = []
    for m in decreases:
        dec_rows.append([
            m["app"],
            _fmt_abbreviated(m["forecast"]),
            _fmt_abbreviated(m["current_month"]),
            _fmt_abbreviated(m["var_change"]),
            _fmt_pct(m["var_pct"]),
        ])

    if dec_rows:
        dec_top = inc_table_bottom + Inches(0.8)  # buffer gap between tables
        n_rows = 1 + len(dec_rows)
        dec_height = Inches(0.40 * n_rows)
        shape_dec = slide.shapes.add_table(
            n_rows, 5, Inches(0.5), dec_top, Inches(12.3), dec_height,
        )
        tbl_dec = shape_dec.table
        tbl_dec._tbl.tblPr.set("firstRow", "1")
        tbl_dec._tbl.tblPr.set("lastRow", "0")
        tbl_dec._tbl.tblPr.set("bandRow", "1")
        tbl_dec._tbl.tblPr.set("bandCol", "0")
        tbl_dec._tbl.tblPr.set("firstCol", "1")
        for i, w in enumerate(_TABLE_COL_WIDTHS):
            tbl_dec.columns[i].width = Inches(w)

        dec_headers = ["Top Decreases vs Forecast"] + headers[1:]
        for col_idx, h in enumerate(dec_headers):
            align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
            _set_cell(tbl_dec.cell(0, col_idx), h, bold=True, alignment=align)
        for row_idx, row_data in enumerate(dec_rows, start=1):
            for col_idx, val in enumerate(row_data):
                align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
                _set_cell(tbl_dec.cell(row_idx, col_idx), val, alignment=align)
        _apply_table_style(tbl_dec)


# ---------------------------------------------------------------------------
# Phase 3: COGS drill-down slide builders
# ---------------------------------------------------------------------------

_DRILLDOWN_COL_WIDTHS = [4.5, 2.0, 2.0, 2.0, 1.8]


def _add_drilldown_slides(
    prs: Any,
    layout: Any,
    drilldown_metrics: list[dict],
    drilldown_totals: dict,
    month_label: str = "",
) -> None:
    """Add COGS Drill-Down slide(s) showing top MoM movers."""
    prev_label, current_label = _compute_mom_headers(month_label)
    headers = ["App | Service | Operation", prev_label, current_label, "Change $", "Change %"]

    data_rows = []
    for m in drilldown_metrics:
        pct_display = m["mom_label"] if m.get("mom_label") else _fmt_pct(m["mom_pct"])
        data_rows.append([
            m["label"],
            _fmt_abbreviated(m["previous_month"]),
            _fmt_abbreviated(m["current_month"]),
            _fmt_abbreviated(m["delta_cost"]),
            pct_display,
        ])

    total_row = [
        "Total",
        _fmt_abbreviated(drilldown_totals["previous_month"]),
        _fmt_abbreviated(drilldown_totals["current_month"]),
        _fmt_abbreviated(drilldown_totals["delta_cost"]),
        _fmt_pct(drilldown_totals["mom_pct"]),
    ]

    _add_paginated_table_slides(
        prs, layout, "AWN COGS Drill-Down: Top MoM Movers",
        headers, data_rows, total_row, _DRILLDOWN_COL_WIDTHS,
        font_size=11,
    )


def _add_ec2_purchase_slides(
    prs: Any,
    layout: Any,
    ec2_metrics: list[dict],
    ec2_totals: dict,
    month_label: str = "",
) -> None:
    """Add EC2 RunInstances purchase option breakdown slide(s)."""
    prev_label, current_label = _compute_mom_headers(month_label)
    headers = ["Purchase Option | Region", prev_label, current_label, "Change $", "Change %"]

    data_rows = []
    for m in ec2_metrics:
        pct_display = m["mom_label"] if m.get("mom_label") else _fmt_pct(m["mom_pct"])
        data_rows.append([
            m["label"],
            _fmt_abbreviated(m["previous_month"]),
            _fmt_abbreviated(m["current_month"]),
            _fmt_abbreviated(m["delta_cost"]),
            pct_display,
        ])

    total_row = [
        "Total",
        _fmt_abbreviated(ec2_totals["previous_month"]),
        _fmt_abbreviated(ec2_totals["current_month"]),
        _fmt_abbreviated(ec2_totals["delta_cost"]),
        _fmt_pct(ec2_totals["mom_pct"]),
    ]

    _add_paginated_table_slides(
        prs, layout, "EC2 RunInstances: Purchase Option Breakdown",
        headers, data_rows, total_row, _DRILLDOWN_COL_WIDTHS,
        font_size=11,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_pptx(
    narratives: dict[str, str],
    all_metrics: list[dict],
    config: dict[str, Any],
    month_label: str = "",
    project_root: Path | None = None,
    chart_images: dict[str, bytes] | None = None,
    app_data: dict[str, Any] | None = None,
) -> bytes:
    """Build a PowerPoint deck from scratch using template layouts.

    Removes all slides in the template, then constructs:
      1. A title slide (``Title Slide A`` layout).
      2. For each section in ``config["pptx"]["sections"]``:
         a. A transition slide (``Transition Slide B`` layout).
         b. A content slide (``Title, Body`` layout) with the
            bucket's narrative text.
         c. (optional) A chart slide (``Title Only`` layout) with the
            6-month trend chart PNG.
         d. (COGS only, optional) Phase 2 table slides if *app_data*
            is provided.

    Args:
        narratives: Bucket name -> narrative text.
        all_metrics: Metric dicts (reserved for future chart use).
        config: Parsed config.yaml.
        month_label: e.g. ``"January 2026"`` -- used to format the
            title slide heading and subtitle via ``{month}``/``{year}``.
        project_root: Optional project root for template resolution.
        chart_images: Optional dict mapping bucket name to PNG bytes
            for the 6-month trend chart. When provided, a chart slide
            is inserted after each content slide.
        app_data: Optional Phase 2 data dict with keys:
            ``app_metrics``, ``category_rollup``, ``top_movers``,
            ``app_totals``, ``category_totals``.

    Returns:
        The .pptx file as raw bytes.
    """
    template_path = _resolve_template_path(config, project_root=project_root)
    prs = Presentation(str(template_path))

    # Resolve layouts before deleting slides
    title_layout = _get_layout_by_name(prs, _LAYOUT_TITLE_SLIDE)
    transition_layout = _get_layout_by_name(prs, _LAYOUT_TRANSITION)
    content_layout = _get_layout_by_name(prs, _LAYOUT_CONTENT)

    # Title Only layout -- used for chart slides and Phase 2 table slides
    chart_layout = None
    table_layout = None
    try:
        title_only_layout = _get_layout_by_name(prs, _LAYOUT_TITLE_ONLY)
        chart_layout = title_only_layout if chart_images else None
        table_layout = title_only_layout if app_data else None
    except ValueError:
        logger.warning(
            "Layout '%s' not found; chart and table slides will be skipped.",
            _LAYOUT_TITLE_ONLY,
        )
        if chart_images:
            chart_images = None
        if app_data:
            app_data = None

    _remove_all_slides(prs)

    pptx_cfg = config.get("pptx", {})

    # Build format dict from month_label (e.g. "January 2026")
    parts = month_label.split() if month_label else []
    fmt: dict[str, str] = {
        "month": parts[0] if parts else "",
        "year": parts[1] if len(parts) > 1 else "",
    }

    # -- 1. Title slide ----------------------------------------------------
    title_cfg = pptx_cfg.get("title", {})
    title_slide = prs.slides.add_slide(title_layout)
    _set_text_by_idx(title_slide, 0, title_cfg.get("heading", "").format(**fmt))
    _set_text_by_idx(title_slide, 10, title_cfg.get("sub_title", "").format(**fmt))

    # -- 2. Section slides (transition + content + chart) ------------------
    for section in pptx_cfg.get("sections", []):
        section_title = section.get("title", "")
        bucket_name = section.get("bucket", "")

        # Transition slide
        trans_slide = prs.slides.add_slide(transition_layout)
        _set_text_by_idx(trans_slide, 0, section_title)

        # Content slide
        content_slide = prs.slides.add_slide(content_layout)
        _set_text_by_idx(content_slide, 0, section_title)
        _set_text_by_idx(content_slide, 11, narratives.get(bucket_name, ""))

        # Chart slide (if image available)
        if chart_images and chart_layout and bucket_name in chart_images:
            chart_slide = prs.slides.add_slide(chart_layout)
            _set_text_by_idx(chart_slide, 0, section_title)

            png_stream = BytesIO(chart_images[bucket_name])
            chart_slide.shapes.add_picture(
                png_stream,
                left=Inches(0.53),
                top=Inches(1.40),
                width=Inches(12.27),
                height=Inches(5.80),
            )

        # Phase 2: App breakdown tables (after COGS section only)
        if bucket_name == "COGS" and app_data and table_layout:
            # Transition slide for breakdown section
            trans = prs.slides.add_slide(transition_layout)
            _set_text_by_idx(trans, 0, "AWN COGS Breakdown")

            _add_category_slides(
                prs, table_layout,
                app_data["category_rollup"],
                app_data["category_totals"],
            )
            _add_mom_table_slides(
                prs, table_layout,
                app_data["app_metrics"],
                app_data["app_totals"],
                month_label=month_label,
            )
            _add_forecast_vs_spend_slides(
                prs, table_layout,
                app_data["app_metrics"],
                app_data["app_totals"],
            )
            _add_what_changed_slides(
                prs, table_layout,
                app_data["top_movers"],
            )

            # Phase 3: COGS drill-down slides
            if app_data.get("drilldown"):
                dd = app_data["drilldown"]
                _add_drilldown_slides(
                    prs, table_layout,
                    dd["drilldown_metrics"],
                    dd["drilldown_totals"],
                    month_label=month_label,
                )
                if dd.get("ec2_metrics"):
                    _add_ec2_purchase_slides(
                        prs, table_layout,
                        dd["ec2_metrics"],
                        dd["ec2_totals"],
                        month_label=month_label,
                    )

    # -- Serialize ---------------------------------------------------------
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.read()


# ---------------------------------------------------------------------------
# Debug helper
# ---------------------------------------------------------------------------

def inspect_template(
    template_path: str,
) -> dict[str, Any]:
    """Enumerate all slides, shapes, and placeholders in a template."""
    path = Path(template_path)
    if not path.is_file():
        raise FileNotFoundError(f"Template file not found: {template_path}")

    prs = Presentation(str(path))
    result: dict[str, Any] = {
        "slide_count": len(prs.slides),
        "slide_width": str(prs.slide_width),
        "slide_height": str(prs.slide_height),
        "slides": [],
    }

    for slide_idx, slide in enumerate(prs.slides):
        slide_info: dict[str, Any] = {
            "slide_index": slide_idx,
            "layout_name": slide.slide_layout.name,
            "shapes": [],
        }

        for shape in slide.shapes:
            shape_info: dict[str, Any] = {
                "shape_id": shape.shape_id,
                "name": shape.name,
                "shape_type": str(shape.shape_type),
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
                "has_text_frame": shape.has_text_frame,
            }

            if shape.is_placeholder:
                ph_fmt = shape.placeholder_format
                shape_info["is_placeholder"] = True
                shape_info["placeholder_idx"] = ph_fmt.idx
                shape_info["placeholder_type"] = str(ph_fmt.type)
            else:
                shape_info["is_placeholder"] = False

            if shape.has_text_frame:
                shape_info["text_preview"] = shape.text_frame.text[:200]

            slide_info["shapes"].append(shape_info)

        result["slides"].append(slide_info)

    return result
