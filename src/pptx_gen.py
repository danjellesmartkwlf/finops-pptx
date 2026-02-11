"""Config-driven PowerPoint slide builder for the FinOps Report Generator.

Reads slide definitions from ``slides_config.yaml`` and dispatches each
entry to a typed handler.  All formatting and layout helpers live in
``pptx_utils.py``.
"""

from __future__ import annotations

import logging
from io import BytesIO
from pathlib import Path
from typing import Any

import yaml
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from src.pptx_utils import (
    add_paginated_table_slides,
    apply_table_style,
    compute_mom_headers,
    fmt_abbreviated,
    fmt_pct,
    get_layout_by_name,
    remove_all_slides,
    resolve_template_path,
    set_cell,
    set_text_by_idx,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Slide config loading
# ---------------------------------------------------------------------------

_SLIDES_CONFIG: dict[str, Any] | None = None


def _load_slides_config(project_root: Path | None = None) -> dict[str, Any]:
    """Load and cache slides_config.yaml."""
    global _SLIDES_CONFIG
    if _SLIDES_CONFIG is not None:
        return _SLIDES_CONFIG

    if project_root is None:
        project_root = Path(__file__).resolve().parent.parent
    config_path = project_root / "slides_config.yaml"
    with open(config_path, "r") as f:
        _SLIDES_CONFIG = yaml.safe_load(f)
    return _SLIDES_CONFIG


# ---------------------------------------------------------------------------
# Formatter registry
# ---------------------------------------------------------------------------

_FORMATTER_MAP: dict[str, Any] = {
    "abbreviated_currency": fmt_abbreviated,
    "signed_percentage": fmt_pct,
    "raw": None,
}


def _format_value(value: Any, format_name: str | None) -> str:
    """Apply a named formatter to a value."""
    if format_name is None or format_name == "raw":
        return str(value) if value is not None else ""
    func = _FORMATTER_MAP.get(format_name)
    if func is None:
        return str(value) if value is not None else ""
    return func(value)


# ---------------------------------------------------------------------------
# Data resolution helpers
# ---------------------------------------------------------------------------

def _resolve_data_path(data_source: str, context: dict[str, Any]) -> Any:
    """Resolve a dotted path like 'app_data.category_rollup' from context."""
    parts = data_source.split(".")
    obj = context
    for part in parts:
        if isinstance(obj, dict):
            obj = obj.get(part)
        else:
            return None
        if obj is None:
            return None
    return obj


def _check_requires(requires: str | None, context: dict[str, Any]) -> bool:
    """Return True if the ``requires`` dotted path is truthy (or absent)."""
    if requires is None:
        return True
    return bool(_resolve_data_path(requires, context))


def _resolve_headers(
    headers: list[str],
    prev_month: str,
    curr_month: str,
) -> list[str]:
    """Replace ``{prev_month}`` and ``{curr_month}`` tokens in headers."""
    return [
        h.replace("{prev_month}", prev_month).replace("{curr_month}", curr_month)
        for h in headers
    ]


# ---------------------------------------------------------------------------
# Row builders
# ---------------------------------------------------------------------------

def _build_data_rows(
    data: list[dict],
    columns: list[dict],
) -> list[list[str]]:
    """Build formatted table rows from data dicts and column config."""
    rows = []
    for record in data:
        row = []
        for col in columns:
            field = col["field"]
            fmt_name = col.get("format")
            fallback_field = col.get("fallback_field")

            value = record.get(field)

            # fallback_field convention: if the record has a non-None
            # fallback value, use that raw string instead of formatting.
            if fallback_field:
                fallback_val = record.get(fallback_field)
                if fallback_val:
                    row.append(fallback_val)
                    continue

            row.append(_format_value(value, fmt_name))
        rows.append(row)
    return rows


def _build_total_row(
    totals: dict,
    totals_columns: list[dict],
) -> list[str]:
    """Build a formatted total row from totals dict and column config."""
    row = []
    for col in totals_columns:
        if col.get("field") is None:
            row.append(col.get("text", "Total"))
        else:
            row.append(_format_value(totals.get(col["field"]), col.get("format")))
    return row


# ---------------------------------------------------------------------------
# Slide handlers
# ---------------------------------------------------------------------------

def _handle_title(
    prs: Any,
    slide_cfg: dict[str, Any],
    layouts: dict[str, Any],
    context: dict[str, Any],
) -> None:
    """Render the title slide."""
    layout = layouts.get(slide_cfg["layout"])
    if layout is None:
        return
    slide = prs.slides.add_slide(layout)
    fmt = context["fmt"]
    set_text_by_idx(slide, 0, slide_cfg.get("heading", "").format(**fmt))
    set_text_by_idx(slide, 10, slide_cfg.get("subtitle", "").format(**fmt))


def _handle_transition(
    prs: Any,
    slide_cfg: dict[str, Any],
    layouts: dict[str, Any],
    context: dict[str, Any],
) -> None:
    """Render a transition (section divider) slide."""
    if not _check_requires(slide_cfg.get("requires"), context):
        return
    layout = layouts.get(slide_cfg["layout"])
    if layout is None:
        return
    slide = prs.slides.add_slide(layout)
    set_text_by_idx(slide, 0, slide_cfg.get("title", ""))


def _handle_content(
    prs: Any,
    slide_cfg: dict[str, Any],
    layouts: dict[str, Any],
    context: dict[str, Any],
) -> None:
    """Render a content / narrative slide."""
    layout = layouts.get(slide_cfg["layout"])
    if layout is None:
        return
    slide = prs.slides.add_slide(layout)
    set_text_by_idx(slide, 0, slide_cfg.get("title", ""))

    bucket = slide_cfg.get("bucket", "")
    narratives = context.get("narratives", {})
    set_text_by_idx(slide, 11, narratives.get(bucket, ""))


def _handle_chart(
    prs: Any,
    slide_cfg: dict[str, Any],
    layouts: dict[str, Any],
    context: dict[str, Any],
) -> None:
    """Render a chart image slide.  Skipped when no chart for the bucket."""
    chart_images = context.get("chart_images")
    bucket = slide_cfg.get("bucket", "")
    if not chart_images or bucket not in chart_images:
        return

    layout = layouts.get(slide_cfg["layout"])
    if layout is None:
        return
    slide = prs.slides.add_slide(layout)
    set_text_by_idx(slide, 0, slide_cfg.get("title", ""))

    png_stream = BytesIO(chart_images[bucket])
    slide.shapes.add_picture(
        png_stream,
        left=Inches(0.53),
        top=Inches(1.40),
        width=Inches(12.27),
        height=Inches(5.80),
    )


def _handle_table(
    prs: Any,
    slide_cfg: dict[str, Any],
    layouts: dict[str, Any],
    context: dict[str, Any],
) -> None:
    """Render a paginated table slide from YAML config."""
    if not _check_requires(slide_cfg.get("requires"), context):
        return

    layout = layouts.get(slide_cfg["layout"])
    if layout is None:
        return

    data = _resolve_data_path(slide_cfg["data_source"], context)
    if data is None:
        return

    totals_source = slide_cfg.get("totals_source")
    totals = _resolve_data_path(totals_source, context) if totals_source else None

    headers = _resolve_headers(
        slide_cfg["headers"],
        context.get("prev_month", ""),
        context.get("curr_month", ""),
    )

    data_rows = _build_data_rows(data, slide_cfg["columns"])

    total_row = None
    if totals and "totals_columns" in slide_cfg:
        total_row = _build_total_row(totals, slide_cfg["totals_columns"])

    add_paginated_table_slides(
        prs,
        layout,
        slide_cfg["title"],
        headers,
        data_rows,
        total_row=total_row,
        col_widths=slide_cfg.get("col_widths"),
        font_size=slide_cfg.get("font_size"),  # None → auto-compute
    )


def _handle_split_table(
    prs: Any,
    slide_cfg: dict[str, Any],
    layouts: dict[str, Any],
    context: dict[str, Any],
) -> None:
    """Render two tables stacked vertically on one slide.

    The lower table's vertical position is computed dynamically from the
    upper table's rendered height plus ``gap_below``.
    """
    if not _check_requires(slide_cfg.get("requires"), context):
        return

    layout = layouts.get(slide_cfg["layout"])
    if layout is None:
        return
    slide = prs.slides.add_slide(layout)
    set_text_by_idx(slide, 0, slide_cfg.get("title", ""))

    tables_cfg = slide_cfg["tables"]
    upper_bottom = Inches(1.4)  # tracks bottom edge for lower table placement

    for tbl_cfg in tables_cfg:
        data = _resolve_data_path(tbl_cfg["data_source"], context)
        if data is None:
            data = []

        data_rows = _build_data_rows(data, tbl_cfg["columns"])
        if not data_rows:
            continue

        headers = tbl_cfg["headers"]
        col_widths = tbl_cfg.get("col_widths", [3.0, 2.3, 2.3, 2.3, 2.4])
        font_size = tbl_cfg.get("font_size", 18)
        row_height_factor = tbl_cfg.get("row_height_factor", 0.40)
        n_rows = 1 + len(data_rows)
        n_cols = len(headers)

        if tbl_cfg["position"] == "upper":
            tbl_top = Inches(tbl_cfg.get("top", 1.4))
        else:
            gap = tables_cfg[0].get("gap_below", 0.8)
            tbl_top = upper_bottom + Inches(gap)

        tbl_height = Inches(row_height_factor * n_rows)

        shape = slide.shapes.add_table(
            n_rows, n_cols, Inches(0.5), tbl_top, Inches(12.3), tbl_height,
        )
        table = shape.table
        table._tbl.tblPr.set("firstRow", "1")
        table._tbl.tblPr.set("lastRow", "0")
        table._tbl.tblPr.set("bandRow", "1")
        table._tbl.tblPr.set("bandCol", "0")
        table._tbl.tblPr.set("firstCol", "1")
        for i, w in enumerate(col_widths):
            table.columns[i].width = Inches(w)

        for col_idx, h in enumerate(headers):
            align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
            set_cell(
                table.cell(0, col_idx), h,
                font_size=font_size, bold=True, alignment=align,
            )

        for row_idx, row_data in enumerate(data_rows, start=1):
            for col_idx, val in enumerate(row_data):
                align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
                set_cell(
                    table.cell(row_idx, col_idx), val,
                    font_size=font_size, alignment=align,
                )

        apply_table_style(table)

        if tbl_cfg["position"] == "upper":
            upper_bottom = tbl_top + tbl_height


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------

_HANDLER_MAP: dict[str, Any] = {
    "title": _handle_title,
    "transition": _handle_transition,
    "content": _handle_content,
    "chart": _handle_chart,
    "table": _handle_table,
    "split_table": _handle_split_table,
}


def _dispatch_slide(
    prs: Any,
    slide_cfg: dict[str, Any],
    layouts: dict[str, Any],
    context: dict[str, Any],
) -> None:
    """Dispatch a single slide config entry to its handler."""
    slide_type = slide_cfg.get("type", "")
    handler = _HANDLER_MAP.get(slide_type)
    if handler is None:
        logger.warning("Unknown slide type '%s' -- skipping.", slide_type)
        return
    handler(prs, slide_cfg, layouts, context)


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
    """Build a PowerPoint deck from ``slides_config.yaml`` definitions.

    Args:
        narratives: Bucket name -> narrative text.
        all_metrics: Metric dicts (reserved for future use).
        config: Parsed config.yaml (used for template resolution).
        month_label: e.g. ``"January 2026"`` for title formatting.
        project_root: Optional project root for template resolution.
        chart_images: Optional dict mapping bucket name to PNG bytes.
        app_data: Optional Phase 2/3 data dict.

    Returns:
        The ``.pptx`` file as raw bytes.
    """
    slides_config = _load_slides_config(project_root=project_root)
    template_path = resolve_template_path(config, project_root=project_root)
    prs = Presentation(str(template_path))

    # Resolve layout names → layout objects
    layout_map: dict[str, Any] = {}
    for key, name in slides_config.get("layouts", {}).items():
        try:
            layout_map[key] = get_layout_by_name(prs, name)
        except ValueError:
            logger.warning("Layout '%s' (%s) not found.", name, key)
            layout_map[key] = None

    remove_all_slides(prs)

    # Build format dict from month_label (e.g. "January 2026")
    parts = month_label.split() if month_label else []
    fmt = {
        "month": parts[0] if parts else "",
        "year": parts[1] if len(parts) > 1 else "",
    }

    # Compute MoM header labels for template token resolution
    prev_month = ""
    curr_month = ""
    if month_label:
        prev_month, curr_month = compute_mom_headers(month_label)

    # Rendering context available to all handlers
    context: dict[str, Any] = {
        "config": config,
        "narratives": narratives,
        "all_metrics": all_metrics,
        "month_label": month_label,
        "chart_images": chart_images,
        "app_data": app_data,
        "fmt": fmt,
        "prev_month": prev_month,
        "curr_month": curr_month,
    }

    # Walk the slide sequence
    for slide_cfg in slides_config.get("slides", []):
        _dispatch_slide(prs, slide_cfg, layout_map, context)

    # Serialize
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.read()


# ---------------------------------------------------------------------------
# Debug helper
# ---------------------------------------------------------------------------

def inspect_template(template_path: str) -> dict[str, Any]:
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
