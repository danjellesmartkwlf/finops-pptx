"""PowerPoint slide builder for the FinOps Report Generator.

Builds a deck from scratch using the corporate template's slide layouts.
Each run:
  1. Opens the .pptx template to access its slide layouts.
  2. Removes all pre-existing slides.
  3. Adds a Title slide, Transition slides, and Content slides
     based on the sections defined in config.yaml.
"""

from __future__ import annotations

import logging
from io import BytesIO
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

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
# Public API
# ---------------------------------------------------------------------------

def generate_pptx(
    narratives: dict[str, str],
    all_metrics: list[dict],
    config: dict[str, Any],
    month_label: str = "",
    project_root: Path | None = None,
    chart_images: dict[str, bytes] | None = None,
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

    Returns:
        The .pptx file as raw bytes.
    """
    template_path = _resolve_template_path(config, project_root=project_root)
    prs = Presentation(str(template_path))

    # Resolve layouts before deleting slides
    title_layout = _get_layout_by_name(prs, _LAYOUT_TITLE_SLIDE)
    transition_layout = _get_layout_by_name(prs, _LAYOUT_TRANSITION)
    content_layout = _get_layout_by_name(prs, _LAYOUT_CONTENT)

    # Chart layout is optional -- only resolve if we have chart images
    chart_layout = None
    if chart_images:
        try:
            chart_layout = _get_layout_by_name(prs, _LAYOUT_TITLE_ONLY)
        except ValueError:
            logger.warning(
                "Layout '%s' not found; chart slides will be skipped.",
                _LAYOUT_TITLE_ONLY,
            )
            chart_images = None

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
