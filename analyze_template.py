"""Analyze a visual resume template (PDF/image) using Claude Vision.

Sends the template image to Claude and extracts format settings that
match the visual layout — fonts, colors, spacing, columns, section order.
Returns a dict compatible with rendering.DEFAULT_FORMAT.
"""
import base64
import io
import json
import os

import anthropic
from PIL import Image

from config import MODEL_NAME, MAX_TOKENS, PROMPTS_DIR

ANALYSIS_PROMPT = """You are a resume layout analysis engine. Study this resume template image carefully and extract its visual formatting properties.

Return ONLY a valid JSON object with these exact keys. Analyze the image precisely:

{
  "layout_style": "single" or "two-column",
  "font_family": the font family used (e.g. "Calibri", "Arial", "Garamond", "Times New Roman", "Georgia"),
  "name_size": estimated font size of the person's name in points (integer, typically 18-28),
  "body_size": estimated body text font size in points (number, typically 10-12),
  "section_header_size": estimated section header font size in points (integer, typically 11-14),
  "date_size": estimated date/detail text font size in points (integer, typically 9-11),
  "line_spacing": estimated line spacing as a decimal (1.0, 1.1, 1.15, or 1.2),
  "margin_top": estimated top margin in inches (0.3 to 1.0),
  "margin_bottom": estimated bottom margin in inches (0.3 to 1.0),
  "margin_left": estimated left margin in inches (0.4 to 1.0),
  "margin_right": estimated right margin in inches (0.4 to 1.0),
  "section_spacing_before": space before section headers in points (6-14),
  "section_spacing_after": space after section headers in points (2-6),
  "paragraph_spacing": space between body paragraphs in points (1-4),
  "bullet_indent": bullet indentation in inches (0.2-0.5),
  "show_section_lines": true if section headers have a line/border underneath, false otherwise,
  "name_color": hex color of the name text (e.g. "#1a1a1a"),
  "header_color": hex color of section headers (e.g. "#1a1a2e"),
  "body_color": hex color of body text (e.g. "#333333"),
  "accent_color": hex color of dates, contact info, secondary text (e.g. "#666666"),
  "section_line_color": hex color of section divider lines (e.g. "#cccccc"),
  "section_order": array of section keys in the order they appear on the template, using these keys: "summary", "skills", "experience", "education", "certifications",
  "contact_layout": "center", "left", or "right" (how contact info is positioned relative to the name),
  "header_banner_color": hex color if there's a colored banner/header area behind the name (e.g. "#d4c5a9"), or "" if no banner,
  "tagline": "yes" if there's a subtitle/tagline line under the name (like a job title), or "" if not,
  "sidebar_sections": if layout_style is "two-column", list which sections appear in the sidebar (left column), e.g. ["education", "skills", "certifications"]. If single column, use an empty array [],
  "sidebar_width_pct": if two-column, the approximate sidebar width as a percentage (25-45). If single column, use 35,
  "template_description": a brief 1-sentence description of the template's overall style (e.g. "Clean modern single-column with blue section headers and thin divider lines")
}

Rules:
- Analyze colors by looking at the actual pixels in the image. Use hex codes.
- For fonts, identify if it's serif or sans-serif first, then guess the most likely specific font.
- For sizes, estimate based on the relative proportions in the image.
- section_order should only include sections that are visible in the template.
- Return ONLY the JSON object, no markdown fences, no explanation.
"""


def _image_to_base64(img: Image.Image) -> str:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.standard_b64encode(buf.getvalue()).decode("ascii")


def analyze_template_image(img: Image.Image) -> dict:
    """Send a template image to Claude Vision and extract format settings.

    Args:
        img: PIL Image of the resume template

    Returns:
        Dict of format settings compatible with rendering.DEFAULT_FORMAT
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        from config import ANTHROPIC_API_KEY
        api_key = ANTHROPIC_API_KEY
    if not api_key:
        raise RuntimeError("No API key available for template analysis.")

    client = anthropic.Anthropic(api_key=api_key)
    b64 = _image_to_base64(img)

    response = client.messages.create(
        model=MODEL_NAME,
        max_tokens=MAX_TOKENS,
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": ANALYSIS_PROMPT},
                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            ],
        }],
    )

    raw = response.content[0].text.strip()

    # Strip markdown fences if present
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1] if "\n" in raw else raw[3:]
        if raw.endswith("```"):
            raw = raw[:-3]
        raw = raw.strip()

    try:
        result = json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"Claude returned invalid JSON for template analysis:\n{raw[:500]}") from e

    # Extract and clean the description before returning
    description = result.pop("template_description", "Imported template")

    # If tagline is "yes", set it to empty string (user fills in their own)
    if result.get("tagline") == "yes":
        result["tagline"] = ""

    return result, description


def load_template_image(filepath: str) -> Image.Image:
    """Load a template file (PDF or image) as a PIL Image."""
    from pathlib import Path
    p = Path(filepath)
    ext = p.suffix.lower()

    if ext == ".pdf":
        from pdf2image import convert_from_path
        from config import SCAN_DPI
        images = convert_from_path(str(p), dpi=min(SCAN_DPI, 200), first_page=1, last_page=1)
        return images[0]
    elif ext in (".png", ".jpg", ".jpeg"):
        return Image.open(p).convert("RGB")
    else:
        raise ValueError(f"Unsupported template format: {ext}. Use PDF, PNG, or JPG.")
