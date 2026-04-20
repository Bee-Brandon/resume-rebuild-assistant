"""Call Claude to extract structured resume data from raw content."""
import base64
import io
import json
import logging
from datetime import datetime, timezone
from pathlib import Path

import anthropic
from PIL import Image

from config import ANTHROPIC_API_KEY, MODEL_NAME, MAX_TOKENS, API_LOG_FILE, PROMPTS_DIR

logger = logging.getLogger(__name__)

# File-based API call log
_log_path = Path(API_LOG_FILE)
_log_path.parent.mkdir(exist_ok=True)


def _log_api_call(input_type: str, page_count: int, tokens_in: int, tokens_out: int):
    cost_in = tokens_in * 3.0 / 1_000_000   # sonnet input pricing estimate
    cost_out = tokens_out * 15.0 / 1_000_000  # sonnet output pricing estimate
    total = cost_in + cost_out
    line = (
        f"{datetime.now(timezone.utc).isoformat()} | model={MODEL_NAME} | "
        f"type={input_type} | pages={page_count} | "
        f"tokens_in={tokens_in} tokens_out={tokens_out} | "
        f"est_cost=${total:.4f}\n"
    )
    with open(_log_path, "a") as f:
        f.write(line)


def _load_prompt() -> str:
    return (PROMPTS_DIR / "extraction.txt").read_text(encoding="utf-8")


def _image_to_base64(img: Image.Image) -> str:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.standard_b64encode(buf.getvalue()).decode("ascii")


def extract_structured(content: dict) -> dict:
    """Take ingestion output and return structured resume JSON.

    Args:
        content: {"type": "text"|"images", "content": str|list[Image]}

    Returns:
        Parsed resume dict matching the schema in prompts/extraction.txt

    Raises:
        RuntimeError: if Claude API fails or returns unparseable JSON after retry
    """
    # Use runtime env var (may be set via UI) over import-time config
    import os
    api_key = os.environ.get("ANTHROPIC_API_KEY") or ANTHROPIC_API_KEY
    if not api_key:
        raise RuntimeError("No API key configured. Enter your Anthropic API key in the sidebar.")
    client = anthropic.Anthropic(api_key=api_key)
    system_prompt = _load_prompt()

    # Build the user message content
    if content["type"] == "text":
        user_content = [{"type": "text", "text": f"Here is the resume text:\n\n{content['content']}"}]
        page_count = content["content"].count("\n\n") + 1  # rough estimate
    else:
        # Images — send each page as a vision attachment
        user_content = [{"type": "text", "text": "Here is the scanned resume. Extract all information you can see:"}]
        for img in content["content"]:
            b64 = _image_to_base64(img)
            user_content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/png", "data": b64},
            })
        page_count = len(content["content"])

    messages = [{"role": "user", "content": user_content}]

    # First attempt
    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=MAX_TOKENS,
            system=system_prompt,
            messages=messages,
        )
    except anthropic.APIError as e:
        raise RuntimeError(f"Claude API error: {e}") from e

    _log_api_call(
        content["type"], page_count,
        response.usage.input_tokens, response.usage.output_tokens,
    )

    raw = response.content[0].text.strip()

    # Try parsing
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # Retry with correction nudge
    messages.append({"role": "assistant", "content": raw})
    messages.append({
        "role": "user",
        "content": "Your previous response was not valid JSON. Please return only the JSON object, no markdown fences or extra text.",
    })

    try:
        response2 = client.messages.create(
            model=MODEL_NAME,
            max_tokens=MAX_TOKENS,
            system=system_prompt,
            messages=messages,
        )
    except anthropic.APIError as e:
        raise RuntimeError(f"Claude API error on retry: {e}") from e

    _log_api_call(
        content["type"] + "_retry", page_count,
        response2.usage.input_tokens, response2.usage.output_tokens,
    )

    raw2 = response2.content[0].text.strip()
    # Strip markdown fences if Claude wrapped it anyway
    if raw2.startswith("```"):
        raw2 = raw2.split("\n", 1)[1] if "\n" in raw2 else raw2[3:]
        if raw2.endswith("```"):
            raw2 = raw2[:-3]
        raw2 = raw2.strip()

    try:
        return json.loads(raw2)
    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"Claude returned invalid JSON after retry.\nParse error: {e}\nRaw output:\n{raw2[:500]}"
        ) from e
