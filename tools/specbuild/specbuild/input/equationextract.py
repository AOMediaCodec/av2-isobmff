"""Extract equations from Formula-style paragraphs in Word documents.

Equations in standards documents like H.265/HEVC are typically authored as
``Formula`` or ``Equation`` styled paragraphs.  The equation body may be
plain text, Office MathML (OMML), or a mixture.  An equation number in
parentheses (e.g. ``(5-1)``) usually appears at the right margin.

This module extracts the equation content and number, and formats them
for Bikeshed output.
"""

from __future__ import annotations

import re

# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

#: Matches an equation number at the end of text: (5-1), (8-123), (A-2).
#: Supports both hyphens and en-dashes as separators.
EQ_NUMBER_RE = re.compile(r"\(([A-Za-z0-9]+[\-\u2013][A-Za-z0-9]+)\)\s*$")

#: Matches a simple standalone equation number: (1), (42).
EQ_SIMPLE_NUMBER_RE = re.compile(r"\((\d+)\)\s*$")


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------


def extract_equation(paragraph) -> dict:
    """Extract equation content and number from a Formula paragraph.

    Args:
        paragraph: A python-docx ``Paragraph`` with a Formula/Equation style.

    Returns:
        Dict with keys:

        * ``text`` — the equation body (stripped of the number)
        * ``number`` — the equation number string (e.g. ``"5-1"``) or ``None``
        * ``has_omml`` — ``True`` if Office MathML is present
    """
    text = paragraph.text.strip()

    # Try compound number first (section-number), then simple
    eq_num: str | None = None
    eq_text = text

    m = EQ_NUMBER_RE.search(text)
    if m is None:
        m = EQ_SIMPLE_NUMBER_RE.search(text)

    if m:
        eq_num = m.group(1)
        eq_text = text[: m.start()].strip()

    return {
        "text": eq_text,
        "number": eq_num,
        "has_omml": _has_omml(paragraph),
    }


def format_equation_bs(eq: dict) -> str:
    """Format an extracted equation for Bikeshed output.

    Wraps the equation text in ``$$`` delimiters.  If an equation number
    is present it is placed in a right-aligned span.

    Args:
        eq: Dict returned by :func:`extract_equation`.

    Returns:
        Bikeshed-compatible equation markup.
    """
    text = eq.get("text", "").strip()
    number = eq.get("number")

    if not text:
        return ""

    lines: list[str] = []
    lines.append("<div class='equation'>")
    lines.append(f"  $$ {text} $$")
    if number:
        lines.append(f"  <span class='equation-number'>({number})</span>")
    lines.append("</div>")
    lines.append("")

    return "\n".join(lines)


def batch_extract_equations(paragraphs: list) -> list[dict]:
    """Extract equations from a list of Formula/Equation paragraphs.

    Args:
        paragraphs: List of python-docx paragraphs already classified as
            equation type.

    Returns:
        List of equation dicts (see :func:`extract_equation`).
    """
    results: list[dict] = []
    for para in paragraphs:
        eq = extract_equation(para)
        if eq["text"]:
            results.append(eq)
    return results


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _has_omml(paragraph) -> bool:
    """Check if a paragraph contains Office Math ML elements.

    Args:
        paragraph: A python-docx ``Paragraph``.

    Returns:
        ``True`` if the paragraph's underlying XML contains ``oMath``.
    """
    try:
        xml = paragraph._element.xml
        return "oMath" in xml
    except Exception:
        return False
