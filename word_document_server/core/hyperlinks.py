"""
Hyperlink helpers for Word Document Server.

Builds real Word hyperlink elements (w:hyperlink with an r:id pointing at an
external relationship) inside python-docx paragraphs. Uses python-docx's
``part.relate_to`` so the docx rels part is maintained correctly and duplicate
URLs are deduped into a single relationship.
"""
from typing import Optional

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


_DEFAULT_HYPERLINK_COLOR = "0563C1"


def _normalize_url(url: str) -> str:
    url = url.strip()
    if not url:
        raise ValueError("Hyperlink url must not be empty")
    if "://" not in url and not url.startswith(("mailto:", "tel:", "#")):
        url = "https://" + url
    return url


def _has_style(paragraph, style_name: str) -> bool:
    try:
        doc = paragraph.part.document
    except AttributeError:
        return False
    try:
        _ = doc.styles[style_name]
        return True
    except KeyError:
        return False


def add_hyperlink_run(
    paragraph,
    url: str,
    text: str,
    *,
    style: Optional[str] = "Hyperlink",
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = True,
    color: Optional[str] = _DEFAULT_HYPERLINK_COLOR,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
):
    """Append a clickable hyperlink run to ``paragraph``.

    Args:
        paragraph: python-docx Paragraph to append to.
        url: Target URL. A scheme is added if missing.
        text: Visible link text. Empty falls back to ``url``.
        style: Run style to apply (default ``"Hyperlink"``). If the doc does
            not define it, falls back to inline color/underline formatting.
        bold/italic/underline: Optional run formatting overrides.
        color: Hex RGB (no ``#``) applied when the Hyperlink style is absent
            or to force a color.
        font_name/font_size: Optional font overrides.

    Returns:
        The inserted ``w:hyperlink`` lxml element.
    """
    if not isinstance(url, str):
        raise TypeError("url must be a string")
    url = _normalize_url(url)
    display_text = text if text else url

    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    use_style = bool(style) and _has_style(paragraph, style)
    if use_style:
        rStyle = OxmlElement("w:rStyle")
        rStyle.set(qn("w:val"), style)
        rPr.append(rStyle)
    else:
        if color:
            c = OxmlElement("w:color")
            c.set(qn("w:val"), color.lstrip("#"))
            rPr.append(c)
        if underline:
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            rPr.append(u)

    if bold:
        rPr.append(OxmlElement("w:b"))
    if italic:
        rPr.append(OxmlElement("w:i"))
    if font_name:
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)
        rPr.append(rFonts)
    if font_size:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(font_size) * 2))  # half-points
        rPr.append(sz)

    new_run.append(rPr)

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = display_text
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def wrap_run_as_hyperlink(run, url: str, *, preserve_run: bool = True):
    """Wrap an existing python-docx ``Run`` so it becomes a hyperlink.

    Replaces the ``<w:r>`` element in-place with ``<w:hyperlink><w:r/></w:hyperlink>``
    whose ``r:id`` points at the registered external relationship. Existing
    run formatting (``rPr``) is preserved when ``preserve_run`` is True.
    """
    if not isinstance(url, str):
        raise TypeError("url must be a string")
    url = _normalize_url(url)

    paragraph_part = run.part
    r_id = paragraph_part.relate_to(url, RT.HYPERLINK, is_external=True)

    r_elem = run._r
    parent = r_elem.getparent()
    index = list(parent).index(r_elem)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    parent.remove(r_elem)
    if preserve_run:
        hyperlink.append(r_elem)
    parent.insert(index, hyperlink)
    return hyperlink
