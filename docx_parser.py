"""
Parses a Jaywing-style Word content brief (.docx) into a structured dict
that xlsx_builder.py can turn into a Myer-format content-brief spreadsheet.

The agency brief template has two halves:
  1. A metadata/instructions half (Job Purpose, Target Audience, URL/Meta
     tables, keyword table, then an instructional "Content Brief" section
     with character counts and writer directions).
  2. A "Draft Content" half containing the actual finished copy, laid out
     as "H1:", "H2:", "H3:" headings each followed by one or more
     paragraphs, with occasional bracketed design notes like
     "[Myer to product carousel grid to single mattresses]".

Only the "Draft Content" half is used for the heading/copy rows in the
output brief; the metadata half is used for the reference sheet.
"""

import re
from dataclasses import dataclass

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.run import Run


@dataclass
class TextRun:
    """A run of text with the formatting that matters for the Excel output:
    bold (Word direct-bold), underline (always True for hyperlinks), and the
    hyperlink target URL (if this run is a hyperlink)."""

    text: str
    bold: bool = False
    underline: bool = False
    hyperlink: str = None

HEADING_RE = re.compile(r"^H([1-3])\s*:?\s*(.*)$", re.IGNORECASE)
DASH_LINE_RE = re.compile(r"^-{3,}\s*$")
ANCHOR_LINE_RE = re.compile(r"^Myer to Insert Anchor Links\s*:\s*(.*)$", re.IGNORECASE)
DRAFT_INSTRUCTION_RE = re.compile(r"^Copywriter to draft the content", re.IGNORECASE)

KNOWN_SECTIONS = [
    "Job Purpose",
    "Type of content",
    "Brand & TOV Guidelines",
    "Target Audience",
    "URL",
    "Meta Title",
    "Meta Description",
    "Main Keywords",
    "Semantically Related Keywords",
    "Word Count",
    "Inspirational content",
    "Page Template",
    "Type of content to include",
    "SEO Best Practice Recommendations",
    "Content Brief",
    "Draft Content",
]
KNOWN_SECTIONS_LOWER = {s.lower(): s for s in KNOWN_SECTIONS}


class BriefParseError(ValueError):
    """Raised when a .docx doesn't look like the expected brief template."""


def iter_block_items(document):
    """Yield paragraphs and tables from the document body, in document order."""
    parent_elm = document.element.body
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


def _extract_runs(paragraph):
    """Walk a paragraph's runs and hyperlinks in document order, preserving
    bold and hyperlink formatting. Uses python-docx's iter_inner_content()
    (1.1+) so hyperlink-wrapped runs aren't silently dropped the way plain
    paragraph.runs would drop them."""
    runs = []
    for item in paragraph.iter_inner_content():
        if isinstance(item, Run):
            if item.text:
                runs.append(TextRun(text=item.text, bold=bool(item.bold)))
        elif isinstance(item, Hyperlink):
            text = item.text
            if text:
                bold = bool(item.runs[0].bold) if item.runs else False
                runs.append(TextRun(text=text, bold=bold, underline=True, hyperlink=item.address))
    return runs


def _is_bullet_paragraph(paragraph):
    """True if the paragraph is a Word bulleted-list item (has a <w:numPr>
    list-numbering reference in its paragraph properties)."""
    p_pr = paragraph._p.find(qn("w:pPr"))
    if p_pr is None:
        return False
    return p_pr.find(qn("w:numPr")) is not None


def _plain_text_from_rich_paragraphs(rich_paragraphs):
    """Flatten rich paragraphs to plain text (bullets rendered as '• ', a
    blank line inserted between a bullet list and the body copy either side
    of it), used for row-height estimation and as a fallback if rich text
    can't be rendered. Kept in sync with xlsx_builder._build_rich_copy's
    paragraph-join rules so estimated row heights match what's rendered."""
    lines = []
    prev_bullet = None
    for para in rich_paragraphs:
        text = "".join(run.text for run in para["runs"]).strip()
        if not text:
            continue  # drop blank/whitespace-only paragraphs (soft-break noise)
        if para["bullet"]:
            text = "• " + text
        if lines and para["bullet"] != prev_bullet:
            lines.append("")  # blank line between a bullet list and body copy
        lines.append(text)
        prev_bullet = para["bullet"]
    return "\n".join(lines)


def _table_rows_text(table):
    """Flatten a table's rows, including rows from any nested tables (Word
    lets a table cell contain its own sub-table, which python-docx's
    cell.text ignores)."""
    rows = []
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        rows.append(cells)
        for cell in row.cells:
            for nested in cell.tables:
                rows.extend(_table_rows_text(nested))
    return rows


def parse_docx(path):
    document = Document(path)
    items = list(iter_block_items(document))

    sections = {}  # header text -> {"paragraphs": [...], "tables": [...]}
    header_table = None
    current_section = None
    draft_start_index = None

    for idx, item in enumerate(items):
        if isinstance(item, Paragraph):
            text = item.text.strip()
            matched_header = KNOWN_SECTIONS_LOWER.get(text.lower())
            if matched_header:
                current_section = matched_header
                sections.setdefault(current_section, {"paragraphs": [], "tables": []})
                if matched_header == "Draft Content":
                    draft_start_index = idx
                    break
                continue
            if not text or DASH_LINE_RE.match(text):
                continue
            if current_section:
                sections[current_section]["paragraphs"].append(text)
        else:  # Table
            if current_section:
                sections[current_section]["tables"].append(_table_rows_text(item))
            elif header_table is None:
                header_table = _table_rows_text(item)

    if draft_start_index is None:
        raise BriefParseError(
            "Couldn't find a 'Draft Content' section in this document. "
            "This app expects the standard Jaywing content-brief template."
        )

    brief_info = _extract_brief_info(header_table, sections)
    draft_sections, anchors = _parse_draft_content(items[draft_start_index + 1 :])

    if not draft_sections:
        raise BriefParseError(
            "Found a 'Draft Content' section but no H1/H2/H3 headings inside it."
        )

    content_page_name = draft_sections[0]["heading"] if draft_sections else None

    return {
        "content_page_name": content_page_name,
        "anchors": anchors,
        "sections": draft_sections,
        "brief_info": brief_info,
    }


def _first_row_value(table_rows, col=0):
    """For a simple 2-row [header, value] table, return the value cell."""
    if not table_rows:
        return ""
    data_row = table_rows[1] if len(table_rows) > 1 else table_rows[0]
    return data_row[col] if col < len(data_row) else ""


def _extract_brief_info(header_table, sections):
    info = {
        "copywriter": "",
        "client": "",
        "date": "",
        "job_purpose": "",
        "type_of_content": "",
        "brand_tov": [],
        "target_audience": "",
        "existing_url": "",
        "meta_title": "",
        "meta_title_pixels": "",
        "meta_description": "",
        "meta_description_pixels": "",
        "keywords": [],  # list of (type, keyword, avg_monthly_searches, ranking)
        "related_keywords": [],
    }

    if header_table:
        for row in header_table:
            if len(row) == 2:
                key, value = row[0].strip().lower(), row[1].strip()
                if "copywriter" in key:
                    info["copywriter"] = value
                elif "client" in key:
                    info["client"] = value
                elif key == "date":
                    info["date"] = value

    if "Job Purpose" in sections:
        info["job_purpose"] = "\n".join(sections["Job Purpose"]["paragraphs"])
    if "Type of content" in sections:
        info["type_of_content"] = "\n".join(sections["Type of content"]["paragraphs"])
    if "Brand & TOV Guidelines" in sections:
        info["brand_tov"] = sections["Brand & TOV Guidelines"]["paragraphs"]
    if "Target Audience" in sections:
        info["target_audience"] = "\n".join(sections["Target Audience"]["paragraphs"])

    if "URL" in sections and sections["URL"]["tables"]:
        info["existing_url"] = _first_row_value(sections["URL"]["tables"][0])

    if "Meta Title" in sections and sections["Meta Title"]["tables"]:
        rows = sections["Meta Title"]["tables"][0]
        info["meta_title"] = _first_row_value(rows, col=0)
        info["meta_title_pixels"] = _first_row_value(rows, col=1)

    if "Meta Description" in sections and sections["Meta Description"]["tables"]:
        rows = sections["Meta Description"]["tables"][0]
        info["meta_description"] = _first_row_value(rows, col=0)
        info["meta_description_pixels"] = _first_row_value(rows, col=1)

    if "Main Keywords" in sections and sections["Main Keywords"]["tables"]:
        rows = sections["Main Keywords"]["tables"][0]
        for row in rows[1:]:  # skip header row
            if len(row) >= 4 and any(c for c in row):
                info["keywords"].append(tuple(row[:4]))

    if "Semantically Related Keywords" in sections:
        info["related_keywords"] = sections["Semantically Related Keywords"]["paragraphs"]

    return info


def _parse_draft_content(items):
    anchors = []
    draft_sections = []
    current = None

    note_buffer = None  # lines of a bracketed note that spans multiple paragraphs

    for item in items:
        if not isinstance(item, Paragraph):
            continue  # tables inside Draft Content aren't expected; skip safely
        text = item.text.strip()
        if not text or DASH_LINE_RE.match(text):
            continue

        if note_buffer is not None:
            note_buffer.append(text)
            if text.endswith("]"):
                if current:
                    current["notes"].append(" ".join(note_buffer))
                note_buffer = None
            continue

        if DRAFT_INSTRUCTION_RE.match(text):
            continue

        anchor_match = ANCHOR_LINE_RE.match(text)
        if anchor_match:
            anchors = [a.strip() for a in anchor_match.group(1).split("|") if a.strip()]
            continue

        heading_match = HEADING_RE.match(text)
        if heading_match:
            current = {
                "level": int(heading_match.group(1)),
                "heading": heading_match.group(2).strip(),
                "rich_paragraphs": [],
                "notes": [],
            }
            draft_sections.append(current)
            continue

        if text.startswith("["):
            if text.endswith("]"):
                if current:
                    current["notes"].append(text)
            else:
                note_buffer = [text]
            continue

        if current:
            runs = _extract_runs(item)
            if runs:
                current["rich_paragraphs"].append(
                    {"runs": runs, "bullet": _is_bullet_paragraph(item)}
                )

    for section in draft_sections:
        rich_paragraphs = section.pop("rich_paragraphs")
        intro_paragraphs, products = _split_product_entries(rich_paragraphs)
        section["copy_runs"] = intro_paragraphs
        section["copy"] = _plain_text_from_rich_paragraphs(intro_paragraphs)
        section["products"] = [
            {
                "name": p["name"],
                "url": p["url"],
                "copy_runs": p["paragraphs"],
                "copy": _plain_text_from_rich_paragraphs(p["paragraphs"]),
            }
            for p in products
        ]
        section["note"] = "\n".join(section.pop("notes"))

    return draft_sections, anchors


MYER_PRODUCT_URL_RE = re.compile(r"^https?://(?:www\.)?myer\.com\.au/p/\S+$", re.IGNORECASE)


def _product_url_marker(para):
    """If `para` is a dedicated product-URL marker line, return that URL,
    else None. Matches a Word hyperlink run whose target is a Myer product
    page (`/p/...`), or a plain-text paragraph that is nothing but such a
    URL (copy/pasted or generated content doesn't always carry a real
    hyperlink field) -- independent of bullet/bold formatting, since that
    varies between briefs and isn't a reliable signal on its own.
    """
    runs = [r for r in para["runs"] if r.text and r.text.strip()]
    if not runs:
        return None
    if len(runs) == 1 and runs[0].hyperlink and MYER_PRODUCT_URL_RE.match(runs[0].hyperlink.strip()):
        return runs[0].hyperlink.strip()
    full_text = "".join(r.text for r in runs).strip()
    if MYER_PRODUCT_URL_RE.match(full_text):
        return full_text
    return None


def _split_product_entries(rich_paragraphs):
    """Detect a repeating "product recommendation" pattern within a
    section's body copy: a paragraph that is nothing but a Myer product
    URL (`_product_url_marker`), immediately followed by the product's
    name, followed by plain description lines (best-for/pros/cons/rating/
    cost, or whatever the brief includes) up to the next such URL marker
    (or the end of the section). This shape is how briefs write out named
    product picks/comparison grids (e.g. "The Best Kettles To Buy") -- one
    URL+name+description group per product -- and is generic enough to
    apply to any brief that follows the same convention, not just one
    specific document or product category.

    Returns (intro_paragraphs, products), where intro_paragraphs is
    whatever body copy precedes the first product (or all of it, if no
    product pattern is found) and products is a list of
    {"name": str, "url": str, "paragraphs": [rich paragraphs...]}.
    """
    starts = [(i, url) for i, p in enumerate(rich_paragraphs) if (url := _product_url_marker(p))]
    if not starts:
        return rich_paragraphs, []

    intro = rich_paragraphs[: starts[0][0]]
    products = []
    for idx, (start, url) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(rich_paragraphs)
        name_idx = start + 1
        if name_idx < end:
            name = "".join(r.text for r in rich_paragraphs[name_idx]["runs"]).strip()
            desc_paragraphs = rich_paragraphs[name_idx + 1 : end]
        else:
            name = ""
            desc_paragraphs = []
        products.append({"name": name, "url": url, "paragraphs": desc_paragraphs})

    return intro, products
