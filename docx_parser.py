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
from urllib.parse import unquote

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
                runs.append(TextRun(text=text, bold=bold, underline=True, hyperlink=clean_url(item.url)))
    return runs


def clean_url(url):
    """Repair hyperlink targets mangled by copy/pasting from Word Online,
    e.g. 'https://auc-word-edit.officeapps.live.com/we/8.%09https:/www.myer
    .com.au/p/...' -> 'https://www.myer.com.au/p/...': the real link is
    whatever follows the *last* 'http(s):/', with its missing slash restored."""
    if not url:
        return None
    url = unquote(url).strip()
    starts = [m.start() for m in re.finditer(r"https?:/", url, re.IGNORECASE)]
    if starts:
        url = url[starts[-1]:]
    url = re.sub(r"^(https?):/+", r"\1://", url, flags=re.IGNORECASE)
    return url.strip() or None


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
        intro_paragraphs, products, outro_paragraphs = _split_product_entries(rich_paragraphs)
        section["copy_runs"] = intro_paragraphs
        section["copy"] = _plain_text_from_rich_paragraphs(intro_paragraphs)
        section["copy_after_runs"] = outro_paragraphs
        section["copy_after"] = _plain_text_from_rich_paragraphs(outro_paragraphs)
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
    """If `para` is a dedicated product-URL marker line, return (url,
    inline_name), else None. Matches a Word hyperlink run whose target is a
    Myer product page (`/p/...`), or a plain-text paragraph that is nothing
    but such a URL (copy/pasted or generated content doesn't always carry a
    real hyperlink field) -- independent of bullet/bold formatting, since
    that varies between briefs and isn't a reliable signal on its own.

    inline_name is the hyperlink's display text when it's a product name
    rather than the URL itself (e.g. "Philips Baristina" linked to its
    product page), in which case the name line *is* the marker line.
    """
    runs = [r for r in para["runs"] if r.text and r.text.strip()]
    if not runs:
        return None
    if len(runs) == 1 and runs[0].hyperlink and MYER_PRODUCT_URL_RE.match(runs[0].hyperlink.strip()):
        text = runs[0].text.strip()
        return runs[0].hyperlink.strip(), (None if URL_IN_TEXT_RE.match(text) else text)
    full_text = "".join(r.text for r in runs).strip()
    if MYER_PRODUCT_URL_RE.match(full_text):
        return full_text, None
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

    If no URL markers are found, falls back to _split_field_records (the
    "Name: / Type: / Pros: / ..." layout).

    Returns (intro_paragraphs, products, outro_paragraphs), where
    intro_paragraphs is whatever body copy precedes the first product (or
    all of it, if no product pattern is found), outro_paragraphs is body
    copy after the last product (field-record layout only), and products
    is a list of {"name": str, "url": str, "paragraphs": [rich paragraphs...]}.
    """
    starts = [(i, *marker) for i, p in enumerate(rich_paragraphs) if (marker := _product_url_marker(p))]
    if not starts:
        return _split_field_records(rich_paragraphs)

    intro = rich_paragraphs[: starts[0][0]]
    products = []
    for idx, (start, url, inline_name) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(rich_paragraphs)
        name_idx = start + 1
        if inline_name:
            name = inline_name
            desc_paragraphs = rich_paragraphs[start + 1 : end]
        elif name_idx < end:
            name = "".join(r.text for r in rich_paragraphs[name_idx]["runs"]).strip()
            desc_paragraphs = rich_paragraphs[name_idx + 1 : end]
        else:
            name = ""
            desc_paragraphs = []
        products.append({"name": name, "url": url, "paragraphs": desc_paragraphs})

    return intro, products, []


# "Label: value" lines, e.g. "Pros: Easy milk drinks" -- label is 1-4 words.
FIELD_LINE_RE = re.compile(r"^\s*([A-Za-z][A-Za-z'&/\-]*(?:\s+[A-Za-z'&/\-]+){0,3})\s*:\s*(\S.*)$")
NAME_LABELS = {"name", "product", "product name", "model"}
URL_IN_TEXT_RE = re.compile(r"https?://\S+", re.IGNORECASE)
MIN_RECORDS = 2
MIN_REPEATED_FIELDS = 3


def _para_text(para):
    return "".join(r.text for r in para["runs"]).strip()


def _field_label(para):
    match = FIELD_LINE_RE.match(_para_text(para))
    return match.group(1).strip().lower() if match else None


def _split_field_records(rich_paragraphs):
    """Detect a product comparison written as repeated labelled fields, e.g.

        Name: Nespresso Vertuo Creatista   (name often hyperlinked)
        Type: Capsule
        Pros: ...  Cons: ...  Star Rating: 4.5/5  Price: $879.20
        Name: De'Longhi Eletta Ultra
        ...

    A section is treated as a product grid when the same set of field
    labels (at least MIN_REPEATED_FIELDS of them) repeats for at least
    MIN_RECORDS items -- ordinary copy like a "How to choose" list of
    one-off "Budget: / Size:" tips never repeats its labels, so it's left
    alone. A record starts at each occurrence of the "Name:" field if there
    is one, otherwise at each occurrence of the first field label (taking an
    unlabelled one-line name directly above it as the product name).

    The product URL is the record's first hyperlink (usually on the name),
    falling back to a plain-text URL anywhere in the record.
    """
    labels = [_field_label(p) for p in rich_paragraphs]
    counts = {}
    for label in labels:
        if label:
            counts[label] = counts.get(label, 0) + 1
    repeated = {label for label, n in counts.items() if n >= MIN_RECORDS}
    if len(repeated) < MIN_REPEATED_FIELDS:
        return rich_paragraphs, [], []

    name_labels = [l for l in repeated if l in NAME_LABELS]
    if name_labels:
        start_label = name_labels[0]
    else:
        start_label = next(l for l in labels if l in repeated)

    starts = [i for i, label in enumerate(labels) if label == start_label]

    # Unlabelled name line directly above each start (e.g. "Philips Baristina"
    # then "Type: ..."): include it in the record as the name. Decided from
    # the 2nd record on, since only those are sandwiched between records --
    # the 1st record's line above could just as well be intro copy.
    name_above = (
        start_label not in NAME_LABELS
        and starts[0] > 0
        and all(labels[i - 1] is None and _looks_like_name(rich_paragraphs[i - 1]) for i in starts[1:])
    )
    if name_above:
        starts = [i - 1 for i in starts]

    # The last record ends at the first paragraph that isn't one of the
    # repeated fields (closing copy after the grid).
    end = starts[-1] + (2 if name_above else 1)
    while end < len(rich_paragraphs) and labels[end] in repeated:
        end += 1

    intro = rich_paragraphs[: starts[0]]
    outro = rich_paragraphs[end:]
    products = []
    for idx, start in enumerate(starts):
        stop = starts[idx + 1] if idx + 1 < len(starts) else end
        products.append(_field_record_to_product(rich_paragraphs[start:stop], name_above))
    return intro, products, outro


def _looks_like_name(para):
    text = _para_text(para)
    return 0 < len(text) <= 80 and not text.endswith((".", "?", "!", ":"))


def _field_record_to_product(paragraphs, name_above):
    name, name_idx = "", None
    if name_above:
        name, name_idx = _para_text(paragraphs[0]), 0
    else:
        for i, para in enumerate(paragraphs):
            label = _field_label(para)
            if label in NAME_LABELS:
                name, name_idx = FIELD_LINE_RE.match(_para_text(para)).group(2).strip(), i
                break

    url, url_only_idx = None, None
    ordered = ([paragraphs[name_idx]] if name_idx is not None else []) + paragraphs
    for para in ordered:
        link = next((r.hyperlink for r in para["runs"] if r.hyperlink), None)
        if link:
            url = link
            break
    if not url:
        for i, para in enumerate(paragraphs):
            match = URL_IN_TEXT_RE.search(_para_text(para))
            if match:
                url = match.group(0)
                if match.group(0) == _para_text(para):
                    url_only_idx = i
                break

    desc = [p for i, p in enumerate(paragraphs) if i not in (name_idx, url_only_idx)]
    return {"name": name, "url": url, "paragraphs": desc}
