import io
import csv
from docx.enum.text import WD_COLOR_INDEX
from openpyxl import Workbook
from openpyxl.styles import Alignment


def is_heading(para, level):
    name = (para.style.name or "").lower()
    return f"heading {level}" in name or name == f"heading{level}"

def paragraph_has_marked_run(para):
    for run in para.runs:
        if not run.text or not run.text.strip():
            continue
        if run.font.underline:
            return True
        highlight = run.font.highlight_color
        if highlight and highlight != WD_COLOR_INDEX.AUTO:
            return True
    return False


def is_citation_like(text):
    t = text.strip()
    if not t:
        return False
    lower = t.lower()
    if t.startswith("[") or t.endswith("]"):
        return True
    if "http://" in lower or "https://" in lower or "www." in lower:
        return True
    if "accessed" in lower or "doa" in lower:
        return True
    return False


def extract_cards(doc, title_level, tag_level, include_underlined=True, include_highlighted=False):
    cards = []

    current_title = ""
    current_tag = ""
    current_citation = ""
    current_marked_runs = []
    current_underlined_runs = []
    current_highlighted_runs = []
    current_full_quote_lines = []

    in_card = False
    collect_quote_marked = False

    def flush_card():
        if current_tag.strip() and current_citation.strip():
            cards.append(
                {
                    "title": current_title.strip(),
                    "tag": current_tag.strip(),
                    "citation": current_citation.strip(),
                    "marked_text": " ".join(
                        t.strip() for t in current_marked_runs if t.strip()
                    ),
                    "underlined_text": " ".join(
                        t.strip() for t in current_underlined_runs if t.strip()
                    ),
                    "highlighted_text": " ".join(
                        t.strip() for t in current_highlighted_runs if t.strip()
                    ),
                    "full_quote": "\n".join(
                        t.strip() for t in current_full_quote_lines if t.strip()
                    ),
                }
            )

    paragraphs = doc.paragraphs
    n = len(paragraphs)
    i = 0

    while i < n:
        para = paragraphs[i]
        text = para.text.strip()

        try:
            if is_heading(para, title_level):
                current_title = text
                i += 1
                continue

            if is_heading(para, tag_level):
                if in_card:
                    flush_card()

                current_tag = text
                current_citation = ""
                current_marked_runs = []
                current_underlined_runs = []
                current_highlighted_runs = []
                current_full_quote_lines = []
                in_card = True
                collect_quote_marked = False

                j = i + 1
                while j < n:
                    nxt = paragraphs[j]
                    nxt_text = nxt.text.strip()
                    if not nxt_text:
                        j += 1
                        continue
                    if is_heading(nxt, title_level) or is_heading(nxt, tag_level):
                        break
                    citation_lines = [nxt_text]

                    k = j + 1
                    while k < n:
                        nxt2 = paragraphs[k]
                        nxt2_text = nxt2.text.strip()
                        if not nxt2_text:
                            break
                        if is_heading(nxt2, title_level) or is_heading(nxt2, tag_level):
                            break
                        if paragraph_has_marked_run(nxt2):
                            break
                        if is_citation_like(nxt2_text):
                            citation_lines.append(nxt2_text)
                            k += 1
                            continue
                        break

                    current_citation = "\n".join(citation_lines)
                    collect_quote_marked = True
                    i = k - 1
                    break

                i += 1
                continue

            if in_card and collect_quote_marked:
                if is_heading(para, title_level) or is_heading(para, tag_level):
                    collect_quote_marked = False
                    i += 1
                    continue

                for run in para.runs:
                    if not run.text or not run.text.strip():
                        continue

                    is_underlined = bool(run.font.underline)
                    highlight = run.font.highlight_color
                    is_highlighted = bool(highlight) and highlight != WD_COLOR_INDEX.AUTO

                    if (include_underlined and is_underlined) or (
                        include_highlighted and is_highlighted
                    ):
                        current_marked_runs.append(run.text)
                    if is_underlined:
                        current_underlined_runs.append(run.text)
                    if is_highlighted:
                        current_highlighted_runs.append(run.text)

                if para.text and para.text.strip():
                    current_full_quote_lines.append(para.text)

            i += 1

        except Exception:
            i += 1

    if in_card:
        flush_card()

    return cards


def format_txt(cards):
    lines = []
    for c in cards:
        if c["title"].strip():
            lines.append(c["title"].strip())
        if c["tag"].strip():
            lines.append(c["tag"].strip())
        if c["citation"].strip():
            lines.append(c["citation"].strip())
        if c["marked_text"].strip():
            lines.append(c["marked_text"].strip())
        lines.append("")
        lines.append("=" * 60)
        lines.append("")
    text = "\n".join(lines).rstrip() + "\n"
    return text


def format_csv(cards):
    output = io.StringIO()
    writer = csv.writer(output)
    for c in cards:
        col1 = c["title"].strip()
        parts = []
        if c["tag"].strip():
            parts.append(c["tag"].strip())
        if c["citation"].strip():
            parts.append(c["citation"].strip())
        if c["marked_text"].strip():
            parts.append(c["marked_text"].strip())
        col2 = "\n".join(parts)
        writer.writerow([col1, col2])
    return output.getvalue()


def format_csv_detailed(cards):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(
        [
            "title",
            "tag",
            "citation",
            "highlighted",
            "underlined",
            "full_quote",
        ]
    )
    for c in cards:
        writer.writerow(
            [
                c["title"].strip(),
                c["tag"].strip(),
                c["citation"].strip(),
                c["highlighted_text"].strip(),
                c["underlined_text"].strip(),
                c["full_quote"].strip(),
            ]
        )
    return output.getvalue()


def _workbook_to_bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def format_xlsx(cards):
    wb = Workbook()
    ws = wb.active
    ws.title = "cards"

    ws.append(["title", "card"])
    for c in cards:
        parts = []
        if c["tag"].strip():
            parts.append(c["tag"].strip())
        if c["citation"].strip():
            parts.append(c["citation"].strip())
        if c["marked_text"].strip():
            parts.append(c["marked_text"].strip())
        card_text = "\n".join(parts)
        ws.append([c["title"].strip(), card_text])

    wrap = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.alignment = wrap

    return _workbook_to_bytes(wb)


def format_xlsx_detailed(cards):
    wb = Workbook()
    ws = wb.active
    ws.title = "cards_detailed"

    ws.append(
        [
            "title",
            "tag",
            "citation",
            "highlighted",
            "underlined",
            "full_quote",
        ]
    )
    for c in cards:
        ws.append(
            [
                c["title"].strip(),
                c["tag"].strip(),
                c["citation"].strip(),
                c["highlighted_text"].strip(),
                c["underlined_text"].strip(),
                c["full_quote"].strip(),
            ]
        )

    wrap = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.alignment = wrap

    return _workbook_to_bytes(wb)
