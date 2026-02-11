import io
import csv
from docx.enum.text import WD_COLOR_INDEX


def is_heading(para, level):
    name = (para.style.name or "").lower()
    return f"heading {level}" in name or name == f"heading{level}"


def extract_cards(doc, title_level, tag_level, include_underlined=True, include_highlighted=False):
    cards = []

    current_title = ""
    current_tag = ""
    current_citation = ""
    current_underlined_runs = []

    in_card = False
    collect_quote_marked = False

    def flush_card():
        if current_tag.strip() and current_citation.strip():
            cards.append(
                {
                    "title": current_title.strip(),
                    "tag": current_tag.strip(),
                    "citation": current_citation.strip(),
                    "underlined_text": " ".join(
                        t.strip() for t in current_underlined_runs if t.strip()
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
                current_underlined_runs = []
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
                    current_citation = nxt_text
                    collect_quote_marked = True
                    i = j
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

                    is_underlined = bool(run.font.underline) if include_underlined else False
                    is_highlighted = False
                    if include_highlighted:
                        highlight = run.font.highlight_color
                        is_highlighted = bool(highlight) and highlight != WD_COLOR_INDEX.AUTO

                    if is_underlined or is_highlighted:
                        current_underlined_runs.append(run.text)

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
        if c["underlined_text"].strip():
            lines.append(c["underlined_text"].strip())
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
        if c["underlined_text"].strip():
            parts.append(c["underlined_text"].strip())
        col2 = "\n".join(parts)
        writer.writerow([col1, col2])
    return output.getvalue()
