import io
import csv
import os
import re
from docx.enum.text import WD_COLOR_INDEX
from openpyxl import Workbook
from openpyxl.styles import Alignment


def is_heading(para, level_or_levels):
    name = (para.style.name or "").lower()
    if isinstance(level_or_levels, (list, tuple, set)):
        levels = level_or_levels
    else:
        levels = [level_or_levels]
    for level in levels:
        if f"heading {level}" in name or name == f"heading{level}":
            return True
    return False

def normalize_separators(s: str) -> str:
    dash_chars = "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"
    s = re.sub(f"[{re.escape(dash_chars)}]", "-", s)
    s = s.replace("~~", "-")
    s = re.sub(r"-{2,}", "-", s)
    s = s.strip(" -")
    return s


def strip_numeric_prefix_tokens(tokens):
    out = tokens[:]
    while out and re.fullmatch(r"\[?\d+\]?", out[0]):
        out.pop(0)
    return out


def tournament_short_name(raw: str) -> str:
    r = raw.strip()
    r_low = r.lower()
    if "university-of-pennsylvania" in r_low or re.search(r"\bupenn\b", r_low):
        if "liberty-bell" in r_low:
            return "UPenn Liberty Bell"
        return "UPenn"
    if "stanford" in r_low:
        if "invitational" in r_low or "invitationals" in r_low:
            return "Stanford Invitational"
        if "online" in r_low:
            return "Stanford Online Invitational"
        return "Stanford"
    if "bellaire" in r_low:
        return "Bellaire"
    if "pennsbury" in r_low:
        return "Pennsbury"
    if "unlv" in r_low or "golden-desert" in r_low:
        return "UNLV Golden Desert"
    cleaned = re.sub(r"\s+", " ", r.replace("-", " ")).strip()
    return " ".join(w.upper() if w.isupper() and len(w) <= 5 else w.capitalize() for w in cleaned.split())


ROUND_NUM_RE = re.compile(r"(?:^|-)Round-(\d+)$", re.IGNORECASE)
SPECIAL_ROUNDS = {
    "Octas",
    "Quarters",
    "Semis",
    "Finals",
    "Doubles",
    "Triples",
    "All-Rounds",
    "All",
}


def split_school_team_side(name_no_ext):
    s = normalize_separators(name_no_ext)
    parts = [p for p in s.split("-") if p]
    if len(parts) < 4:
        return None, None, None, ""
    school, team_code, side = parts[0], parts[1], parts[2]
    rest = "-".join(parts[3:])
    return school, team_code, side, rest


def parse_round_and_tournament(rest):
    rest_norm = normalize_separators(rest)
    tokens = [t for t in rest_norm.split("-") if t]
    tokens = strip_numeric_prefix_tokens(tokens)
    if not tokens:
        return None, None, None

    m = ROUND_NUM_RE.search("-" + tokens[-2] + "-" + tokens[-1]) if len(tokens) >= 2 else None
    if m:
        round_number = int(m.group(1))
        tournament_tokens = tokens[:-2]
        tournament_raw = "-".join(tournament_tokens) if tournament_tokens else None
        return tournament_raw, round_number, "Round"

    last = tokens[-1]
    if last in SPECIAL_ROUNDS:
        tournament_tokens = tokens[:-1]
        tournament_raw = "-".join(tournament_tokens) if tournament_tokens else None
        return tournament_raw, None, last

    return "-".join(tokens), None, None


def parse_filename_metadata(source_name):
    if ":" in source_name:
        source_name = source_name.split(":", 1)[1]
    base = os.path.basename(source_name)
    name_no_ext = os.path.splitext(base)[0]

    school, team_code, side, rest = split_school_team_side(name_no_ext)
    side_norm = None
    if side:
        side_upper = side.strip().lower()
        if side_upper in ("pro", "con"):
            side_norm = side_upper.capitalize()
        else:
            side_norm = side

    tournament_raw, round_number, round_label = parse_round_and_tournament(rest)
    tournament = tournament_short_name(tournament_raw) if tournament_raw else None

    if round_number is not None:
        round_value = f"Round {round_number}"
    elif round_label:
        round_value = round_label
    else:
        round_value = None

    return {
        "filename": base,
        "school": school,
        "team_code": team_code,
        "side": side_norm,
        "tournament_raw": tournament_raw,
        "tournament": tournament,
        "round": round_value,
    }


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
    if isinstance(title_level, (list, tuple, set)):
        title_levels = list(title_level)
    else:
        title_levels = [title_level]
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
            if is_heading(para, title_levels):
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
                    if is_heading(nxt, title_levels) or is_heading(nxt, tag_level):
                        break
                    citation_lines = [nxt_text]

                    k = j + 1
                    while k < n:
                        nxt2 = paragraphs[k]
                        nxt2_text = nxt2.text.strip()
                        if not nxt2_text:
                            break
                        if is_heading(nxt2, title_levels) or is_heading(nxt2, tag_level):
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
                if is_heading(para, title_levels) or is_heading(para, tag_level):
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
    raise NotImplementedError("Use format_csv_detailed_with_source instead.")


def format_csv_detailed_with_source(cards_with_source):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(
        [
            "filename",
            "school",
            "team_code",
            "side",
            "tournament_raw",
            "tournament",
            "round",
            "title",
            "tag",
            "citation",
            "highlighted",
            "underlined",
            "full_quote",
        ]
    )
    for source_name, c in cards_with_source:
        meta = parse_filename_metadata(source_name)
        writer.writerow(
            [
                meta["filename"] or "",
                meta["school"] or "",
                meta["team_code"] or "",
                meta["side"] or "",
                meta["tournament_raw"] or "",
                meta["tournament"] or "",
                meta["round"] or "",
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


def format_xlsx_detailed_with_source(cards_with_source):
    wb = Workbook()
    ws = wb.active
    ws.title = "cards_detailed"

    ws.append(
        [
            "filename",
            "school",
            "team_code",
            "side",
            "tournament_raw",
            "tournament",
            "round",
            "title",
            "tag",
            "citation",
            "highlighted",
            "underlined",
            "full_quote",
        ]
    )
    for source_name, c in cards_with_source:
        meta = parse_filename_metadata(source_name)
        ws.append(
            [
                meta["filename"] or "",
                meta["school"] or "",
                meta["team_code"] or "",
                meta["side"] or "",
                meta["tournament_raw"] or "",
                meta["tournament"] or "",
                meta["round"] or "",
                c["title"].strip(),
                c["tag"].strip(),
                c["citation"].strip(),
                c["highlighted_text"].strip(),
                c["underlined_text"].strip(),
                c["full_quote"].strip(),
            ]
        )

    wrap = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=13):
        for cell in row:
            cell.alignment = wrap

    return _workbook_to_bytes(wb)
