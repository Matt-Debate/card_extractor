from docx import Document
import os

doc_path = "/Users/matthewalanfarmer/Downloads/BasisLangleyIndependent-RaCh-Pro-02---40th-Annual-Stanford-Invitational-Finals.docx"

try:
    doc = Document(doc_path)
except Exception as e:
    raise SystemExit(f"Error loading document: {e}")

cards = []

current_title = ""
current_tag = ""
current_citation = ""
current_underlined_runs = []

in_card = False
collect_quote_underlines = False  # IMPORTANT: starts only after citation line

def flush_card():
    """Save current card if it has at least tag + citation."""
    if current_tag.strip() and current_citation.strip():
        cards.append({
            "title": current_title.strip(),
            "tag": current_tag.strip(),
            "citation": current_citation.strip(),
            "underlined_text": " ".join(t.strip() for t in current_underlined_runs if t.strip())
        })

# helper: robust heading detection
def is_heading(para, level):
    # catches normal style names and occasional custom variants
    name = (para.style.name or "").lower()
    return f"heading {level}" in name or name == f"heading{level}"

paragraphs = doc.paragraphs
n = len(paragraphs)
i = 0

while i < n:
    para = paragraphs[i]
    text = para.text.strip()

    try:
        # Title block (argument-level title)
        if is_heading(para, 1):
            current_title = text
            i += 1
            continue

        # New card starts at Heading 4 (tag)
        if is_heading(para, 4):
            # close previous card
            if in_card:
                flush_card()

            # start new card
            current_tag = text
            current_citation = ""
            current_underlined_runs = []
            in_card = True
            collect_quote_underlines = False  # don't collect tag underlines

            # citation is first non-empty non-heading paragraph after tag
            j = i + 1
            while j < n:
                nxt = paragraphs[j]
                nxt_text = nxt.text.strip()
                if not nxt_text:
                    j += 1
                    continue
                if is_heading(nxt, 1) or is_heading(nxt, 4):
                    # no citation found before next heading
                    break
                current_citation = nxt_text
                collect_quote_underlines = True  # start collecting AFTER citation
                i = j  # move pointer to citation line
                break

            i += 1
            continue

        # collect quote underlines only inside card and only after citation
        if in_card and collect_quote_underlines:
            # stop collecting if we hit another heading
            if is_heading(para, 1) or is_heading(para, 4):
                # do not consume; loop will process heading next iteration
                collect_quote_underlines = False
                i += 1
                continue

            # collect underlined runs from quote/body
            for run in para.runs:
                if run.font.underline and run.text and run.text.strip():
                    current_underlined_runs.append(run.text)

        i += 1

    except Exception as e:
        print(f"Warning: failed at paragraph {i}: {e}")
        i += 1

# flush final card
if in_card:
    flush_card()

# write output (no labels, keep citation)
out_dir = "/Users/matthewalanfarmer/Documents/Cards_Extracted"
os.makedirs(out_dir, exist_ok=True)
out_path = os.path.join(out_dir, "cards_extracted.txt")

try:
    with open(out_path, "w", encoding="utf-8") as f:
        for c in cards:
            # order: title -> tag -> citation -> underlined text
            if c["title"].strip():
                f.write(c["title"].strip() + "\n")
            if c["tag"].strip():
                f.write(c["tag"].strip() + "\n")
            if c["citation"].strip():
                f.write(c["citation"].strip() + "\n")
            if c["underlined_text"].strip():
                f.write(c["underlined_text"].strip() + "\n")

            # separator between cards
            f.write("\n" + "=" * 60 + "\n\n")

    print(f"Wrote {len(cards)} cards to {out_path}")
except Exception as e:
    raise SystemExit(f"Error writing output file: {e}")