from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from docx import Document
from extractor import extract_cards, format_txt, format_csv, format_csv_detailed

SAMPLE_PATH = ROOT / "samples" / "sample.docx"


def main():
    doc = Document(str(SAMPLE_PATH))
    cards = extract_cards(
        doc,
        title_level=1,
        tag_level=4,
        include_underlined=True,
        include_highlighted=True,
    )

    print(f"Cards found: {len(cards)}")
    if cards:
        print("First card:")
        print("Title:", cards[0]["title"])
        print("Tag:", cards[0]["tag"])
        print("Citation:", cards[0]["citation"])
        print("Marked:", cards[0]["marked_text"])
        print("Highlighted:", cards[0]["highlighted_text"])
        print("Underlined:", cards[0]["underlined_text"])
        print("Full quote:", cards[0]["full_quote"])

    txt = format_txt(cards)
    csv_data = format_csv(cards)
    detailed_csv_data = format_csv_detailed(cards)

    print("\nTXT preview:\n")
    print("\n".join(txt.splitlines()[:10]))

    print("\nCSV preview:\n")
    print("\n".join(csv_data.splitlines()[:5]))

    print("\nDetailed CSV preview:\n")
    print("\n".join(detailed_csv_data.splitlines()[:5]))


if __name__ == "__main__":
    main()
