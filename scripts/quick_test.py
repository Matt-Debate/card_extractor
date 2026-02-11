from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from docx import Document
from extractor import extract_cards, format_txt, format_csv

SAMPLE_PATH = ROOT / "samples" / "sample.docx"


def main():
    doc = Document(str(SAMPLE_PATH))
    cards = extract_cards(doc, title_level=1, tag_level=4)

    print(f"Cards found: {len(cards)}")
    if cards:
        print("First card:")
        print("Title:", cards[0]["title"])
        print("Tag:", cards[0]["tag"])
        print("Citation:", cards[0]["citation"])
        print("Underlined:", cards[0]["underlined_text"])

    txt = format_txt(cards)
    csv_data = format_csv(cards)

    print("\nTXT preview:\n")
    print("\n".join(txt.splitlines()[:10]))

    print("\nCSV preview:\n")
    print("\n".join(csv_data.splitlines()[:5]))


if __name__ == "__main__":
    main()
