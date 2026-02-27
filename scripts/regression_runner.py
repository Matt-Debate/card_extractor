from __future__ import annotations

import argparse
import io
import json
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from docx import Document

from extractor import extract_cards


@dataclass
class RunConfig:
    title_headings: list[int]
    tag_headings: list[int]
    include_underlined: bool
    include_highlighted: bool


def _parse_heading_list(values: Iterable[str]) -> list[int]:
    out: list[int] = []
    for v in values:
        for part in v.split(","):
            part = part.strip()
            if not part:
                continue
            out.append(int(part))
    if not out:
        raise ValueError("At least one heading level is required.")
    return out


def _load_docs(path: Path) -> list[tuple[str, bytes]]:
    if path.suffix.lower() == ".docx":
        return [(path.name, path.read_bytes())]

    if path.suffix.lower() == ".zip":
        docs: list[tuple[str, bytes]] = []
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                lower_name = info.filename.lower()
                if lower_name.startswith("__macosx/") or "/__macosx/" in lower_name:
                    continue
                if lower_name.rsplit("/", 1)[-1].startswith("._"):
                    continue
                if not lower_name.endswith(".docx"):
                    continue
                docs.append((f"{path.name}:{info.filename}", zf.read(info)))
        return docs

    raise ValueError("Input must be a .docx or .zip file.")


def _normalize_card(card: dict) -> dict:
    # Keep stable fields only to reduce noisy diffs between runs.
    return {
        "title": card.get("title", "").strip(),
        "tag": card.get("tag", "").strip(),
        "citation": card.get("citation", "").strip(),
        "marked_text": card.get("marked_text", "").strip(),
        "underlined_text": card.get("underlined_text", "").strip(),
        "highlighted_text": card.get("highlighted_text", "").strip(),
        "full_quote": card.get("full_quote", "").strip(),
    }


def run_extraction(input_path: Path, config: RunConfig) -> dict:
    docs = _load_docs(input_path)
    cards_with_source: list[dict] = []
    errors: list[str] = []

    for source_name, data in docs:
        try:
            doc = Document(io.BytesIO(data))
        except Exception as e:
            errors.append(f"{source_name}: {e}")
            continue

        cards = extract_cards(
            doc,
            title_level=config.title_headings,
            tag_level=config.tag_headings,
            include_underlined=config.include_underlined,
            include_highlighted=config.include_highlighted,
        )
        for card in cards:
            cards_with_source.append(
                {
                    "source": source_name,
                    "card": _normalize_card(card),
                }
            )

    return {
        "input": str(input_path),
        "config": asdict(config),
        "summary": {
            "docs_found": len(docs),
            "docs_failed": len(errors),
            "cards_found": len(cards_with_source),
        },
        "errors": errors,
        "cards": cards_with_source,
    }


def _print_preview(snapshot: dict, preview_count: int) -> None:
    summary = snapshot["summary"]
    print(f"Docs found:   {summary['docs_found']}")
    print(f"Docs failed:  {summary['docs_failed']}")
    print(f"Cards found:  {summary['cards_found']}")

    if snapshot["errors"]:
        print("\nErrors:")
        for err in snapshot["errors"]:
            print(f"- {err}")

    cards = snapshot["cards"][:preview_count]
    if not cards:
        return

    print("\nPreview:")
    for i, item in enumerate(cards, start=1):
        c = item["card"]
        print(f"\n[{i}] {item['source']}")
        for key in ("title", "tag", "citation", "marked_text"):
            value = c.get(key, "")
            if value:
                print(f"{key}: {value[:240]}")


def _compare_snapshots(current: dict, baseline: dict) -> tuple[bool, list[str]]:
    diffs: list[str] = []

    cur_cfg = current.get("config")
    base_cfg = baseline.get("config")
    if cur_cfg != base_cfg:
        diffs.append("Config differs between current run and baseline.")

    cur_cards = current.get("cards", [])
    base_cards = baseline.get("cards", [])
    if len(cur_cards) != len(base_cards):
        diffs.append(f"Card count changed: {len(base_cards)} -> {len(cur_cards)}")

    max_len = min(len(cur_cards), len(base_cards))
    for i in range(max_len):
        if cur_cards[i] != base_cards[i]:
            diffs.append(f"First card diff at index {i} (source={cur_cards[i].get('source')}).")
            break

    return (len(diffs) == 0), diffs


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run the card extractor on a real .docx/.zip and optionally compare snapshots."
    )
    parser.add_argument("input", help="Path to a .docx file or a .zip containing .docx files")
    parser.add_argument(
        "--title-headings",
        nargs="+",
        default=["1"],
        help="Title heading levels (space or comma separated), e.g. '1 2' or '1,2,3'",
    )
    parser.add_argument(
        "--tag-headings",
        nargs="+",
        default=["4"],
        help="Tag heading levels (space or comma separated), e.g. '3 4'",
    )
    parser.add_argument(
        "--include-underlined",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Include underlined runs in marked_text (default: true)",
    )
    parser.add_argument(
        "--include-highlighted",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="Include highlighted runs in marked_text (default: false)",
    )
    parser.add_argument(
        "--preview-count",
        type=int,
        default=3,
        help="How many cards to print as a preview (default: 3)",
    )
    parser.add_argument(
        "--save-snapshot",
        help="Write a JSON snapshot of this run to a file (for later regression comparison)",
    )
    parser.add_argument(
        "--compare-snapshot",
        help="Compare current run to a previously saved JSON snapshot",
    )

    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    if not input_path.exists():
        print(f"Input not found: {input_path}", file=sys.stderr)
        return 2

    try:
        config = RunConfig(
            title_headings=_parse_heading_list(args.title_headings),
            tag_headings=_parse_heading_list(args.tag_headings),
            include_underlined=bool(args.include_underlined),
            include_highlighted=bool(args.include_highlighted),
        )
    except ValueError as e:
        print(str(e), file=sys.stderr)
        return 2

    try:
        snapshot = run_extraction(input_path, config)
    except Exception as e:
        print(f"Run failed: {e}", file=sys.stderr)
        return 1

    _print_preview(snapshot, args.preview_count)

    if args.save_snapshot:
        out_path = Path(args.save_snapshot).expanduser()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(snapshot, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"\nSaved snapshot: {out_path}")

    if args.compare_snapshot:
        base_path = Path(args.compare_snapshot).expanduser()
        if not base_path.exists():
            print(f"Baseline snapshot not found: {base_path}", file=sys.stderr)
            return 2
        baseline = json.loads(base_path.read_text(encoding="utf-8"))
        same, diffs = _compare_snapshots(snapshot, baseline)
        if same:
            print("\nSnapshot comparison: MATCH")
        else:
            print("\nSnapshot comparison: DIFFERENT")
            for diff in diffs:
                print(f"- {diff}")
            return 3

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
