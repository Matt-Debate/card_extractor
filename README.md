# Card Extractor

Extracts debate cards from `.docx` files and exports them as TXT/CSV/XLSX.

The default extraction flow is heading-based:
- Title: configurable heading level(s)
- Tag: configurable heading level(s)
- Citation: first non-empty paragraph after tag (with continuation heuristics)
- Quote text: underlined and/or highlighted runs after citation until next heading

## Quick Start

```bash
pip install -r requirements.txt
streamlit run streamlit.py
```

Then upload one or more `.docx` files, or a `.zip` containing `.docx` files.

## Streamlit App

Main app: `streamlit.py`

Features:
- Configurable title/tag heading detection
- Underline/highlight extraction toggles
- Multi-file and ZIP upload
- Progress + processing status
- Export options:
  - `cards_extracted.txt`
  - `cards_extracted.xlsx`
  - `cards_extracted_detailed.xlsx`
  - Optional CSV exports
- Parse warning surface (sampled) for debugging malformed content

## Output Formats

- **Standard TXT/CSV/XLSX**
  - Title
  - Tag
  - Citation
  - Marked text (based on underline/highlight settings)

- **Detailed CSV/XLSX**
  - Source filename + parsed filename metadata (school/team/side/tournament/round)
  - Citation metadata (oral citation + URL)
  - Highlighted text, underlined text, and full quote text

## Scripts

- `scripts/quick_test.py`
  - Smoke test using `samples/sample.docx`
  - Prints previews and verifies exporters

- `scripts/regression_runner.py`
  - Runs extraction on a real `.docx` or `.zip`
  - Saves JSON snapshots for regression testing
  - Compares current extraction vs baseline snapshot

Example:

```bash
python3 scripts/regression_runner.py "/path/to/input.zip" \
  --title-headings 1 2 \
  --tag-headings 3 4 \
  --include-underlined \
  --include-highlighted \
  --save-snapshot tmp/baseline.snapshot.json
```

Compare later:

```bash
python3 scripts/regression_runner.py "/path/to/input.zip" \
  --title-headings 1 2 \
  --tag-headings 3 4 \
  --include-underlined \
  --include-highlighted \
  --compare-snapshot tmp/baseline.snapshot.json
```

## Project Structure

- `streamlit.py`: UI and batch processing flow
- `extractor.py`: extraction + formatting logic
- `scripts/`: local test and regression helpers
- `samples/sample.docx`: test fixture

## Notes

- `main.py` and checked-in generated output files were removed as legacy artifacts.
- For large mixed-format batches, use `scripts/regression_runner.py` to tune settings and compare behavior across extractor changes.
