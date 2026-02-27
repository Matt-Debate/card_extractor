import io
import zipfile
import streamlit as st
from docx import Document

from extractor import (
    extract_cards,
    format_txt,
    format_csv,
    format_csv_detailed_with_source,
    format_xlsx,
    format_xlsx_detailed_with_source,
)

st.set_page_config(page_title="Card Extractor", layout="wide")

st.title("Card Extractor")
st.write("Upload a .docx file to extract debate cards. Configure which headings mark titles and tags.")

with st.sidebar:
    st.subheader("Header Settings")
    title_heading_options = {
        "Auto (Heading 1-3)": [1, 2, 3],
        "Heading 1 or 2": [1, 2],
        "Heading 1": 1,
        "Heading 2": 2,
    }
    tag_heading_options = {
        "Heading 3 or 4": [3, 4],
        "Heading 3": 3,
        "Heading 4": 4,
    }
    title_heading = st.selectbox(
        "Title Heading",
        list(title_heading_options.keys()),
        index=0,
    )
    tag_heading = st.selectbox(
        "Tag Heading",
        list(tag_heading_options.keys()),
        index=2,
    )
    st.subheader("Text Extraction")
    include_underlined = st.checkbox("Include underlined text", value=False)
    include_highlighted = st.checkbox("Include highlighted text", value=True)


title_level = title_heading_options[title_heading]
tag_level = tag_heading_options[tag_heading]

title_levels = title_level if isinstance(title_level, (list, tuple, set)) else [title_level]
tag_levels = tag_level if isinstance(tag_level, (list, tuple, set)) else [tag_level]

if set(title_levels) & set(tag_levels):
    st.warning("Title Heading and Tag Heading are the same. This may reduce accuracy.")
if not include_underlined and not include_highlighted:
    st.warning("Both text extraction options are disabled. No marked text will be captured.")

uploaded_files = st.file_uploader(
    "Upload .docx files or a .zip containing .docx files",
    type=["docx", "zip"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Upload a .docx file to begin.")
else:
    zip_payloads = {}
    total_docs = 0
    intake_errors = []

    for upload_idx, uploaded in enumerate(uploaded_files):
        name = uploaded.name
        lower_name = name.lower()
        if lower_name.endswith(".docx"):
            total_docs += 1
            continue

        if lower_name.endswith(".zip"):
            try:
                zip_bytes = uploaded.getvalue()
                member_names = []
                with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                    for info in zf.infolist():
                        if info.is_dir():
                            continue
                        filename = info.filename
                        member_lower = filename.lower()
                        if member_lower.startswith("__macosx/") or "/__macosx/" in member_lower:
                            continue
                        base_name = member_lower.rsplit("/", 1)[-1]
                        if base_name.startswith("._"):
                            continue
                        if not member_lower.endswith(".docx"):
                            continue
                        member_names.append(filename)
                if member_names:
                    zip_payloads[upload_idx] = (zip_bytes, member_names)
                    total_docs += len(member_names)
            except zipfile.BadZipFile as e:
                intake_errors.append(f"Invalid ZIP file: {name} ({e})")
            continue

        intake_errors.append(f"Unsupported file type: {name}")

    if not total_docs:
        st.info("No .docx files found. Upload .docx files or a .zip containing .docx files.")
        st.stop()

    cards = []
    cards_with_source = []
    errors = []
    parse_warnings = []
    progress = st.progress(0)
    status = st.empty()
    processed_docs = [0]

    def process_document(source_name, data):
        status.text(f"Processing {processed_docs[0] + 1}/{total_docs}: {source_name}")
        try:
            doc = Document(io.BytesIO(data))
        except Exception as e:
            errors.append(f"{source_name}: {e}")
            processed_docs[0] += 1
            progress.progress(int(processed_docs[0] / total_docs * 100))
            return

        doc_parse_errors = []
        extracted = extract_cards(
            doc,
            title_levels,
            tag_level,
            include_underlined=include_underlined,
            include_highlighted=include_highlighted,
            parse_errors=doc_parse_errors,
        )
        for c in extracted:
            cards.append(c)
            cards_with_source.append((source_name, c))

        if doc_parse_errors:
            parse_warnings.append((source_name, doc_parse_errors))

        processed_docs[0] += 1
        progress.progress(int(processed_docs[0] / total_docs * 100))

    for upload_idx, uploaded in enumerate(uploaded_files):
        name = uploaded.name
        lower_name = name.lower()
        if lower_name.endswith(".docx"):
            process_document(name, uploaded.getvalue())
            continue

        if lower_name.endswith(".zip"):
            payload = zip_payloads.get(upload_idx)
            if payload is None:
                continue
            zip_bytes, member_names = payload
            try:
                with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                    for member_name in member_names:
                        source_name = f"{name}:{member_name}"
                        try:
                            member_data = zf.read(member_name)
                        except Exception as e:
                            errors.append(f"{source_name}: {e}")
                            processed_docs[0] += 1
                            progress.progress(int(processed_docs[0] / total_docs * 100))
                            continue
                        process_document(source_name, member_data)
            except zipfile.BadZipFile as e:
                errors.append(f"Invalid ZIP file during processing: {name} ({e})")

    status.text(f"Processed {processed_docs[0]} document(s).")

    st.write(f"Uploads received: {len(uploaded_files)}")
    st.write(f"Documents processed: {processed_docs[0]}")
    st.write(f"Cards found: {len(cards)}")

    if intake_errors:
        st.warning("Some uploads could not be indexed:")
        for err in intake_errors:
            st.text(err)

    if errors:
        st.warning("Some files could not be read:")
        for err in errors:
            st.text(err)

    if parse_warnings:
        total_parse_warnings = sum(len(w) for _, w in parse_warnings)
        st.warning(f"Parser warnings encountered: {total_parse_warnings}")
        with st.expander("Parser warning details (sample)"):
            shown = 0
            max_shown = 100
            for source_name, warnings_for_doc in parse_warnings:
                for warning in warnings_for_doc:
                    st.text(f"{source_name}: {warning}")
                    shown += 1
                    if shown >= max_shown:
                        st.text(f"... truncated at {max_shown} warnings")
                        break
                if shown >= max_shown:
                    break

    if not cards:
        st.info("No cards extracted. Check heading selections and document formatting.")
    else:
        txt_data = format_txt(cards)
        csv_data = format_csv(cards)
        detailed_csv_data = format_csv_detailed_with_source(cards_with_source)
        xlsx_data = format_xlsx(cards)
        detailed_xlsx_data = format_xlsx_detailed_with_source(cards_with_source)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                "Download TXT",
                data=txt_data,
                file_name="cards_extracted.txt",
                mime="text/plain",
            )
        with col2:
            st.download_button(
                "Download XLSX",
                data=xlsx_data,
                file_name="cards_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col3:
            st.download_button(
                "Download Detailed XLSX",
                data=detailed_xlsx_data,
                file_name="cards_extracted_detailed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with st.expander("CSV (optional)"):
            st.write("CSV can lose formatting in Excel; XLSX is recommended.")
            csv_col1, csv_col2 = st.columns(2)
            with csv_col1:
                st.download_button(
                    "Download CSV",
                    data=csv_data,
                    file_name="cards_extracted.csv",
                    mime="text/csv",
                )
            with csv_col2:
                st.download_button(
                    "Download Detailed CSV",
                    data=detailed_csv_data,
                    file_name="cards_extracted_detailed.csv",
                    mime="text/csv",
                )

        with st.expander("Preview (first 3 cards)"):
            for idx, (source_name, c) in enumerate(cards_with_source[:3], start=1):
                st.markdown(f"**Card {idx}**")
                if len(uploaded_files) > 1:
                    st.caption(f"Source: {source_name}")
                preview_lines = []
                if c["title"].strip():
                    preview_lines.append(c["title"].strip())
                if c["tag"].strip():
                    preview_lines.append(c["tag"].strip())
                if c["citation"].strip():
                    preview_lines.append(c["citation"].strip())
                if c["marked_text"].strip():
                    preview_lines.append(c["marked_text"].strip())
                st.text("\n".join(preview_lines))
