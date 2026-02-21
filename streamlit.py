import io
import streamlit as st
from docx import Document

from extractor import (
    extract_cards,
    format_txt,
    format_csv,
    format_csv_detailed,
    format_xlsx,
    format_xlsx_detailed,
)

st.set_page_config(page_title="Card Extractor", layout="wide")

st.title("Card Extractor")
st.write("Upload a .docx file to extract debate cards. Configure which headings mark titles and tags.")

with st.sidebar:
    st.subheader("Header Settings")
    heading_options = {
        "Heading 1 or 2": [1, 2],
        "Heading 1": 1,
        "Heading 2": 2,
        "Heading 3": 3,
        "Heading 4": 4,
    }
    title_heading = st.selectbox(
        "Title Heading",
        list(heading_options.keys()),
        index=0,
    )
    tag_heading = st.selectbox(
        "Tag Heading",
        list(heading_options.keys()),
        index=3,
    )
    st.subheader("Text Extraction")
    include_underlined = st.checkbox("Include underlined text", value=False)
    include_highlighted = st.checkbox("Include highlighted text", value=True)


title_level = heading_options[title_heading]
tag_level = heading_options[tag_heading]

title_levels = title_level if isinstance(title_level, (list, tuple, set)) else [title_level]

if tag_level in title_levels:
    st.warning("Title Heading and Tag Heading are the same. This may reduce accuracy.")
if not include_underlined and not include_highlighted:
    st.warning("Both text extraction options are disabled. No marked text will be captured.")

uploaded_files = st.file_uploader("Upload one or more .docx files", type=["docx"], accept_multiple_files=True)

if not uploaded_files:
    st.info("Upload a .docx file to begin.")
else:
    cards = []
    cards_with_source = []
    errors = []

    for uploaded in uploaded_files:
        try:
            doc = Document(io.BytesIO(uploaded.getvalue()))
        except Exception as e:
            errors.append(f"{uploaded.name}: {e}")
            continue

        extracted = extract_cards(
            doc,
            title_levels,
            tag_level,
            include_underlined=include_underlined,
            include_highlighted=include_highlighted,
        )
        for c in extracted:
            cards.append(c)
            cards_with_source.append((uploaded.name, c))

    st.write(f"Files processed: {len(uploaded_files)}")
    st.write(f"Cards found: {len(cards)}")

    if errors:
        st.warning("Some files could not be read:")
        for err in errors:
            st.text(err)

    if not cards:
        st.info("No cards extracted. Check heading selections and document formatting.")
    else:
        txt_data = format_txt(cards)
        csv_data = format_csv(cards)
        detailed_csv_data = format_csv_detailed(cards)
        xlsx_data = format_xlsx(cards)
        detailed_xlsx_data = format_xlsx_detailed(cards)

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
