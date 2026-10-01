import streamlit as st
import fitz
from docx import Document
import requests
from langdetect import detect
import tempfile
import os
from pathlib import Path
import time


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DocMorph | Smart Document Translator",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# SESSION STATE
# ============================================================

DEFAULTS = {
    "extracted_pages": [],
    "source_language": "",
    "translated_file": None,
    "translated_paragraphs": [],
    "page_count": 0,
    "input_extension": "",
    "original_file_bytes": None,
    "original_file_name": "",
    "last_uploaded_file": ""
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# SIMPLE CSS
# ============================================================

st.markdown(
    """
    <style>

    .main {
        padding-top: 1rem;
    }

    .hero-box {
        padding: 30px;
        border-radius: 24px;
        margin-bottom: 25px;
        border: 1px solid rgba(255,255,255,0.12);
        background: linear-gradient(
            135deg,
            rgba(80,70,180,0.22),
            rgba(0,170,255,0.12)
        );
    }

    .step-card {
        padding: 18px 12px;
        border-radius: 18px;
        border: 1px solid rgba(255,255,255,0.10);
        background: rgba(255,255,255,0.045);
        text-align: center;
        min-height: 100px;
    }

    .step-number {
        font-size: 11px;
        color: #8f96b5;
        letter-spacing: 1px;
        font-weight: 700;
    }

    .step-icon {
        font-size: 25px;
        margin-top: 5px;
    }

    .step-title {
        font-size: 15px;
        font-weight: 700;
        margin-top: 5px;
    }

    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.045);
        border: 1px solid rgba(255,255,255,0.08);
        padding: 14px;
        border-radius: 16px;
    }

    button[kind="primary"] {
        border-radius: 14px !important;
        font-weight: 700 !important;
    }

    section[data-testid="stFileUploaderDropzone"] {
        border-radius: 18px;
        border: 1px dashed rgba(130,120,255,0.45);
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HERO
# ============================================================

st.title("🌐 DocMorph")

st.subheader("✨ AI-Assisted Document Workspace")

st.write(
    "Smart Document Language Converter — upload a PDF, DOCX, or TXT "
    "file, detect its language, translate it into your selected "
    "language, preview the result, and export it as a DOCX document."
)


# ============================================================
# WORKFLOW
# ============================================================

st.markdown("### 🔄 How DocMorph Works")

steps = [
    ("01", "📥", "INPUT"),
    ("02", "🔍", "ANALYZE"),
    ("03", "🌐", "TRANSLATE"),
    ("04", "👁️", "PREVIEW & EXPORT")
]

workflow_cols = st.columns(4)

for workflow_index, (
    column,
    step
) in enumerate(zip(workflow_cols, steps)):

    number, icon, title = step

    with column:

        st.markdown(
            f"""
            <div class="step-card">
                <div class="step-number">{number}</div>
                <div class="step-icon">{icon}</div>
                <div class="step-title">{title}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

st.write("")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown("## ⚙️ DocMorph")

    st.caption(
        "Smart document translation workspace"
    )

    st.divider()

    # ========================================================
    # ADDED ONLY: SIMPLE DOCX UPLOAD
    # ========================================================

    st.markdown("### 📄 Upload Document")

    st.caption(
        "Upload your DOCX file here."
    )

    sidebar_file = st.file_uploader(
        "Choose DOCX",
        type=["docx"],
        key="sidebar_docx_upload"
    )

    if sidebar_file is not None:

        st.success(
            "✅ Document uploaded!"
        )

    # ========================================================
    # ORIGINAL SIDEBAR CODE
    # ========================================================

    st.markdown("### 📄 Supported Files")

    st.write("• PDF")
    st.write("• DOCX")
    st.write("• TXT")

    st.markdown("### 🌐 Target Languages")

    st.write("• English")
    st.write("• Kannada")
    st.write("• Telugu")
    st.write("• Tamil")

    st.divider()

    st.caption(
        "Translation powered by the MyMemory API."
    )


# ============================================================
# LANGUAGE MAP
# ============================================================

language_map = {
    "English": "en",
    "Kannada": "kn",
    "Telugu": "te",
    "Tamil": "ta"
}


source_names = {
    "en": "English",
    "kn": "Kannada",
    "te": "Telugu",
    "ta": "Tamil",
    "hi": "Hindi",
    "fr": "French",
    "de": "German",
    "es": "Spanish",
    "it": "Italian",
    "pt": "Portuguese",
    "ar": "Arabic",
    "ja": "Japanese",
    "ko": "Korean",
    "zh-cn": "Chinese",
    "zh-tw": "Chinese",
    "pa": "Punjabi"
}


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf_pages(file_bytes):

    pages = []

    pdf = fitz.open(
        stream=file_bytes,
        filetype="pdf"
    )

    for page_number, page in enumerate(pdf):

        text = page.get_text("text").strip()

        if text:

            pages.append(
                {
                    "page": page_number + 1,
                    "text": text
                }
            )

    pdf.close()

    return pages


# ============================================================
# DOCX EXTRACTION
# ============================================================

def extract_docx(file_bytes):

    temp_path = os.path.join(
        tempfile.gettempdir(),
        "docmorph_input.docx"
    )

    with open(temp_path, "wb") as file:

        file.write(file_bytes)

    document = Document(temp_path)

    paragraphs = []

    for index, paragraph in enumerate(
        document.paragraphs
    ):

        text = paragraph.text.strip()

        if text:

            paragraphs.append(
                {
                    "index": index,
                    "page": 1,
                    "text": text
                }
            )

    try:

        os.remove(temp_path)

    except OSError:

        pass

    return paragraphs


# ============================================================
# TXT EXTRACTION
# ============================================================

def extract_txt(file_bytes):

    text = file_bytes.decode(
        "utf-8",
        errors="ignore"
    ).strip()

    if not text:

        return []

    return [
        {
            "page": 1,
            "text": text
        }
    ]


# ============================================================
# TEXT CHUNKING
# ============================================================

def create_chunks(
    text,
    max_chars=450
):

    text = text.strip()

    if not text:

        return []

    chunks = []

    start = 0

    while start < len(text):

        end = min(
            start + max_chars,
            len(text)
        )

        if end < len(text):

            space_position = text.rfind(
                " ",
                start,
                end
            )

            if space_position > start:

                end = space_position

        chunk = text[
            start:end
        ].strip()

        if chunk:

            chunks.append(chunk)

        start = end

    return chunks


# ============================================================
# TRANSLATE ONE CHUNK
# ============================================================

def translate_chunk(
    text,
    source_language,
    target_language
):

    if not text.strip():

        return ""

    if source_language == target_language:

        return text

    url = (
        "https://api.mymemory.translated.net/get"
    )

    params = {
        "q": text,
        "langpair": (
            f"{source_language}|"
            f"{target_language}"
        )
    }

    response = requests.get(
        url,
        params=params,
        timeout=30
    )

    if response.status_code != 200:

        raise Exception(
            "Translation server returned "
            f"status {response.status_code}"
        )

    result = response.json()

    response_data = result.get(
        "responseData",
        {}
    )

    translated_text = response_data.get(
        "translatedText"
    )

    if not translated_text:

        raise Exception(
            "Translation service did not "
            "return translated text."
        )

    return translated_text


# ============================================================
# TRANSLATE COMPLETE TEXT
# ============================================================

def translate_text(
    text,
    source_language,
    target_language
):

    if not text.strip():

        return ""

    chunks = create_chunks(
        text,
        max_chars=450
    )

    translated_chunks = []

    for chunk_index, chunk in enumerate(
        chunks
    ):

        translated = translate_chunk(
            chunk,
            source_language,
            target_language
        )

        translated_chunks.append(
            translated
        )

        if chunk_index < len(chunks) - 1:

            time.sleep(0.4)

    return "\n\n".join(
        translated_chunks
    )


# ============================================================
# CREATE DOCX PRESERVING ORIGINAL PARAGRAPHS
# ============================================================

def create_translated_docx_preserve_structure(
    original_bytes,
    translated_paragraphs
):

    temp_input = os.path.join(
        tempfile.gettempdir(),
        "docmorph_original.docx"
    )

    output_path = os.path.join(
        tempfile.gettempdir(),
        "DocMorph_Translated.docx"
    )

    with open(temp_input, "wb") as file:

        file.write(original_bytes)

    document = Document(temp_input)

    translated_map = {
        item["index"]: item["translated"]
        for item in translated_paragraphs
    }

    for index, paragraph in enumerate(
        document.paragraphs
    ):

        if index not in translated_map:

            continue

        translated_text = translated_map[
            index
        ]

        if paragraph.runs:

            paragraph.runs[0].text = (
                translated_text
            )

            for run in paragraph.runs[1:]:

                run.text = ""

        else:

            paragraph.add_run(
                translated_text
            )

    document.save(output_path)

    try:

        os.remove(temp_input)

    except OSError:

        pass

    return output_path


# ============================================================
# CREATE DOCX FOR PDF / TXT
# ============================================================

def create_simple_translated_docx(
    translated_items
):

    output_path = os.path.join(
        tempfile.gettempdir(),
        "DocMorph_Translated.docx"
    )

    document = Document()

    document.add_heading(
        "DocMorph - Translated Document",
        level=0
    )

    document.add_paragraph(
        "Generated using DocMorph"
    )

    for item in translated_items:

        translated = item.get(
            "translated",
            ""
        )

        if not translated.strip():

            continue

        paragraphs = translated.split(
            "\n\n"
        )

        for paragraph_text in paragraphs:

            if paragraph_text.strip():

                document.add_paragraph(
                    paragraph_text.strip()
                )

    document.save(output_path)

    return output_path


# ============================================================
# DOCUMENT INPUT
# ============================================================

st.markdown("## 📥 Document Input")

uploaded_file = st.file_uploader(
    "Upload your document",
    type=[
        "pdf",
        "docx",
        "txt"
    ],
    help=(
        "Supported formats: "
        "PDF, DOCX and TXT"
    ),
    key="document_uploader"
)


# ============================================================
# HANDLE UPLOAD
# ============================================================

if uploaded_file is not None:

    current_file_name = uploaded_file.name

    if (
        st.session_state.last_uploaded_file
        != current_file_name
    ):

        st.session_state.extracted_pages = []

        st.session_state.source_language = ""

        st.session_state.translated_file = None

        st.session_state.translated_paragraphs = []

        st.session_state.page_count = 0

        st.session_state.input_extension = ""

        st.session_state.original_file_bytes = None

        st.session_state.original_file_name = ""

        st.session_state.last_uploaded_file = (
            current_file_name
        )

    st.success(
        "✅ Document uploaded successfully!"
    )

    # ========================================================
    # ANALYZE BUTTON
    # ========================================================

    if st.button(
        "🔍 Analyze Document",
        use_container_width=True,
        key="analyze_document_button"
    ):

        file_bytes = uploaded_file.getvalue()

        extension = Path(
            uploaded_file.name
        ).suffix.lower()

        try:

            # ------------------------------------------------
            # Extract text
            # ------------------------------------------------

            if extension == ".pdf":

                pages = extract_pdf_pages(
                    file_bytes
                )

            elif extension == ".docx":

                pages = extract_docx(
                    file_bytes
                )

            elif extension == ".txt":

                pages = extract_txt(
                    file_bytes
                )

            else:

                raise Exception(
                    "Unsupported file format."
                )

            # ------------------------------------------------
            # Remove empty sections
            # ------------------------------------------------

            pages = [
                page
                for page in pages
                if page["text"].strip()
            ]

            if not pages:

                raise Exception(
                    "No readable text was "
                    "found in the document."
                )

            # ------------------------------------------------
            # Save extracted information
            # ------------------------------------------------

            st.session_state.extracted_pages = (
                pages
            )

            st.session_state.page_count = (
                len(pages)
            )

            st.session_state.input_extension = (
                extension
            )

            st.session_state.original_file_bytes = (
                file_bytes
            )

            st.session_state.original_file_name = (
                uploaded_file.name
            )

            st.session_state.translated_file = (
                None
            )

            st.session_state.translated_paragraphs = (
                []
            )

            # ------------------------------------------------
            # Detect source language
            # ------------------------------------------------

            combined_text = " ".join(
                item["text"]
                for item in pages
                if item["text"]
            )

            try:

                detected = detect(
                    combined_text[:3000]
                )

                st.session_state.source_language = (
                    detected
                )

            except Exception:

                st.session_state.source_language = (
                    "en"
                )

            st.success(
                "✅ Document analyzed successfully!"
            )

        except Exception as error:

            st.error(
                f"❌ Analysis error: {error}"
            )


# ============================================================
# DOCUMENT INTELLIGENCE
# ============================================================

if st.session_state.extracted_pages:

    st.markdown(
        "## 🔍 Document Intelligence"
    )

    info1, info2, info3 = st.columns(3)

    with info1:

        st.metric(
            "Pages / Sections",
            st.session_state.page_count
        )

    with info2:

        source_code = (
            st.session_state.source_language
        )

        source_display = source_names.get(
            source_code,
            source_code
        )

        st.metric(
            "Detected Language",
            source_display
        )

    with info3:

        total_chars = sum(
            len(item["text"])
            for item in
            st.session_state.extracted_pages
        )

        st.metric(
            "Characters",
            total_chars
        )

    st.write("")


    # ========================================================
    # LANGUAGE SELECTION
    # ========================================================

    language_col1, language_col2 = (
        st.columns(2)
    )

    with language_col1:

        st.info(
            f"🔎 Source Language: "
            f"**{source_display}**"
        )

    with language_col2:

        target_language_name = st.selectbox(
            "🌐 Choose Target Language",
            list(language_map.keys()),
            key="target_language_selector"
        )

        target_language = language_map[
            target_language_name
        ]


    # ========================================================
    # ORIGINAL CONTENT PREVIEW
    # ========================================================

    with st.expander(
        "👁️ Preview Extracted Content",
        expanded=False
    ):

        for preview_index, item in enumerate(
            st.session_state.extracted_pages
        ):

            label = item.get(
                "page",
                item.get(
                    "index",
                    preview_index + 1
                )
            )

            st.markdown(
                f"**Section / Page {label}**"
            )

            preview = item["text"]

            if len(preview) > 1000:

                preview = (
                    preview[:1000]
                    + "..."
                )

            st.text_area(
                "Extracted text",
                preview,
                height=150,
                key=(
                    f"original_preview_"
                    f"{preview_index}"
                ),
                disabled=True
            )


    # ========================================================
    # TRANSLATE DOCUMENT
    # ========================================================

    if st.button(
        "🌐 Translate Document",
        type="primary",
        use_container_width=True,
        key="translate_document_button"
    ):

        progress = st.progress(0)

        status = st.empty()

        translated_items = []

        total = len(
            st.session_state.extracted_pages
        )

        try:

            source_language = (
                st.session_state.source_language
            )

            if not source_language:

                raise Exception(
                    "Source language could "
                    "not be detected."
                )

            # ------------------------------------------------
            # Translate each section
            # ------------------------------------------------

            for index, item in enumerate(
                st.session_state.extracted_pages
            ):

                status.write(
                    f"🌐 Translating section "
                    f"{index + 1} of {total}..."
                )

                translated = translate_text(
                    item["text"],
                    source_language,
                    target_language
                )

                translated_items.append(
                    {
                        **item,
                        "translated": translated
                    }
                )

                progress.progress(
                    int(
                        (
                            (index + 1)
                            / total
                        ) * 80
                    )
                )

            # ------------------------------------------------
            # Create output
            # ------------------------------------------------

            status.write(
                "📄 Creating translated DOCX..."
            )

            if (
                st.session_state.input_extension
                == ".docx"
            ):

                output_path = (
                    create_translated_docx_preserve_structure(
                        st.session_state.original_file_bytes,
                        translated_items
                    )
                )

            else:

                output_path = (
                    create_simple_translated_docx(
                        translated_items
                    )
                )

            # ------------------------------------------------
            # Save result
            # ------------------------------------------------

            st.session_state.translated_paragraphs = (
                translated_items
            )

            st.session_state.translated_file = (
                output_path
            )

            progress.progress(100)

            status.success(
                "✅ Translation completed!"
            )

            st.success(
                "🎉 Your translated document "
                "is ready!"
            )

        except Exception as error:

            st.error(
                f"❌ Translation error: {error}"
            )


# ============================================================
# TRANSLATED DOCUMENT PREVIEW
# ============================================================

if st.session_state.translated_file:

    st.markdown(
        "## 👁️ Translated Document Preview"
    )

    st.caption(
        "Review the translated document "
        "before downloading."
    )

    try:

        preview_doc = Document(
            st.session_state.translated_file
        )

        preview_paragraphs = [
            paragraph.text
            for paragraph in
            preview_doc.paragraphs
            if paragraph.text.strip()
        ]

        if preview_paragraphs:

            for translated_index, paragraph_text in enumerate(
                preview_paragraphs
            ):

                st.text_area(
                    f"Translated Section "
                    f"{translated_index + 1}",
                    paragraph_text,
                    height=110,
                    key=(
                        f"translated_preview_"
                        f"{translated_index}"
                    ),
                    disabled=True
                )

        else:

            st.info(
                "The translated document "
                "contains no previewable text."
            )

    except Exception as error:

        st.warning(
            f"Preview could not be generated: "
            f"{error}"
        )


# ============================================================
# EXPORT
# ============================================================

if st.session_state.translated_file:

    st.markdown("## 📤 Export")

    st.success(
        "✅ Translation complete — "
        "your DOCX document is ready."
    )

    try:

        with open(
            st.session_state.translated_file,
            "rb"
        ) as file:

            document_bytes = file.read()

        st.download_button(
            label="⬇️ Download Translated DOCX",
            data=document_bytes,
            file_name="DocMorph_Translated.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
            key="download_translated_docx"
        )

    except Exception as error:

        st.error(
            f"❌ Download preparation failed: "
            f"{error}"
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "🌐 DocMorph • Smart Document Language Converter • "
    "Built with Streamlit"
)
