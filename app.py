import streamlit as st
import fitz
from docx import Document
from deep_translator import GoogleTranslator
from langdetect import detect
import tempfile
import os
from pathlib import Path


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DocMorph",
    page_icon="🌐",
    layout="wide"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 42px;
        font-weight: 700;
        text-align: center;
        margin-bottom: 5px;
    }

    .subtitle {
        text-align: center;
        font-size: 18px;
        margin-bottom: 30px;
    }

    .step-card {
        padding: 15px;
        border-radius: 12px;
        text-align: center;
        border: 1px solid #444;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# TITLE
# ============================================================

st.markdown(
    '<div class="main-title">🌐 DocMorph</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Smart Document Language Converter'
    '</div>',
    unsafe_allow_html=True
)


# ============================================================
# SESSION STATE
# ============================================================

if "extracted_pages" not in st.session_state:
    st.session_state.extracted_pages = []

if "source_language" not in st.session_state:
    st.session_state.source_language = ""

if "translated_file" not in st.session_state:
    st.session_state.translated_file = None

if "page_count" not in st.session_state:
    st.session_state.page_count = 0


# ============================================================
# WORKFLOW HEADER
# ============================================================

step1, step2, step3, step4 = st.columns(4)

with step1:
    st.info("01\n\n📥 INPUT")

with step2:
    st.info("02\n\n🔍 ANALYZE")

with step3:
    st.info("03\n\n🌐 TRANSLATE")

with step4:
    st.info("04\n\n📤 EXPORT")


# ============================================================
# UPLOAD SECTION
# ============================================================

st.subheader("📥 Document Input")

uploaded_file = st.file_uploader(
    "Upload your document",
    type=["pdf", "docx", "txt"],
    help="Supported formats: PDF, DOCX and TXT"
)


# ============================================================
# TARGET LANGUAGES
# ============================================================

language_map = {
    "English": "en",
    "Hindi": "hi",
    "Kannada": "kn",
    "Tamil": "ta",
    "Telugu": "te",
    "Malayalam": "ml",
    "Marathi": "mr",
    "Bengali": "bn",
    "Gujarati": "gu",
    "Punjabi": "pa",
    "Urdu": "ur",
    "French": "fr",
    "German": "de",
    "Spanish": "es",
    "Italian": "it",
    "Portuguese": "pt",
    "Russian": "ru",
    "Arabic": "ar",
    "Chinese": "zh-CN",
    "Japanese": "ja",
    "Korean": "ko"
}


# ============================================================
# DOCUMENT EXTRACTION
# ============================================================

def extract_pdf_pages(file_bytes):

    pages = []

    pdf = fitz.open(
        stream=file_bytes,
        filetype="pdf"
    )

    for page_number, page in enumerate(pdf):

        text = page.get_text("text")

        pages.append({
            "page": page_number + 1,
            "text": text.strip()
        })

    pdf.close()

    return pages


def extract_docx(file_bytes):

    temp_path = os.path.join(
        tempfile.gettempdir(),
        "input_document.docx"
    )

    with open(temp_path, "wb") as f:
        f.write(file_bytes)

    document = Document(temp_path)

    paragraphs = []

    for paragraph in document.paragraphs:

        text = paragraph.text.strip()

        if text:

            paragraphs.append({
                "page": 1,
                "text": text
            })

    os.remove(temp_path)

    return paragraphs


def extract_txt(file_bytes):

    text = file_bytes.decode(
        "utf-8",
        errors="ignore"
    )

    return [{
        "page": 1,
        "text": text
    }]


# ============================================================
# TEXT CHUNKING
# ============================================================

def create_chunks(text, max_chars=3000):

    """
    Split large text into manageable chunks.

    This prevents sending extremely large text
    in a single translation request.
    """

    text = text.strip()

    if not text:
        return []

    chunks = []

    start = 0

    while start < len(text):

        end = start + max_chars

        # If this is not the last chunk,
        # try to break at a space.
        if end < len(text):

            space_position = text.rfind(
                " ",
                start,
                end
            )

            if space_position > start:

                end = space_position

        chunk = text[start:end].strip()

        if chunk:
            chunks.append(chunk)

        start = end

    return chunks


# ============================================================
# BATCH TRANSLATION
# ============================================================

def translate_text(text, target_language):

    if not text.strip():
        return ""

    # Create larger chunks instead of sending
    # many small translation requests.
    chunks = create_chunks(
        text,
        max_chars=3000
    )

    if not chunks:
        return ""

    translator = GoogleTranslator(
        source="auto",
        target=target_language
    )

    translated_chunks = []

    # --------------------------------------------------------
    # Translate in batches
    # --------------------------------------------------------

    batch_size = 5

    for i in range(
        0,
        len(chunks),
        batch_size
    ):

        batch = chunks[
            i:i + batch_size
        ]

        translated_batch = translator.translate_batch(
            batch
        )

        translated_chunks.extend(
            translated_batch
        )

    return "\n\n".join(
        translated_chunks
    )


# ============================================================
# DOCX CREATION
# ============================================================

def create_translated_docx(
    translated_pages,
    source_language,
    target_language,
    original_name
):

    output_path = os.path.join(
        tempfile.gettempdir(),
        "DocMorph_Translated.docx"
    )

    document = Document()

    # --------------------------------------------------------
    # Title
    # --------------------------------------------------------

    document.add_heading(
        "DocMorph - Translated Document",
        level=0
    )

    document.add_paragraph(
        f"Original Document: {original_name}"
    )

    document.add_paragraph(
        f"Source Language: {source_language}"
    )

    document.add_paragraph(
        f"Target Language: {target_language}"
    )

    document.add_paragraph(
        "Generated using DocMorph"
    )

    document.add_page_break()

    # --------------------------------------------------------
    # Translated pages
    # --------------------------------------------------------

    current_page = None

    for item in translated_pages:

        page_number = item["page"]

        if page_number != current_page:

            if current_page is not None:

                document.add_page_break()

            document.add_heading(
                f"Page {page_number}",
                level=1
            )

            current_page = page_number

        document.add_paragraph(
            item["translated"]
        )

    document.save(output_path)

    return output_path


# ============================================================
# ANALYZE BUTTON
# ============================================================

if uploaded_file is not None:

    st.success(
        f"📄 {uploaded_file.name} uploaded"
    )

    col_a, col_b = st.columns(2)

    with col_a:

        if st.button(
            "🔍 Analyze Document",
            use_container_width=True
        ):

            file_bytes = uploaded_file.getvalue()

            extension = Path(
                uploaded_file.name
            ).suffix.lower()

            try:

                # PDF
                if extension == ".pdf":

                    pages = extract_pdf_pages(
                        file_bytes
                    )

                # DOCX
                elif extension == ".docx":

                    pages = extract_docx(
                        file_bytes
                    )

                # TXT
                elif extension == ".txt":

                    pages = extract_txt(
                        file_bytes
                    )

                else:

                    st.error(
                        "Unsupported file format."
                    )

                    st.stop()

                st.session_state.extracted_pages = pages

                st.session_state.page_count = len(
                    pages
                )

                # ------------------------------------------------
                # Detect language
                # ------------------------------------------------

                combined_text = " ".join(
                    item["text"]
                    for item in pages
                    if item["text"]
                )

                if combined_text:

                    try:

                        detected = detect(
                            combined_text[:3000]
                        )

                        st.session_state.source_language = (
                            detected
                        )

                    except Exception:

                        st.session_state.source_language = (
                            "Unknown"
                        )

                st.success(
                    "✅ Document analyzed successfully!"
                )

            except Exception as e:

                st.error(
                    f"❌ Analysis error: {str(e)}"
                )


# ============================================================
# DOCUMENT INFORMATION
# ============================================================

if st.session_state.extracted_pages:

    st.subheader(
        "🔍 Document Intelligence"
    )

    info1, info2, info3 = st.columns(3)

    with info1:

        st.metric(
            "Pages / Sections",
            st.session_state.page_count
        )

    with info2:

        st.metric(
            "Detected Language",
            st.session_state.source_language
        )

    with info3:

        total_chars = sum(
            len(item["text"])
            for item in st.session_state.extracted_pages
        )

        st.metric(
            "Characters",
            total_chars
        )


# ============================================================
# TARGET LANGUAGE
# ============================================================

if st.session_state.extracted_pages:

    st.subheader(
        "🌐 Translation Settings"
    )

    target_language_name = st.selectbox(
        "Choose Target Language",
        list(language_map.keys())
    )

    target_language = language_map[
        target_language_name
    ]

    # ========================================================
    # PREVIEW
    # ========================================================

    with st.expander(
        "👁 Preview Extracted Content"
    ):

        for item in (
            st.session_state.extracted_pages
        ):

            st.markdown(
                f"**Page / Section {item['page']}**"
            )

            preview = item["text"]

            if len(preview) > 500:

                preview = (
                    preview[:500]
                    + "..."
                )

            st.write(preview)

    # ========================================================
    # TRANSLATE BUTTON
    # ========================================================

    if st.button(
        "🌐 Translate Document",
        type="primary",
        use_container_width=True
    ):

        progress = st.progress(0)

        status = st.empty()

        translated_pages = []

        total = len(
            st.session_state.extracted_pages
        )

        try:

            for index, item in enumerate(
                st.session_state.extracted_pages
            ):

                status.write(
                    f"🌐 Translating page / section "
                    f"{index + 1} of {total}..."
                )

                translated = translate_text(
                    item["text"],
                    target_language
                )

                translated_pages.append({
                    "page": item["page"],
                    "translated": translated
                })

                progress.progress(
                    int(
                        ((index + 1) / total) * 100
                    )
                )

            # ------------------------------------------------
            # Create DOCX
            # ------------------------------------------------

            status.write(
                "📄 Creating translated DOCX..."
            )

            output_path = create_translated_docx(
                translated_pages,
                st.session_state.source_language,
                target_language_name,
                uploaded_file.name
            )

            st.session_state.translated_file = (
                output_path
            )

            status.write(
                "✅ Translation completed!"
            )

            st.success(
                "🎉 Your translated document is ready!"
            )

        except Exception as e:

            st.error(
                f"❌ Translation error: {str(e)}"
            )


# ============================================================
# DOWNLOAD
# ============================================================

if st.session_state.translated_file:

    st.subheader("📤 Export")

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
        use_container_width=True
    )
