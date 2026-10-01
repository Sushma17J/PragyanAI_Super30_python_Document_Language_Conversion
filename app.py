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

if "extracted_pages" not in st.session_state:
    st.session_state.extracted_pages = []

if "source_language" not in st.session_state:
    st.session_state.source_language = ""

if "translated_file" not in st.session_state:
    st.session_state.translated_file = None

if "translated_paragraphs" not in st.session_state:
    st.session_state.translated_paragraphs = []

if "page_count" not in st.session_state:
    st.session_state.page_count = 0

if "input_extension" not in st.session_state:
    st.session_state.input_extension = ""

if "original_file_bytes" not in st.session_state:
    st.session_state.original_file_bytes = None

if "original_file_name" not in st.session_state:
    st.session_state.original_file_name = ""


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    /* Main page */

    .main {
        padding-top: 1rem;
    }

    /* Hero */

    .hero-box {
        padding: 30px;
        border-radius: 24px;
        margin-bottom: 25px;
        border: 1px solid rgba(255,255,255,0.12);
        background:
            linear-gradient(
                135deg,
                rgba(80,70,180,0.22),
                rgba(0,170,255,0.12)
            );
    }

    .hero-badge {
        display: inline-block;
        padding: 7px 14px;
        border-radius: 30px;
        background: rgba(120,100,255,0.18);
        border: 1px solid rgba(140,120,255,0.35);
        color: #c9c2ff;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 1px;
        margin-bottom: 12px;
    }

    .hero-title {
        font-size: 48px;
        font-weight: 800;
        margin: 0;
        letter-spacing: -1px;
    }

    .hero-subtitle {
        color: #b9bfd4;
        font-size: 17px;
        line-height: 1.6;
        margin-top: 10px;
        max-width: 900px;
    }

    /* Workflow cards */

    .step-card {
        padding: 18px 12px;
        border-radius: 18px;
        border: 1px solid rgba(255,255,255,0.10);
        background: rgba(255,255,255,0.045);
        text-align: center;
        min-height: 105px;
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

    /* Preview boxes */

    .preview-box {
        padding: 18px;
        border-radius: 16px;
        border: 1px solid rgba(255,255,255,0.10);
        background: rgba(0,0,0,0.18);
        margin-bottom: 12px;
    }

    /* Success box */

    .success-box {
        padding: 18px;
        border-radius: 18px;
        border: 1px solid rgba(0,255,170,0.25);
        background: rgba(0,255,170,0.06);
    }

    /* Metrics */

    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.045);
        border: 1px solid rgba(255,255,255,0.08);
        padding: 14px;
        border-radius: 16px;
    }

    /* Buttons */

    button[kind="primary"] {
        border-radius: 14px !important;
        font-weight: 700 !important;
    }

    /* File uploader */

    section[data-testid="stFileUploaderDropzone"] {
        border-radius: 18px;
        border: 1px dashed rgba(130,120,255,0.45);
    }

    /* Footer */

    .footer-text {
        text-align: center;
        color: #777e99;
        padding-top: 30px;
        padding-bottom: 15px;
        font-size: 13px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HERO
# ============================================================

st.markdown(
    """
    <div class="hero-box">
        <div class="hero-badge">
            ✨ AI-ASSISTED DOCUMENT WORKSPACE
        </div>

        <div class="hero-title">
            🌐 DocMorph
        </div>

        <div class="hero-subtitle">
            Smart Document Language Converter — upload a PDF, DOCX, or TXT
            file, detect its language, translate it into your selected
            language, preview the result, and export it as a DOCX document.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# WORKFLOW
# ============================================================

st.markdown("### 🔄 How DocMorph Works")

steps = [
    ("01", "📥", "INPUT"),
    ("02", "🔍", "ANALYZE"),
    ("03", "🌐", "TRANSLATE"),
    ("04", "👁️", "PREVIEW & EXPORT"),
]

cols = st.columns(4)

for col, (number, icon, title) in zip(cols, steps):
    with col:
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
    st.caption("Smart document translation workspace")

    st.divider()

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

    st.caption("Translation powered by the MyMemory API.")


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
# DOCUMENT EXTRACTION
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


def extract_docx(file_bytes):

    temp_path = os.path.join(
        tempfile.gettempdir(),
        "docmorph_input.docx"
    )

    with open(temp_path, "wb") as f:
        f.write(file_bytes)

    document = Document(temp_path)

    paragraphs = []

    for index, paragraph in enumerate(document.paragraphs):

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

def create_chunks(text, max_chars=450):

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

        chunk = text[start:end].strip()

        if chunk:
            chunks.append(chunk)

        start = end

    return chunks


# ============================================================
# TRANSLATION
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

    url = "https://api.mymemory.translated.net/get"

    params = {
        "q": text,
        "langpair": f"{source_language}|{target_language}"
    }

    response = requests.get(
        url,
        params=params,
        timeout=30
    )

    if response.status_code != 200:

        raise Exception(
            f"Translation server returned status "
            f"{response.status_code}"
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
            "Translation service did not return translated text."
        )

    return translated_text


def translate_text(
    text,
    source_language,
    target_language
):

    if not text or not text.strip():
        return ""

    chunks = create_chunks(
        text,
        max_chars=450
    )

    translated_chunks = []

    for index, chunk in enumerate(chunks):

        translated = translate_chunk(
            chunk,
            source_language,
            target_language
        )

        translated_chunks.append(translated)

        if index < len(chunks) - 1:
            time.sleep(0.4)

    return "\n\n".join(
        translated_chunks
    )


# ============================================================
# CREATE DOCX WHILE PRESERVING ORIGINAL DOCX STRUCTURE
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

    with open(temp_input, "wb") as f:
        f.write(original_bytes)

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

        translated_text = translated_map[index]

        if paragraph.runs:

            paragraph.runs[0].text = translated_text

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

        if translated.strip():

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
    type=["pdf", "docx", "txt"],
    help="Supported formats: PDF, DOCX and TXT"
)


# ============================================================
# ANALYZE DOCUMENT
# ============================================================

if uploaded_file is not None:

    # IMPORTANT:
    # No filename card / HTML filename display here.
    st.success(
        "✅ Document uploaded successfully!"
    )

    if st.button(
        "🔍 Analyze Document",
        use_container_width=True
    ):

        file_bytes = uploaded_file.getvalue()

        extension = Path(
            uploaded_file.name
        ).suffix.lower()

        try:

            # ----------------------------
            # Extract document
            # ----------------------------

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

                st.error(
                    "Unsupported file format."
                )

                st.stop()

            # ----------------------------
            # Remove empty content
            # ----------------------------

            pages = [
                page
                for page in pages
                if page["text"].strip()
            ]

            if not pages:

                raise Exception(
                    "No readable text was found in the document."
                )

            # ----------------------------
            # Save session data
            # ----------------------------

            st.session_state.extracted_pages = pages

            st.session_state.page_count = len(
                pages
            )

            st.session_state.input_extension = extension

            st.session_state.original_file_bytes = (
                file_bytes
            )

            st.session_state.original_file_name = (
                uploaded_file.name
            )

            # Reset old translation

            st.session_state.translated_file = None

            st.session_state.translated_paragraphs = []

            # ----------------------------
            # Detect language
            # ----------------------------

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

                st.session_state.source_language = "en"

            st.success(
                "✅ Document analyzed successfully!"
            )

        except Exception as e:

            st.error(
                f"❌ Analysis error: {str(e)}"
            )


# ============================================================
# DOCUMENT INTELLIGENCE
# ============================================================

if st.session_state.extracted_pages:

    st.markdown("## 🔍 Document Intelligence")

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
            for item in st.session_state.extracted_pages
        )

        st.metric(
            "Characters",
            total_chars
        )

    st.write("")

    # ========================================================
    # LANGUAGE SELECTION
    # ========================================================

    col1, col2 = st.columns(2)

    with col1:

        st.info(
            f"🔎 Source Language: **{source_display}**"
        )

    with col2:

        target_language_name = st.selectbox(
            "🌐 Choose Target Language",
            list(language_map.keys())
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

        for item in st.session_state.extracted_pages:

            label = item.get(
                "page",
                item.get("index", 1)
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

            # NORMAL STREAMLIT TEXT
            # No HTML tags will appear.

            st.text_area(
                "Extracted text",
                preview,
                height=150,
                key=f"original_preview_{label}",
                disabled=True
            )


    # ========================================================
    # TRANSLATE
    # ========================================================

    if st.button(
        "🌐 Translate Document",
        type="primary",
        use_container_width=True
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
                    "Source language could not be detected."
                )

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
                        ((index + 1) / total)
                        * 80
                    )
                )

            status.write(
                "📄 Creating translated DOCX..."
            )

            # ------------------------------------------------
            # DOCX
            # Preserve original paragraph structure.
            # ------------------------------------------------

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

            # ------------------------------------------------
            # PDF / TXT
            # Create one continuous DOCX.
            # ------------------------------------------------

            else:

                output_path = (
                    create_simple_translated_docx(
                        translated_items
                    )
                )

            st.session_state.translated_paragraphs = (
                translated_items
            )

            st.session_state.translated_file = (
                output_path
            )

            progress.progress(100)

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
# TRANSLATED DOCUMENT PREVIEW
# ============================================================

if st.session_state.translated_file:

    st.markdown("## 👁️ Translated Document Preview")

    st.caption(
        "Review the translated content before downloading your DOCX."
    )

    try:

        preview_doc = Document(
            st.session_state.translated_file
        )

        preview_paragraphs = [
            paragraph.text
            for paragraph in preview_doc.paragraphs
            if paragraph.text.strip()
        ]

        if preview_paragraphs:

            for index, paragraph_text in enumerate(
                preview_paragraphs
            ):

                # Avoid showing the generated metadata
                # heading in a confusing way.

                st.text_area(
                    f"Translated Section {index + 1}",
                    paragraph_text,
                    height=110,
                    key=f"translated_preview_{index}",
                    disabled=True
                )

        else:

            st.info(
                "The translated document contains no previewable paragraphs."
            )

    except Exception as e:

        st.warning(
            f"Preview could not be generated: {str(e)}"
        )


# ============================================================
# EXPORT
# ============================================================

if st.session_state.translated_file:

    st.markdown("## 📤 Export")

    st.success(
        "✅ Translation complete — your DOCX document is ready."
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
            use_container_width=True
        )

    except Exception as e:

        st.error(
            f"❌ Download preparation failed: {str(e)}"
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "🌐 DocMorph • Smart Document Language Converter • Built with Streamlit"
)
