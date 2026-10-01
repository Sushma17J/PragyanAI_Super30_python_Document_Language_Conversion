import streamlit as st
import fitz
from docx import Document
import requests
from langdetect import detect
import speech_recognition as sr
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
    "last_uploaded_file": "",

    # Voice language selection
    "voice_language": "",
    "voice_language_confirmed": False,
    "voice_language_text": ""
}

for key, value in DEFAULTS.items():

    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    .main {
        padding-top: 1rem;
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
# LANGUAGE MAP
# ============================================================

language_map = {

    "English": "en",
    "Hindi": "hi",
    "Kannada": "kn",
    "Telugu": "te",
    "Tamil": "ta",
    "Marathi": "mr",
    "Bengali": "bn",
    "Gujarati": "gu",
    "Punjabi": "pa",
    "Malayalam": "ml",
    "Urdu": "ur",
    "Assamese": "as",
    "Nepali": "ne",
    "Sanskrit": "sa",

    "French": "fr",
    "German": "de",
    "Spanish": "es",
    "Italian": "it",
    "Portuguese": "pt",
    "Russian": "ru",
    "Arabic": "ar",
    "Japanese": "ja",
    "Korean": "ko",
    "Chinese": "zh-CN",
    "Turkish": "tr",
    "Dutch": "nl",
    "Greek": "el",
    "Hebrew": "he",
    "Polish": "pl",
    "Ukrainian": "uk",
    "Vietnamese": "vi",
    "Indonesian": "id",
    "Malay": "ms",
    "Thai": "th",
    "Swedish": "sv",
    "Danish": "da",
    "Finnish": "fi",
    "Norwegian": "no",
    "Czech": "cs",
    "Romanian": "ro",
    "Hungarian": "hu"
}


# ============================================================
# SOURCE LANGUAGE NAMES
# ============================================================

source_names = {

    "en": "English",
    "hi": "Hindi",
    "kn": "Kannada",
    "te": "Telugu",
    "ta": "Tamil",
    "mr": "Marathi",
    "bn": "Bengali",
    "gu": "Gujarati",
    "pa": "Punjabi",
    "ml": "Malayalam",
    "ur": "Urdu",
    "as": "Assamese",
    "ne": "Nepali",
    "sa": "Sanskrit",

    "fr": "French",
    "de": "German",
    "es": "Spanish",
    "it": "Italian",
    "pt": "Portuguese",
    "ru": "Russian",
    "ar": "Arabic",
    "ja": "Japanese",
    "ko": "Korean",

    "zh-cn": "Chinese",
    "zh-CN": "Chinese",
    "zh-tw": "Chinese",

    "tr": "Turkish",
    "nl": "Dutch",
    "el": "Greek",
    "he": "Hebrew",
    "pl": "Polish",
    "uk": "Ukrainian",
    "vi": "Vietnamese",
    "id": "Indonesian",
    "ms": "Malay",
    "th": "Thai",
    "sv": "Swedish",
    "da": "Danish",
    "fi": "Finnish",
    "no": "Norwegian",
    "cs": "Czech",
    "ro": "Romanian",
    "hu": "Hungarian"
}


# ============================================================
# VOICE LANGUAGE ALIASES
# ============================================================

voice_language_aliases = {

    "english": "English",
    "hindi": "Hindi",
    "kannada": "Kannada",
    "telugu": "Telugu",
    "tamil": "Tamil",
    "marathi": "Marathi",
    "bengali": "Bengali",
    "gujarati": "Gujarati",
    "punjabi": "Punjabi",
    "malayalam": "Malayalam",
    "urdu": "Urdu",
    "assamese": "Assamese",
    "nepali": "Nepali",
    "sanskrit": "Sanskrit",

    "french": "French",
    "german": "German",
    "spanish": "Spanish",
    "italian": "Italian",
    "portuguese": "Portuguese",
    "russian": "Russian",
    "arabic": "Arabic",
    "japanese": "Japanese",
    "korean": "Korean",
    "chinese": "Chinese",
    "turkish": "Turkish",
    "dutch": "Dutch",
    "greek": "Greek",
    "hebrew": "Hebrew",
    "polish": "Polish",
    "ukrainian": "Ukrainian",
    "vietnamese": "Vietnamese",
    "indonesian": "Indonesian",
    "malay": "Malay",
    "thai": "Thai",
    "swedish": "Swedish",
    "danish": "Danish",
    "finnish": "Finnish",
    "norwegian": "Norwegian",
    "czech": "Czech",
    "romanian": "Romanian",
    "hungarian": "Hungarian"
}


# ============================================================
# VOICE RECOGNITION
# ============================================================

def recognize_language_from_voice(audio_file):

    recognizer = sr.Recognizer()

    temp_audio = os.path.join(
        tempfile.gettempdir(),
        "docmorph_voice.wav"
    )

    try:

        with open(
            temp_audio,
            "wb"
        ) as file:

            file.write(
                audio_file.getvalue()
            )

        with sr.AudioFile(
            temp_audio
        ) as source:

            audio_data = recognizer.record(
                source
            )

        spoken_text = recognizer.recognize_google(
            audio_data,
            language="en-IN"
        )

        return spoken_text.strip()

    except sr.UnknownValueError:

        return ""

    except sr.RequestError as error:

        raise Exception(
            f"Speech recognition service error: {error}"
        )

    finally:

        try:
            os.remove(temp_audio)

        except OSError:
            pass


# ============================================================
# FIND LANGUAGE FROM SPEECH
# ============================================================

def find_language_from_speech(
    spoken_text
):

    cleaned_text = (
        spoken_text
        .lower()
        .strip()
        .replace(".", "")
        .replace(",", "")
    )

    if cleaned_text in voice_language_aliases:

        return voice_language_aliases[
            cleaned_text
        ]

    for alias, language in (
        voice_language_aliases.items()
    ):

        if alias in cleaned_text:

            return language

    return None


# ============================================================
# HERO
# ============================================================

st.title("🌐 DocMorph")

st.subheader(
    "✨ AI-Assisted Document Workspace"
)

st.write(
    "Smart Document Language Converter — "
    "upload a PDF, DOCX, or TXT file, detect "
    "its language, translate it into your "
    "selected language, preview the result, "
    "and export it as a DOCX document."
)


# ============================================================
# WORKFLOW
# ============================================================

st.markdown("### 🔄 How DocMorph Works")

workflow_col1, workflow_col2, workflow_col3, workflow_col4 = (
    st.columns(4)
)

with workflow_col1:

    st.markdown("### 01")
    st.markdown("## 📥")
    st.markdown("**INPUT**")

with workflow_col2:

    st.markdown("### 02")
    st.markdown("## 🔍")
    st.markdown("**ANALYZE**")

with workflow_col3:

    st.markdown("### 03")
    st.markdown("## 🌐")
    st.markdown("**TRANSLATE**")

with workflow_col4:

    st.markdown("### 04")
    st.markdown("## 👁️")
    st.markdown("**PREVIEW & EXPORT**")

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

    # --------------------------------------------------------
    # SIMPLE DOCX UPLOAD
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # SUPPORTED FILES
    # --------------------------------------------------------

    st.markdown("### 📄 Supported Files")

    st.write("• PDF")
    st.write("• DOCX")
    st.write("• TXT")

    # --------------------------------------------------------
    # TARGET LANGUAGES
    # --------------------------------------------------------

    st.markdown("### 🌐 Target Languages")

    for language in language_map.keys():

        st.write(
            f"• {language}"
        )

    st.divider()

    st.caption(
        "Translation powered by the MyMemory API."
    )


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf_pages(
    file_bytes
):

    pages = []

    pdf = fitz.open(
        stream=file_bytes,
        filetype="pdf"
    )

    for page_number, page in enumerate(
        pdf
    ):

        text = page.get_text(
            "text"
        ).strip()

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

def extract_docx(
    file_bytes
):

    temp_path = os.path.join(
        tempfile.gettempdir(),
        "docmorph_input.docx"
    )

    with open(
        temp_path,
        "wb"
    ) as file:

        file.write(
            file_bytes
        )

    document = Document(
        temp_path
    )

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

def extract_txt(
    file_bytes
):

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

            chunks.append(
                chunk
            )

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

    translated_text = (
        response_data.get(
            "translatedText"
        )
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

        if (
            chunk_index
            <
            len(chunks) - 1
        ):

            time.sleep(0.4)

    return "\n\n".join(
        translated_chunks
    )


# ============================================================
# CREATE DOCX PRESERVING ORIGINAL STRUCTURE
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

    with open(
        temp_input,
        "wb"
    ) as file:

        file.write(
            original_bytes
        )

    document = Document(
        temp_input
    )

    translated_map = {

        item["index"]:
        item["translated"]

        for item in translated_paragraphs

    }

    for index, paragraph in enumerate(
        document.paragraphs
    ):

        if index not in translated_map:

            continue

        translated_text = (
            translated_map[index]
        )

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

    document.save(
        output_path
    )

    try:

        os.remove(
            temp_input
        )

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

    document.save(
        output_path
    )

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

    current_file_name = (
        uploaded_file.name
    )

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

        st.session_state.voice_language = ""

        st.session_state.voice_language_confirmed = False

        st.session_state.voice_language_text = ""

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

        file_bytes = (
            uploaded_file.getvalue()
        )

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
    # TARGET LANGUAGE
    # ========================================================

    st.markdown(
        "### 🌐 Choose Target Language"
    )

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
            "🌐 Select Target Language",
            list(language_map.keys()),
            key="target_language_selector"
        )


    # ========================================================
    # VOICE TARGET LANGUAGE
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### 🎤 Select Target Language by Voice"
    )

    st.caption(
        "Speak the language name, for example: "
        "\"Kannada\", \"Hindi\", \"French\" or \"German\"."
    )

    audio_value = st.audio_input(
        "🎤 Speak Target Language",
        key="language_voice_input"
    )


    # ========================================================
    # PROCESS VOICE
    # ========================================================

    if audio_value is not None:

        try:

            spoken_text = (
                recognize_language_from_voice(
                    audio_value
                )
            )

            if not spoken_text:

                st.warning(
                    "⚠️ I could not understand "
                    "your voice. Please speak "
                    "the language name clearly."
                )

            else:

                st.session_state.voice_language_text = (
                    spoken_text
                )

                detected_language_name = (
                    find_language_from_speech(
                        spoken_text
                    )
                )

                if detected_language_name:

                    st.session_state.voice_language = (
                        detected_language_name
                    )

                    st.session_state.voice_language_confirmed = (
                        False
                    )

                    st.success(
                        f"🎤 You said: "
                        f"**{spoken_text}**"
                    )

                else:

                    st.warning(
                        f"🎤 You said: "
                        f"**{spoken_text}**"
                    )

                    st.error(
                        "❌ I could not match "
                        "that to a supported "
                        "target language."
                    )

        except Exception as error:

            st.error(
                f"❌ Voice recognition error: "
                f"{error}"
            )


    # ========================================================
    # CONFIRM VOICE LANGUAGE
    # ========================================================

    if st.session_state.voice_language:

        st.markdown(
            "#### 🔎 Confirm Your Language"
        )

        st.info(
            f"🎤 You selected: "
            f"**{st.session_state.voice_language}**\n\n"
            f"Is this correct?"
        )

        yes_col, no_col = st.columns(2)

        with yes_col:

            if st.button(
                "✅ Yes, Continue",
                use_container_width=True,
                key="voice_language_yes"
            ):

                st.session_state.voice_language_confirmed = (
                    True
                )

                st.session_state.voice_language_text = (
                    st.session_state.voice_language
                )

                st.rerun()

        with no_col:

            if st.button(
                "❌ No, Speak Again",
                use_container_width=True,
                key="voice_language_no"
            ):

                st.session_state.voice_language = ""

                st.session_state.voice_language_confirmed = (
                    False
                )

                st.session_state.voice_language_text = ""

                st.rerun()


    # ========================================================
    # APPLY CONFIRMED VOICE LANGUAGE
    # ========================================================

    if (
        st.session_state.voice_language
        and
        st.session_state.voice_language_confirmed
    ):

        target_language_name = (
            st.session_state.voice_language
        )

        st.success(
            f"🎯 Target language confirmed: "
            f"**{target_language_name}**"
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

            for (
                translated_index,
                paragraph_text
            ) in enumerate(
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

    st.markdown(
        "## 📤 Export"
    )

    st.success(
        "✅ Translation complete — "
        "your DOCX document is ready."
    )

    try:

        with open(
            st.session_state.translated_file,
            "rb"
        ) as file:

            document_bytes = (
                file.read()
            )

        st.download_button(
            label="⬇️ Download Translated DOCX",
            data=document_bytes,
            file_name=(
                "DocMorph_Translated.docx"
            ),
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
