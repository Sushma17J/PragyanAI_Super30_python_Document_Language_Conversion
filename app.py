# PREVIEW TRANSLATED DOCUMENT
# ============================================================

if st.session_state.translated_file:

    st.markdown(
        '<div class="section-title">👁️ Translated Document Preview</div>',
        unsafe_allow_html=True
    )

    st.caption(
        "Review the translated document before downloading."
    )

    preview_doc = Document(
        st.session_state.translated_file
    )

    st.markdown(
        '<div class="preview-box">',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="preview-label">DOCUMENT PREVIEW</div>',
        unsafe_allow_html=True
    )

    preview_has_content = False

    for paragraph in preview_doc.paragraphs:

        if paragraph.text.strip():

            preview_has_content = True

            st.markdown(
                f"<p style='color:#e8edf7;"
                f"font-size:16px;line-height:1.7;"
                f"margin-bottom:14px;'>"
                f"{paragraph.text}"
                f"</p>",
                unsafe_allow_html=True
            )

    if not preview_has_content:

        st.info(
            "No text available for preview."
        )

    st.markdown(
        '</div>',
        unsafe_allow_html=True
    )


# ============================================================
# EXPORT
# ============================================================

if st.session_state.translated_file:

    st.markdown(
        '<div class="section-title">📤 Export</div>',
        unsafe_allow_html=True
    )

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

    st.caption(
        "✓ Preview checked  •  ✓ DOCX ready  •  ✓ Download available"
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div style="
        text-align:center;
        margin-top:55px;
        padding-top:20px;
        border-top:1px solid rgba(255,255,255,.08);
        color:#75829a;
        font-size:13px;
    ">
        🌐 DocMorph &nbsp;•&nbsp;
        Smart Document Translation Workspace
    </div>
    """,
    unsafe_allow_html=True
)
