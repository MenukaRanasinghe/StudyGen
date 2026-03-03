import os
import tempfile
from pathlib import Path

import streamlit as st

from generate_study_guide_core import (
    DEFAULT_TEMPLATE_FILENAME,
    run_generation_from_pdfs,
)


st.set_page_config(
    page_title="Study Guide Generator",
    page_icon="📘",
    layout="centered",
)

st.title("📘 Study Guide Generator")
st.caption("Upload PDF chapters → generate a study guide DOCX using the bundled template.")

# ---- Inputs ----
with st.form("inputs"):
    pdf_files = st.file_uploader(
        "PDF chapters (one or more)",
        type=["pdf"],
        accept_multiple_files=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        course_name = st.text_input("Course name", placeholder="e.g., Level 5 Diploma in ...")
    with col2:
        unit_no = st.text_input("Unit no", placeholder="e.g., 10")

    base_name = st.text_input("Output file name", value="StudyGuide")

    cover_image = st.file_uploader(
        "Cover image (optional, PNG/JPG)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=False,
    )


    with st.expander("Advanced (optional)"):
        word_limit_mode = st.selectbox("Word limit", options=["auto", "750", "1000"], index=0)
        auto_threshold = st.number_input("Auto threshold (PDF count)", min_value=1, max_value=10, value=3, step=1)
        max_source_chars = st.number_input("Max source chars", min_value=20000, max_value=300000, value=120000, step=5000)
        retry_on_invalid = st.checkbox("Retry if JSON structure is invalid", value=True)
        retry_on_overlimit = st.checkbox("Retry if over word limit", value=True)


    submitted = st.form_submit_button("Generate DOCX")


# ---- Validation / Run ----
if submitted:
    if not pdf_files:
        st.error("Please upload at least one PDF.")
        st.stop()

    if not course_name.strip():
        st.error("Please enter a Course name.")
        st.stop()

    if not unit_no.strip().isdigit():
        st.error("Please enter a numeric Unit no (e.g., 10).")
        st.stop()

    # Template must be present next to app files
    template_path = Path(__file__).with_name(DEFAULT_TEMPLATE_FILENAME)
    if not template_path.exists():
        st.error(
            f"Template file not found: {template_path}\n\n"
            f"Place '{DEFAULT_TEMPLATE_FILENAME}' in the same folder as app.py."
        )
        st.stop()

    if not os.environ.get("OPENAI_API_KEY"):
        st.warning(
            "OPENAI_API_KEY is not set.\n\n"
            "• Local: set it in your environment or add a .env file next to app.py\n"
            "• Streamlit Cloud: add it in App → Settings → Secrets"
        )

    safe_base = "".join([c for c in (base_name or "StudyGuide") if c.isalnum() or c in ("-", "_", " ")]).strip()
    if not safe_base:
        safe_base = "StudyGuide"

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)

        # Save uploaded PDFs
        pdf_paths = []
        for f in pdf_files:
            p = td_path / f.name
            p.write_bytes(f.getbuffer())
            pdf_paths.append(p)

        # Save cover image if provided
        cover_path = None
        if cover_image is not None:
            cover_path = td_path / cover_image.name
            cover_path.write_bytes(cover_image.getbuffer())

        out_docx = td_path / f"{safe_base}.docx"

        with st.spinner("Generating study guide…"):
            try:
                info = run_generation_from_pdfs(
                    pdf_inputs=pdf_paths,
                    out_docx=out_docx,
                    course_name=course_name,
                    unit_no=unit_no,
                    template_path=template_path,
                    cover_image=cover_path,
                    word_limit_mode=word_limit_mode,
                    auto_threshold=int(auto_threshold),
                    max_source_chars=int(max_source_chars),
                    retry_on_invalid=bool(retry_on_invalid),
                    retry_on_overlimit=bool(retry_on_overlimit),
                )
            except SystemExit as e:
                st.error(str(e))
                st.stop()
            except Exception as e:
                st.exception(e)
                st.stop()

        docx_bytes = out_docx.read_bytes()

        st.success("Done!")
        st.write(
            f"Processed **{info.get('pdf_count')}** PDF(s). Word limit: **{info.get('word_limit')}**. "
            f"Estimated output words: **{info.get('estimated_word_count')}**."
        )

        st.download_button(
            label="Download DOCX",
            data=docx_bytes,
            file_name=f"{safe_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


st.divider()

with st.expander("Deployment notes"):
    st.markdown(
        """
**Template**
- Put `Study Guide template.docx` in the same folder as `app.py`.

**Secrets / env vars**
- `OPENAI_API_KEY` (required)
- `OPENAI_MODEL` (optional, defaults to `gpt-4.1-mini`)

**PDF quality**
- This works best with *text-searchable* PDFs.
- If your PDFs are scanned images, OCR them first.
"""
    )
