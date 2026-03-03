# Study Guide Generator (Streamlit)

Upload one or more **PDF** chapter files and generate a **DOCX study guide** using the bundled Word template.

## Files
- `app.py` — Streamlit UI
- `generate_study_guide_core.py` — generation logic (PDF → prompt → DOCX)
- `requirements.txt`
- `Dockerfile`

## Template
Place `Study Guide template.docx` in the same folder as `app.py`.

## Environment variables
- `OPENAI_API_KEY` (required)
- `OPENAI_MODEL` (optional, default: `gpt-4.1-mini`)

### Local (.env)
Create a file named `.env` next to `app.py`:

```
OPENAI_API_KEY=your_key_here
# OPENAI_MODEL=gpt-4.1-mini
```

## Run locally

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Push this folder to a Git repo.
2. In Streamlit Cloud, create a new app pointing at `app.py`.
3. Add Secrets:

```toml
OPENAI_API_KEY = "..."
# OPENAI_MODEL = "gpt-4.1-mini"
```

## Docker

```bash
docker build -t study-guide-generator .
docker run -p 8501:8501 -e OPENAI_API_KEY=... study-guide-generator
```

Then open: `http://localhost:8501`

## Notes about PDFs
- Best results come from **text-searchable** PDFs.
- If a PDF is scanned images, OCR it first (otherwise extraction may be near-empty).
