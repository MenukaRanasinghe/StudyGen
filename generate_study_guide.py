#!/usr/bin/env python3
"""
generate_study_guide_v2.py (clean version)

- Desktop GUI (Tkinter) to generate a Study Guide from a folder of PDFs
- Outputs to a Word template (DOCX) and optional PDF
- Uses OpenAI (ChatGPT via API) only
- Hides Provider/Model in the UI (fixed defaults)

Requirements:
  pip install pdfplumber python-docx openai
Optional for PDF export via Word:
  pip install docx2pdf
Or install LibreOffice for soffice conversion.

API Key:
  setx OPENAI_API_KEY "your_api_key_here"   (Windows PowerShell/CMD)
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pdfplumber
from docx import Document
from docx.shared import Pt


# ----------------------------
# Config (hidden from UI)
# ----------------------------

DEFAULT_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4.1-mini")  # change if needed (UI will not show it)


# ----------------------------
# PDF text extraction
# ----------------------------

def extract_pdf_text(pdf_path: Path) -> str:
    """Extract text from a PDF. (Scanned PDFs will yield little/no text.)"""
    parts: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    text = "\n".join(parts)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ----------------------------
# Prompt building
# ----------------------------

def build_json_prompt(word_limit: int, source_text: str) -> str:
    return f"""
These attached documents are several chapters that are under a particular course.
Create a study guide based ONLY on the source material. Make sure the information is accurate.

Style and constraints:
- Use UK spellings.
- Do NOT include any leading bullet symbols like "-", "•", or numbering like "1)" in your text.
- Limit to {word_limit} words (excluding the JSON keys themselves).
- Produce TWO short sections:
  1) A heading, then 1 paragraph
  2) A heading, then 1 paragraph
- Then: FIVE short quiz questions with answer keys.
- Then: FIVE short practice questions with answer keys.
- Finally: 25 key summary statements which are informative.

OUTPUT FORMAT (STRICT JSON ONLY; no markdown, no commentary):
{{
  "section1": {{"heading": "...", "paragraph": "..."}},
  "section2": {{"heading": "...", "paragraph": "..."}},
  "quiz": [{{"question":"...", "answer":["..."]}}, ... exactly 5 items ...],
  "practice": [{{"question":"...", "answer":["..."]}}, ... exactly 5 items ...],
  "key_summary_statements": ["...", "...", ... exactly 25 items ...]
}}

SOURCE MATERIAL:
{source_text}
""".strip()


# ----------------------------
# OpenAI call (ChatGPT via API)
# ----------------------------

def call_openai(prompt: str, model: str) -> str:
    """
    Requires: pip install openai
    Expects: OPENAI_API_KEY env var set.
    """
    try:
        from openai import OpenAI
    except ImportError as e:
        raise SystemExit("Missing dependency. Install with: pip install openai") from e

    if not os.environ.get("OPENAI_API_KEY"):
        raise SystemExit(
            "OPENAI_API_KEY is not set.\n"
            "Set it (Windows):  setx OPENAI_API_KEY \"your_api_key_here\" \n"
            "Then close & reopen your terminal."
        )

    client = OpenAI()
    resp = client.responses.create(
        model=model,
        input=prompt,
    )
    return resp.output_text or ""


# ----------------------------
# JSON parsing & validation
# ----------------------------

@dataclass
class StudyGuideJSON:
    section1_heading: str
    section1_paragraph: str
    section2_heading: str
    section2_paragraph: str
    quiz: List[Tuple[str, List[str]]]
    practice: List[Tuple[str, List[str]]]
    key_points: List[str]


def _coerce_to_json(text: str) -> Dict[str, Any]:
    text = text.strip()
    m = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if not m:
        raise ValueError("No JSON object found in LLM output.")
    return json.loads(m.group(0))


def parse_study_guide_json(raw: str) -> StudyGuideJSON:
    data = _coerce_to_json(raw)

    def get_section(i: int) -> Tuple[str, str]:
        sec = data.get(f"section{i}", {})
        return (str(sec.get("heading", "")).strip(), str(sec.get("paragraph", "")).strip())

    s1h, s1p = get_section(1)
    s2h, s2p = get_section(2)

    def get_qa(key: str) -> List[Tuple[str, List[str]]]:
        items = data.get(key, [])
        out: List[Tuple[str, List[str]]] = []
        for it in items:
            q = str(it.get("question", "")).strip()
            a = it.get("answer", [])
            if isinstance(a, str):
                a_list = [a.strip()] if a.strip() else []
            else:
                a_list = [str(x).strip() for x in a if str(x).strip()]
            out.append((q, a_list))
        return out

    quiz = get_qa("quiz")
    practice = get_qa("practice")
    key_points = [str(x).strip() for x in data.get("key_summary_statements", []) if str(x).strip()]

    return StudyGuideJSON(
        section1_heading=s1h,
        section1_paragraph=s1p,
        section2_heading=s2h,
        section2_paragraph=s2p,
        quiz=quiz,
        practice=practice,
        key_points=key_points,
    )


def validate_structure(sg: StudyGuideJSON) -> List[str]:
    issues: List[str] = []
    if not sg.section1_heading or not sg.section1_paragraph:
        issues.append("section1 heading/paragraph missing")
    if not sg.section2_heading or not sg.section2_paragraph:
        issues.append("section2 heading/paragraph missing")
    if len(sg.quiz) != 5:
        issues.append(f"quiz has {len(sg.quiz)} items (expected 5)")
    if len(sg.practice) != 5:
        issues.append(f"practice has {len(sg.practice)} items (expected 5)")
    if len(sg.key_points) != 25:
        issues.append(f"key_summary_statements has {len(sg.key_points)} items (expected 25)")
    return issues


def estimate_word_count(sg: StudyGuideJSON) -> int:
    text = " ".join([
        sg.section1_heading, sg.section1_paragraph,
        sg.section2_heading, sg.section2_paragraph,
        " ".join([q + " " + " ".join(a) for q, a in sg.quiz]),
        " ".join([q + " " + " ".join(a) for q, a in sg.practice]),
        " ".join(sg.key_points),
    ])
    return len(re.findall(r"\b\w+\b", text))


# ----------------------------
# DOCX templating helpers
# ----------------------------

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


def clear_body_from_first_content(doc: Document) -> int:
    first_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            first_idx = i
            break
    if first_idx is None:
        return 0
    for p in list(doc.paragraphs)[first_idx:][::-1]:
        delete_paragraph(p)
    return first_idx


def set_run_font(run, name="Garamond", size_pt=12, bold: Optional[bool] = None):
    run.font.name = name
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold


def add_heading_like_template(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, bold=True)
    p.paragraph_format.line_spacing = 1.15
    return p


def add_paragraph_like_template(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, bold=False)
    p.paragraph_format.line_spacing = 1.15
    return p


def add_label_like_template(doc: Document, text: str):
    p = add_heading_like_template(doc, text)
    doc.add_paragraph("")
    return p


def add_question_like_template(doc: Document, question: str):
    p = add_paragraph_like_template(doc, question)
    p.paragraph_format.left_indent = Pt(18)
    return p


def add_key_point_like_template(doc: Document, statement: str):
    p = add_paragraph_like_template(doc, statement)
    p.paragraph_format.left_indent = Pt(18)
    return p


def add_answer_like_template(doc: Document, answers: List[str]):
    joined = "; ".join([a for a in answers if a])
    p = doc.add_paragraph()
    r = p.add_run(f"Answer: {joined}")
    set_run_font(r, bold=False)
    p.paragraph_format.line_spacing = 1.15
    return p


def write_study_guide_docx(
    template_path: Path,
    output_path: Path,
    course_title: str,
    unit_title: str,
    sg: StudyGuideJSON,
):
    doc = Document(str(template_path))
    clear_body_from_first_content(doc)

    if course_title.strip():
        add_heading_like_template(doc, course_title.strip())
    if unit_title.strip():
        add_heading_like_template(doc, unit_title.strip())

    doc.add_paragraph("")

    add_heading_like_template(doc, sg.section1_heading)
    add_paragraph_like_template(doc, sg.section1_paragraph)
    doc.add_paragraph("")

    add_label_like_template(doc, "Questions:")
    for q, a in sg.quiz:
        add_question_like_template(doc, q)
        add_answer_like_template(doc, a)

    doc.add_paragraph("")

    add_heading_like_template(doc, sg.section2_heading)
    add_paragraph_like_template(doc, sg.section2_paragraph)
    doc.add_paragraph("")

    add_label_like_template(doc, "Questions:")
    for q, a in sg.practice:
        add_question_like_template(doc, q)
        add_answer_like_template(doc, a)

    doc.add_paragraph("")

    add_label_like_template(doc, "Key Summary Statements")
    for s in sg.key_points:
        add_key_point_like_template(doc, s)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


# ----------------------------
# DOCX -> PDF conversion
# ----------------------------

def convert_to_pdf(docx_path: Path, pdf_path: Path):
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # 1) Try docx2pdf (Word)
    try:
        from docx2pdf import convert as docx2pdf_convert
        docx2pdf_convert(str(docx_path), str(pdf_path))
        return
    except Exception:
        pass

    # 2) LibreOffice fallback
    try:
        outdir = pdf_path.parent
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(docx_path)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        produced = outdir / (docx_path.stem + ".pdf")
        if produced.exists() and produced != pdf_path:
            produced.replace(pdf_path)
    except FileNotFoundError:
        raise SystemExit(
            "PDF export failed.\n"
            "Install Microsoft Word + 'pip install docx2pdf' OR install LibreOffice (soffice)."
        )


# ----------------------------
# Core run
# ----------------------------

def run_generation(
    pdf_dir: Path,
    template: Path,
    out_docx: Path,
    out_pdf: Optional[Path],
    course_title: str,
    unit_title: str,
    word_limit_mode: str,
    auto_threshold: int,
    max_source_chars: int,
    retry_on_overlimit: bool,
    retry_on_invalid: bool,
    model: str = DEFAULT_MODEL,
):
    pdfs = sorted([p for p in pdf_dir.glob("*.pdf") if p.is_file()])
    if not pdfs:
        raise SystemExit(f"No PDFs found in: {pdf_dir}")

    if word_limit_mode == "750":
        word_limit = 750
    elif word_limit_mode == "1000":
        word_limit = 1000
    else:
        word_limit = 750 if len(pdfs) <= auto_threshold else 1000

    combined: List[str] = []
    total_chars = 0
    for p in pdfs:
        txt = extract_pdf_text(p)
        combined.append(f"\n\n--- FILE: {p.name} ---\n{txt}")
        total_chars += len(txt)

    if total_chars < 500:
        raise SystemExit(
            "Extracted very little text from the PDFs. They may be scanned images.\n"
            "Run OCR first, then try again."
        )

    source_text = "\n".join(combined).strip()[:max_source_chars]
    prompt = build_json_prompt(word_limit=word_limit, source_text=source_text)

    raw = call_openai(prompt, model=model)
    sg = parse_study_guide_json(raw)

    issues = validate_structure(sg)
    if issues and retry_on_invalid:
        fix = f"""
Your previous output had these issues: {issues}.
Return STRICT JSON ONLY in the SAME schema, fixing the issues and keeping the same word limit ({word_limit}).
""".strip()
        raw2 = call_openai(prompt + "\n\n" + fix, model=model)
        sg = parse_study_guide_json(raw2)

    wc = estimate_word_count(sg)
    if wc > word_limit and retry_on_overlimit:
        tighten = f"""
Your previous JSON exceeded {word_limit} words (approx {wc}). Shorten it to within the limit.
Keep the SAME JSON schema. Do not remove required sections or reduce the number of questions/key statements.
Return STRICT JSON ONLY.
""".strip()
        raw2 = call_openai(prompt + "\n\n" + tighten, model=model)
        sg = parse_study_guide_json(raw2)

    write_study_guide_docx(
        template_path=template,
        output_path=out_docx,
        course_title=course_title,
        unit_title=unit_title,
        sg=sg,
    )

    if out_pdf:
        convert_to_pdf(out_docx, out_pdf)


# ----------------------------
# GUI (clean)
# ----------------------------

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    from tkinter import ttk
    import threading

    root = tk.Tk()
    root.title("Study Guide Generator")
    root.geometry("880x520")
    root.minsize(880, 520)

    style = ttk.Style(root)
    for t in ("vista", "clam"):
        if t in style.theme_names():
            style.theme_use(t)
            break

    # Vars
    pdf_dir_var = tk.StringVar()
    template_var = tk.StringVar(value=str(Path(__file__).with_name("Study Guide template.docx")))
    out_dir_var = tk.StringVar(value=str(Path(__file__).with_name("output")))
    base_name_var = tk.StringVar(value="StudyGuide")

    course_title_var = tk.StringVar()
    unit_title_var = tk.StringVar()

    word_limit_var = tk.StringVar(value="auto")
    auto_threshold_var = tk.IntVar(value=3)

    make_pdf_var = tk.BooleanVar(value=True)
    retry_over_var = tk.BooleanVar(value=True)
    retry_invalid_var = tk.BooleanVar(value=True)
    max_source_chars_var = tk.IntVar(value=120000)

    status_var = tk.StringVar(value="Ready.")

    def browse_pdf_dir():
        d = filedialog.askdirectory(title="Select folder containing chapter PDFs")
        if d:
            pdf_dir_var.set(d)

    def browse_template():
        f = filedialog.askopenfilename(
            title="Select Word template (.docx)",
            filetypes=[("Word document", "*.docx")]
        )
        if f:
            template_var.set(f)

    def browse_out_dir():
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            out_dir_var.set(d)

    # Layout
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    main = ttk.Frame(root, padding=14)
    main.grid(row=0, column=0, sticky="nsew")
    main.columnconfigure(0, weight=1)
    main.rowconfigure(3, weight=1)

    header = ttk.Label(
        main,
        text="Generate a Study Guide from a folder of PDFs using ChatGPT, and output to your Word template.",
        font=("Segoe UI", 11, "bold")
    )
    header.grid(row=0, column=0, sticky="w", pady=(0, 10))

    # 1) Inputs
    inputs = ttk.LabelFrame(main, text="1) Inputs", padding=12)
    inputs.grid(row=1, column=0, sticky="ew")
    inputs.columnconfigure(1, weight=1)

    ttk.Label(inputs, text="PDF folder *").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
    ttk.Entry(inputs, textvariable=pdf_dir_var).grid(row=0, column=1, sticky="ew", pady=6)
    ttk.Button(inputs, text="Browse…", command=browse_pdf_dir).grid(row=0, column=2, padx=(8, 0), pady=6)

    ttk.Label(inputs, text="Word template *").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
    ttk.Entry(inputs, textvariable=template_var).grid(row=1, column=1, sticky="ew", pady=6)
    ttk.Button(inputs, text="Browse…", command=browse_template).grid(row=1, column=2, padx=(8, 0), pady=6)

    ttk.Label(inputs, text="Output folder *").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
    ttk.Entry(inputs, textvariable=out_dir_var).grid(row=2, column=1, sticky="ew", pady=6)
    ttk.Button(inputs, text="Browse…", command=browse_out_dir).grid(row=2, column=2, padx=(8, 0), pady=6)

    ttk.Label(inputs, text="Output name").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=6)
    ttk.Entry(inputs, textvariable=base_name_var).grid(row=3, column=1, sticky="ew", pady=6)

    # 2) Word settings (no model/provider shown)
    settings = ttk.LabelFrame(main, text="2) Word settings", padding=12)
    settings.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    settings.columnconfigure(3, weight=1)

    ttk.Label(settings, text="Word limit").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
    ttk.Combobox(settings, textvariable=word_limit_var, values=["auto", "750", "1000"], state="readonly", width=12)\
        .grid(row=0, column=1, sticky="w", pady=6)

    ttk.Label(settings, text="Auto threshold").grid(row=0, column=2, sticky="w", padx=(12, 8), pady=6)
    ttk.Spinbox(settings, from_=1, to=10, textvariable=auto_threshold_var, width=6)\
        .grid(row=0, column=3, sticky="w", pady=6)

    # 3) Options
    opts = ttk.LabelFrame(main, text="3) Options", padding=12)
    opts.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
    opts.columnconfigure(1, weight=1)

    ttk.Label(opts, text="Course title (optional)").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
    ttk.Entry(opts, textvariable=course_title_var).grid(row=0, column=1, sticky="ew", pady=6)

    ttk.Label(opts, text="Unit title (optional)").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
    ttk.Entry(opts, textvariable=unit_title_var).grid(row=1, column=1, sticky="ew", pady=6)

    checks = ttk.Frame(opts)
    checks.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
    ttk.Checkbutton(checks, text="Export PDF", variable=make_pdf_var).grid(row=0, column=0, padx=(0, 14))
    ttk.Checkbutton(checks, text="Retry if over word limit", variable=retry_over_var).grid(row=0, column=1, padx=(0, 14))
    ttk.Checkbutton(checks, text="Retry if JSON invalid", variable=retry_invalid_var).grid(row=0, column=2, padx=(0, 14))

    adv = ttk.Frame(opts)
    adv.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
    ttk.Label(adv, text="Max source chars").grid(row=0, column=0, padx=(0, 8))
    ttk.Spinbox(adv, from_=20000, to=300000, increment=5000, textvariable=max_source_chars_var, width=10)\
        .grid(row=0, column=1)

    # Footer
    footer = ttk.Frame(main, padding=(0, 10, 0, 0))
    footer.grid(row=4, column=0, sticky="ew")
    footer.columnconfigure(0, weight=1)

    ttk.Label(footer, textvariable=status_var).grid(row=0, column=0, sticky="w")
    generate_btn = ttk.Button(footer, text="Generate study guide")
    generate_btn.grid(row=0, column=1, sticky="e")

    hint = ttk.Label(
        main,
        text="API key: set OPENAI_API_KEY   (Model is fixed internally)",
        foreground="#666"
    )
    hint.grid(row=5, column=0, sticky="w", pady=(8, 0))

    def do_generate():
        pdf_dir = Path(pdf_dir_var.get().strip())
        template = Path(template_var.get().strip())
        out_dir = Path(out_dir_var.get().strip())
        base = base_name_var.get().strip() or "StudyGuide"

        if not pdf_dir.exists():
            messagebox.showerror("Missing input", "Please select a valid PDF folder.")
            return
        if not template.exists():
            messagebox.showerror("Missing input", "Please select a valid Word template (.docx).")
            return
        if not out_dir.exists():
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except Exception:
                messagebox.showerror("Output error", "Output folder is invalid or cannot be created.")
                return

        out_docx = out_dir / f"{base}.docx"
        out_pdf = (out_dir / f"{base}.pdf") if make_pdf_var.get() else None

        def worker():
            try:
                status_var.set("Running…")
                generate_btn.config(state="disabled")

                run_generation(
                    pdf_dir=pdf_dir,
                    template=template,
                    out_docx=out_docx,
                    out_pdf=out_pdf,
                    course_title=course_title_var.get().strip(),
                    unit_title=unit_title_var.get().strip(),
                    word_limit_mode=word_limit_var.get(),
                    auto_threshold=int(auto_threshold_var.get()),
                    max_source_chars=int(max_source_chars_var.get()),
                    retry_on_overlimit=retry_over_var.get(),
                    retry_on_invalid=retry_invalid_var.get(),
                    model=DEFAULT_MODEL,
                )

                status_var.set(f"Done. Saved: {out_docx.name}" + (f" and {out_pdf.name}" if out_pdf else ""))
                messagebox.showinfo("Success", f"Created:\n{out_docx}\n" + (f"{out_pdf}" if out_pdf else ""))
            except Exception as e:
                status_var.set("Error.")
                messagebox.showerror("Error", str(e))
            finally:
                generate_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    generate_btn.configure(command=do_generate)
    root.mainloop()


# ----------------------------
# CLI
# ----------------------------

def parse_args(argv: List[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser()
    ap.add_argument("--gui", action="store_true", help="Launch the desktop UI.")
    ap.add_argument("--pdf-dir", type=Path, help="Folder containing chapter PDFs.")
    ap.add_argument("--template", type=Path, help="Word template .docx.")
    ap.add_argument("--out-docx", type=Path, help="Output .docx path.")
    ap.add_argument("--out-pdf", type=Path, default=None, help="Optional output .pdf path.")
    ap.add_argument("--course-title", type=str, default="")
    ap.add_argument("--unit-title", type=str, default="")
    ap.add_argument("--word-limit", choices=["auto", "750", "1000"], default="auto")
    ap.add_argument("--auto-threshold", type=int, default=3)
    ap.add_argument("--max-source-chars", type=int, default=120000)
    ap.add_argument("--retry-on-overlimit", action="store_true")
    ap.add_argument("--retry-on-invalid", action="store_true")
    return ap.parse_args(argv)


def main():
    args = parse_args(sys.argv[1:])

    if args.gui or len(sys.argv) == 1:
        run_gui()
        return

    required = ["pdf_dir", "template", "out_docx"]
    missing = [r for r in required if getattr(args, r) in (None, "")]
    if missing:
        raise SystemExit(f"Missing required CLI args: {missing}. Or run with --gui.")

    run_generation(
        pdf_dir=args.pdf_dir,
        template=args.template,
        out_docx=args.out_docx,
        out_pdf=args.out_pdf,
        course_title=args.course_title,
        unit_title=args.unit_title,
        word_limit_mode=args.word_limit,
        auto_threshold=args.auto_threshold,
        max_source_chars=args.max_source_chars,
        retry_on_overlimit=args.retry_on_overlimit,
        retry_on_invalid=args.retry_on_invalid,
        model=DEFAULT_MODEL,
    )

    print(f"Created: {args.out_docx}")
    if args.out_pdf:
        print(f"Created: {args.out_pdf}")


if __name__ == "__main__":
    main()
