import streamlit as st
import io, os, re, subprocess, tempfile, zipfile, shutil
from pathlib import Path

# ── page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Convert", page_icon="⇄", layout="centered")

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600&display=swap');

*, *::before, *::after { box-sizing: border-box; }

html, body, [data-testid="stAppViewContainer"] {
    background: #0D0D0F !important;
    color: #E8E8E4;
    font-family: 'DM Sans', sans-serif;
}

#MainMenu, footer, header { visibility: hidden; }

[data-testid="stAppViewContainer"] > .main > .block-container {
    max-width: 680px;
    padding: 3.5rem 1.5rem 3rem;
    margin: 0 auto;
}

/* ── typography ── */
h1 {
    font-family: 'DM Mono', monospace !important;
    font-size: 1.15rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.08em;
    color: #888 !important;
    text-transform: uppercase;
    margin-bottom: 0.2rem !important;
}

.tagline {
    font-size: 2.4rem;
    font-weight: 300;
    color: #E8E8E4;
    line-height: 1.15;
    margin-bottom: 3rem;
    letter-spacing: -0.02em;
}
.tagline span { color: #7EE8A2; }

/* ── step label ── */
.step-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #555;
    margin-bottom: 0.7rem;
}

/* ── format grid ── */
.fmt-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 0.5rem;
    margin-bottom: 2rem;
}

/* ── format button ── */
.stButton > button {
    background: #18181C !important;
    color: #C8C8C0 !important;
    border: 1.5px solid #2A2A30 !important;
    border-radius: 6px !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.04em;
    padding: 0.45rem 0.85rem !important;
    min-width: 70px !important;
    transition: all 0.15s ease !important;
    cursor: pointer !important;
}
.stButton > button:hover {
    background: #232329 !important;
    border-color: #4A4A55 !important;
    color: #E8E8E4 !important;
}

/* active / selected state via classes injected by JS */
.active-btn > button {
    background: #7EE8A2 !important;
    border-color: #7EE8A2 !important;
    color: #0D0D0F !important;
}

/* ── arrow display ── */
.arrow-row {
    display: flex;
    align-items: center;
    gap: 1rem;
    margin: 1.5rem 0 2rem;
    font-family: 'DM Mono', monospace;
    font-size: 1.1rem;
}
.pill {
    background: #7EE8A2;
    color: #0D0D0F;
    border-radius: 5px;
    padding: 0.3rem 0.7rem;
    font-weight: 600;
    font-size: 0.95rem;
}
.pill-dim {
    background: #7EE8A222;
    color: #7EE8A2;
    border-radius: 5px;
    padding: 0.3rem 0.7rem;
    font-size: 0.95rem;
}
.arrow { color: #444; font-size: 1.3rem; }

/* ── upload zone ── */
[data-testid="stFileUploader"] {
    border: 1.5px dashed #2A2A30 !important;
    border-radius: 10px;
    padding: 1rem;
    background: #18181C;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #7EE8A2 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    color: #555 !important;
}

/* ── progress / status ── */
[data-testid="stSpinner"] { color: #7EE8A2 !important; }

/* ── download button ── */
[data-testid="stDownloadButton"] > button {
    width: 100% !important;
    background: #7EE8A2 !important;
    color: #0D0D0F !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.95rem !important;
    font-weight: 600 !important;
    padding: 0.75rem 1rem !important;
    margin-top: 1rem;
    letter-spacing: 0.02em;
    transition: opacity 0.15s !important;
}
[data-testid="stDownloadButton"] > button:hover {
    opacity: 0.85 !important;
}

/* ── divider ── */
hr { border-color: #1E1E24 !important; margin: 1.5rem 0 !important; }

/* ── alerts ── */
[data-testid="stAlert"] {
    border-radius: 8px !important;
    font-size: 0.88rem;
}

/* ── info text ── */
.ocr-note {
    font-size: 0.78rem;
    color: #666;
    margin-top: 0.4rem;
    font-family: 'DM Mono', monospace;
}

/* ── reset link ── */
.reset-btn > button {
    background: transparent !important;
    color: #555 !important;
    border: 1px solid #2A2A30 !important;
    font-size: 0.78rem !important;
    padding: 0.3rem 0.6rem !important;
    min-width: 0 !important;
}
.reset-btn > button:hover {
    color: #E8E8E4 !important;
    border-color: #555 !important;
}
</style>
""", unsafe_allow_html=True)

# ── conversion map ────────────────────────────────────────────────────────────
CONVERSIONS = {
    "PDF":   ["PPT", "DOCX", "TXT", "IMG"],
    "DOCX":  ["PDF", "TXT"],
    "PPT":   ["PDF", "IMG", "TXT"],
    "IMG":   ["PDF", "JPG", "PNG", "TXT (OCR)"],
    "JPG":   ["PNG", "PDF"],
    "PNG":   ["JPG", "PDF"],
    "TXT":   ["PDF", "DOCX"],
    "CSV":   ["XLSX"],
    "XLSX":  ["CSV"],
}

FORMAT_ACCEPT = {
    "PDF":  ["pdf"],
    "DOCX": ["docx"],
    "PPT":  ["pptx", "ppt"],
    "IMG":  ["jpg", "jpeg", "png", "webp", "bmp", "tiff"],
    "JPG":  ["jpg", "jpeg"],
    "PNG":  ["png"],
    "TXT":  ["txt"],
    "CSV":  ["csv"],
    "XLSX": ["xlsx"],
}

FORMAT_MIME = {
    "PDF":  "application/pdf",
    "DOCX": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "PPT":  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "TXT":  "text/plain",
    "XLSX": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "CSV":  "text/csv",
    "ZIP":  "application/zip",
    "JPG":  "image/jpeg",
    "PNG":  "image/png",
}


# ── converters ────────────────────────────────────────────────────────────────

def pdf_to_ppt(data: bytes) -> bytes:
    import pdfplumber
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    BG   = RGBColor(0x1A, 0x1A, 0x2E)
    ACC  = RGBColor(0x7E, 0xE8, 0xA2)
    WHT  = RGBColor(0xFF, 0xFF, 0xFF)
    blank = prs.slide_layouts[6]

    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = (page.extract_text() or "").strip()
            slide = prs.slides.add_slide(blank)
            fill = slide.background.fill
            fill.solid(); fill.fore_color.rgb = BG

            bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.07), prs.slide_height)
            bar.fill.solid(); bar.fill.fore_color.rgb = ACC; bar.line.fill.background()

            lines = [l.strip() for l in text.splitlines() if l.strip()]
            title = lines[0] if lines else f"Page {i+1}"
            body  = "\n".join(lines[1:]) if len(lines) > 1 else ""

            tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12.5), Inches(1.1))
            p  = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
            r  = p.add_run(); r.text = title; r.font.size = Pt(26); r.font.bold = True; r.font.color.rgb = WHT

            if body:
                bb = slide.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(12.5), Inches(5.8))
                bb.text_frame.word_wrap = True
                for j, ln in enumerate(body.splitlines()):
                    para = bb.text_frame.paragraphs[0] if j == 0 else bb.text_frame.add_paragraph()
                    run  = para.add_run(); run.text = ln.strip()
                    run.font.size = Pt(13); run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)

    buf = io.BytesIO(); prs.save(buf); return buf.getvalue()


def pdf_to_docx(data: bytes) -> bytes:
    import pdfplumber
    from docx import Document
    from docx.shared import Pt, RGBColor
    doc = Document()
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for i, page in enumerate(pdf.pages):
            if i > 0:
                doc.add_page_break()
            text = page.extract_text() or ""
            lines = text.splitlines()
            for j, ln in enumerate(lines):
                ln = ln.strip()
                if not ln:
                    continue
                if j == 0 and len(ln) < 80:
                    doc.add_heading(ln, level=2)
                else:
                    doc.add_paragraph(ln)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


def pdf_to_txt(data: bytes) -> bytes:
    import pdfplumber
    out = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for i, page in enumerate(pdf.pages):
            out.append(f"─── Page {i+1} ───")
            out.append(page.extract_text() or "")
    return "\n".join(out).encode()


def pdf_to_img(data: bytes) -> bytes:
    from pdf2image import convert_from_bytes
    images = convert_from_bytes(data, dpi=150)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for i, img in enumerate(images):
            ib = io.BytesIO()
            img.save(ib, format="PNG")
            zf.writestr(f"page_{i+1:03d}.png", ib.getvalue())
    return buf.getvalue()


def _soffice_convert(data: bytes, in_ext: str, out_fmt: str) -> bytes:
    """Use LibreOffice headless to convert. Returns output bytes."""
    with tempfile.TemporaryDirectory() as tmp:
        infile  = os.path.join(tmp, f"input.{in_ext}")
        with open(infile, "wb") as f:
            f.write(data)
        subprocess.run(
            ["soffice", "--headless", "--norestore", "--nofirststartwizard",
             f"--convert-to", out_fmt, "--outdir", tmp, infile],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=120
        )
        # find output
        for fname in os.listdir(tmp):
            if fname != f"input.{in_ext}":
                with open(os.path.join(tmp, fname), "rb") as f:
                    return f.read()
    raise RuntimeError("LibreOffice conversion produced no output")


def docx_to_pdf(data: bytes) -> bytes:
    return _soffice_convert(data, "docx", "pdf")


def docx_to_txt(data: bytes) -> bytes:
    from docx import Document
    doc = Document(io.BytesIO(data))
    lines = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(lines).encode()


def ppt_to_pdf(data: bytes) -> bytes:
    return _soffice_convert(data, "pptx", "pdf")


def ppt_to_img(data: bytes) -> bytes:
    pdf_data = ppt_to_pdf(data)
    return pdf_to_img(pdf_data)


def ppt_to_txt(data: bytes) -> bytes:
    from pptx import Presentation
    prs = Presentation(io.BytesIO(data))
    out = []
    for i, slide in enumerate(prs.slides):
        out.append(f"─── Slide {i+1} ───")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = para.text.strip()
                    if t:
                        out.append(t)
    return "\n".join(out).encode()


def img_to_pdf(data: bytes) -> bytes:
    from PIL import Image
    img = Image.open(io.BytesIO(data)).convert("RGB")
    buf = io.BytesIO(); img.save(buf, format="PDF"); return buf.getvalue()


def img_convert(data: bytes, fmt: str) -> bytes:
    from PIL import Image
    img = Image.open(io.BytesIO(data))
    if fmt.upper() == "JPG":
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG" if fmt.upper() == "JPG" else fmt.upper())
    return buf.getvalue()


def img_to_txt_ocr(data: bytes) -> bytes:
    import pytesseract
    from PIL import Image
    img = Image.open(io.BytesIO(data))
    text = pytesseract.image_to_string(img)
    return text.encode()


def txt_to_pdf(data: bytes) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.units import inch

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                             leftMargin=inch, rightMargin=inch,
                             topMargin=inch, bottomMargin=inch)
    styles = getSampleStyleSheet()
    body_style = ParagraphStyle("body", parent=styles["Normal"],
                                fontSize=11, leading=16, alignment=TA_LEFT)
    story = []
    text = data.decode("utf-8", errors="replace")
    for line in text.splitlines():
        safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        story.append(Paragraph(safe if safe.strip() else "&nbsp;", body_style))
        story.append(Spacer(1, 2))
    doc.build(story)
    return buf.getvalue()


def txt_to_docx(data: bytes) -> bytes:
    from docx import Document
    doc = Document()
    text = data.decode("utf-8", errors="replace")
    for line in text.splitlines():
        doc.add_paragraph(line)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


def csv_to_xlsx(data: bytes) -> bytes:
    import pandas as pd
    df = pd.read_csv(io.BytesIO(data))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def xlsx_to_csv(data: bytes) -> bytes:
    import pandas as pd
    df = pd.read_excel(io.BytesIO(data))
    buf = io.BytesIO(); df.to_csv(buf, index=False); return buf.getvalue()


# ── router ────────────────────────────────────────────────────────────────────

def convert(src: str, tgt: str, data: bytes) -> tuple[bytes, str, str]:
    """Returns (output_bytes, extension, mime_type)."""
    key = f"{src}→{tgt}"
    dispatch = {
        "PDF→PPT":      lambda d: (pdf_to_ppt(d),      "pptx", FORMAT_MIME["PPT"]),
        "PDF→DOCX":     lambda d: (pdf_to_docx(d),     "docx", FORMAT_MIME["DOCX"]),
        "PDF→TXT":      lambda d: (pdf_to_txt(d),      "txt",  FORMAT_MIME["TXT"]),
        "PDF→IMG":      lambda d: (pdf_to_img(d),      "zip",  FORMAT_MIME["ZIP"]),
        "DOCX→PDF":     lambda d: (docx_to_pdf(d),     "pdf",  FORMAT_MIME["PDF"]),
        "DOCX→TXT":     lambda d: (docx_to_txt(d),     "txt",  FORMAT_MIME["TXT"]),
        "PPT→PDF":      lambda d: (ppt_to_pdf(d),      "pdf",  FORMAT_MIME["PDF"]),
        "PPT→IMG":      lambda d: (ppt_to_img(d),      "zip",  FORMAT_MIME["ZIP"]),
        "PPT→TXT":      lambda d: (ppt_to_txt(d),      "txt",  FORMAT_MIME["TXT"]),
        "IMG→PDF":      lambda d: (img_to_pdf(d),      "pdf",  FORMAT_MIME["PDF"]),
        "IMG→JPG":      lambda d: (img_convert(d,"JPG"),"jpg", FORMAT_MIME["JPG"]),
        "IMG→PNG":      lambda d: (img_convert(d,"PNG"),"png", FORMAT_MIME["PNG"]),
        "IMG→TXT (OCR)":lambda d: (img_to_txt_ocr(d), "txt",  FORMAT_MIME["TXT"]),
        "JPG→PNG":      lambda d: (img_convert(d,"PNG"),"png", FORMAT_MIME["PNG"]),
        "JPG→PDF":      lambda d: (img_to_pdf(d),      "pdf",  FORMAT_MIME["PDF"]),
        "PNG→JPG":      lambda d: (img_convert(d,"JPG"),"jpg", FORMAT_MIME["JPG"]),
        "PNG→PDF":      lambda d: (img_to_pdf(d),      "pdf",  FORMAT_MIME["PDF"]),
        "TXT→PDF":      lambda d: (txt_to_pdf(d),      "pdf",  FORMAT_MIME["PDF"]),
        "TXT→DOCX":     lambda d: (txt_to_docx(d),     "docx", FORMAT_MIME["DOCX"]),
        "CSV→XLSX":     lambda d: (csv_to_xlsx(d),     "xlsx", FORMAT_MIME["XLSX"]),
        "XLSX→CSV":     lambda d: (xlsx_to_csv(d),     "csv",  FORMAT_MIME["CSV"]),
    }
    fn = dispatch.get(key)
    if fn is None:
        raise ValueError(f"Conversion {key} not supported")
    return fn(data)


# ── session state ─────────────────────────────────────────────────────────────
if "src" not in st.session_state:
    st.session_state.src = None
if "tgt" not in st.session_state:
    st.session_state.tgt = None


def reset():
    st.session_state.src = None
    st.session_state.tgt = None


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("# CONVERT")
st.markdown('<div class="tagline">Drop a file.<br><span>Get it back differently.</span></div>',
            unsafe_allow_html=True)

# ── STEP 1 : choose source ────────────────────────────────────────────────────
st.markdown('<div class="step-label">01 — What do you have?</div>', unsafe_allow_html=True)
all_fmts = list(CONVERSIONS.keys())
cols = st.columns(len(all_fmts))
for col, fmt in zip(cols, all_fmts):
    with col:
        label = fmt
        clicked = st.button(label, key=f"src_{fmt}")
        if clicked:
            if st.session_state.src == fmt:
                # deselect
                st.session_state.src = None
                st.session_state.tgt = None
            else:
                st.session_state.src = fmt
                st.session_state.tgt = None

# ── STEP 2 : choose target ────────────────────────────────────────────────────
if st.session_state.src:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="step-label">02 — What do you want?</div>', unsafe_allow_html=True)
    targets = CONVERSIONS[st.session_state.src]
    tcols = st.columns(max(len(targets), 1))
    for col, tgt in zip(tcols, targets):
        with col:
            clicked = st.button(tgt, key=f"tgt_{tgt}")
            if clicked:
                st.session_state.tgt = tgt if st.session_state.tgt != tgt else None

# ── STEP 3 : upload & convert ─────────────────────────────────────────────────
if st.session_state.src and st.session_state.tgt:
    src = st.session_state.src
    tgt = st.session_state.tgt

    st.markdown("<hr>", unsafe_allow_html=True)

    # Arrow display
    st.markdown(
        f'<div class="arrow-row">'
        f'<span class="pill">{src}</span>'
        f'<span class="arrow">→</span>'
        f'<span class="pill-dim">{tgt}</span>'
        f'</div>',
        unsafe_allow_html=True
    )

    if tgt == "TXT (OCR)":
        st.markdown('<div class="ocr-note">⚠ OCR accuracy depends on image quality.</div>',
                    unsafe_allow_html=True)

    accept = FORMAT_ACCEPT.get(src, [])
    accept_str = ", ".join(f".{e}" for e in accept)
    uploaded = st.file_uploader(
        f"Upload your {src} file",
        type=accept,
        key=f"upload_{src}_{tgt}",
        label_visibility="collapsed"
    )

    if uploaded:
        with st.spinner(f"Converting {src} → {tgt}…"):
            try:
                raw = uploaded.read()
                out_bytes, ext, mime = convert(src, tgt, raw)
                base = Path(uploaded.name).stem

                st.success(f"Done — {len(out_bytes):,} bytes")
                st.download_button(
                    label=f"⬇  Download  {base}.{ext}",
                    data=out_bytes,
                    file_name=f"{base}.{ext}",
                    mime=mime,
                )
            except Exception as e:
                st.error(f"Conversion failed: {e}")

    # reset
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="reset-btn">', unsafe_allow_html=True)
    if st.button("← start over", key="reset"):
        reset()
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
