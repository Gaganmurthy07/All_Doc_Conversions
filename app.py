import io
import os
import re
import json
import tempfile
from pathlib import Path
import streamlit as st
import fitz
import pytesseract
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import inch
from langchain_core.prompts import PromptTemplate, FewShotPromptTemplate
from langchain_groq import ChatGroq

st.set_page_config(
    page_title="Universal AI File Converter",
    page_icon="🔄",
    layout="wide"
)
#st.markdown("""
#<style>
#.main .block-container {max-width: 1100px; padding-top: 2rem; padding-bottom: 3rem;}
#.hero {padding: 1.4rem 1.5rem; border-radius: 18px; background: linear-gradient(120deg, #14213d, #244b74); color: white; margin-bottom: 1.2rem;}
#.hero h1 {color: white; margin-bottom: .3rem;}
#.hero p {color: #e6edf7; margin-bottom: 0;}
#.small-note {color: #64748b; font-size: .9rem;}
#</style>
#<div class="hero">
 # <h1>🔄 AveonRush Document Converter</h1>
#</div>
#""", unsafe_allow_html=True)//
st.title(" 🔄 AveonRush Document Converter")
st.write("Choose a conversion, upload a file, and Download the Converted File.")

CONVERSIONS = {
    "PDF → DOCX": (["pdf"], "docx"),
    "PDF → PPTX": (["pdf"], "pptx"),
    "PDF → TXT": (["pdf"], "txt"),
    "DOCX → PDF": (["docx"], "pdf"),
    "DOCX → PPTX": (["docx"], "pptx"),
    "DOCX → TXT": (["docx"], "txt"),
    "PPTX → DOCX": (["pptx"], "docx"),
    "PPTX → PDF": (["pptx"], "pdf"),
    "PPTX → TXT": (["pptx"], "txt"),
    "TXT → DOCX": (["txt"], "docx"),
    "TXT → PPTX": (["txt"], "pptx"),
    "TXT → PDF": (["txt"], "pdf"),
    "Image → DOCX": (["png", "jpg", "jpeg", "webp", "tif", "tiff"], "docx"),
    "Image → PDF": (["png", "jpg", "jpeg", "webp", "tif", "tiff"], "pdf"),
    "Image → PPTX": (["png", "jpg", "jpeg", "webp", "tif", "tiff"], "pptx"),
    "Image → TXT": (["png", "jpg", "jpeg", "webp", "tif", "tiff"], "txt"),
}

with st.sidebar:
    st.header("Settings")
    conversion = st.selectbox("Choose conversion", list(CONVERSIONS.keys()))
    

accepted_extensions, output_format = CONVERSIONS[conversion]
uploaded = st.file_uploader(
    f"Upload a file for {conversion}",
    type=accepted_extensions,
)

st.info(f"Selected conversion: **{conversion}**  ·  Output format: **.{output_format}**")

EXAMPLES = [
    {
        "input_text": "OCR text: The quick br0wn fox jumps over the l0zy dog.",
        "output_text": "The quick brown fox jumps over the lazy dog."
    },
    {
        "input_text": "OCR text: COMPUTER NETWORKS\\nA network connects devices and allow data shanng.",
        "output_text": "COMPUTER NETWORKS\\n\\nA network connects devices and allows data sharing."
    },
    {
        "input_text": "OCR text: Renewable energy comes from naturaly replenished sources.",
        "output_text": "Renewable energy comes from naturally replenished sources."
    }
]

example_prompt = PromptTemplate(
    input_variables=["input_text", "output_text"],
    template="Input example:\n{input_text}\n\nExpected cleaned output:\n{output_text}"
)

few_shot_prompt = FewShotPromptTemplate(
    examples=EXAMPLES,
    example_prompt=example_prompt,
    prefix=(
        "You clean OCR-extracted document text. Correct only clear OCR mistakes, "
        "preserve the original meaning, names, numbers, formulas, headings and lists. "
        "Do not summarize, add facts, or invent missing content. Preserve paragraph "
        "breaks where possible. Return only the corrected text."
    ),
    suffix="Clean the following extracted text:\n\n{document_text}\n\nCorrected text:",
    input_variables=["document_text"]
)

def run_ocr_on_image(image):
    return pytesseract.image_to_string(image, lang="eng")

def extract_pdf_text(file_bytes, ocr_enabled=True):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    pages = []
    for i, page in enumerate(doc, start=1):
        text = page.get_text("text").strip()
        # OCR the page if no selectable text is present and OCR is enabled.
        if not text and ocr_enabled:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            text = run_ocr_on_image(image).strip()
        pages.append(f"[Page {i}]\n{text}" if text else f"[Page {i}]")
    doc.close()
    return "\n\n".join(pages)

def extract_docx_text(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    parts = []
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            parts.append(paragraph.text)
    for table in doc.tables:
        for row in table.rows:
            parts.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n\n".join(parts)

def extract_pptx_text(file_bytes):
    prs = Presentation(io.BytesIO(file_bytes))
    parts = []
    for i, slide in enumerate(prs.slides, start=1):
        parts.append(f"[Slide {i}]")
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.text.strip():
                parts.append(shape.text.strip())
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    parts.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(parts)

def extract_text_file(file_bytes):
    for encoding in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            pass
    return file_bytes.decode("utf-8", errors="replace")

def extract_image_text(file_bytes, ocr_enabled=True):
    if not ocr_enabled:
        raise ValueError("Enable OCR in the sidebar to read text from images.")
    image = Image.open(io.BytesIO(file_bytes))
    return run_ocr_on_image(image)

def groq_clean_text(text, api_key, model):
    if not api_key:
        api_key = os.environ.get("GROQ_API_KEY", "")
    if not api_key:
        raise ValueError("Enter a Groq API key in the sidebar or configure GROQ_API_KEY in Streamlit Secrets.")
    llm = ChatGroq(
        api_key=api_key,
        model=model,
        temperature=0
    )
    chain = few_shot_prompt | llm
    # Keep requests manageable. Chunking by characters is a basic safeguard;
    # larger production projects should split by paragraphs/token count.
    chunks = [text[i:i+12000] for i in range(0, len(text), 12000)] or [""]
    cleaned = []
    for chunk in chunks:
        result = chain.invoke({"document_text": chunk})
        cleaned.append(result.content.strip())
    return "\n\n".join(cleaned)

def split_sections(text):
    # Split at page/slide markers and blank-line groups, retaining useful sections.
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]
    return blocks or [text.strip()]

def make_docx(text):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    for block in split_sections(text):
        lines = block.splitlines()
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if re.match(r"^\[(Page|Slide)\s+\d+\]$", line, re.I):
                doc.add_heading(line.strip("[]"), level=1)
            elif len(line) < 100 and (line.isupper() or line.endswith(":")):
                doc.add_heading(line.rstrip(":"), level=2)
            elif re.match(r"^[-•*]\s+", line):
                doc.add_paragraph(re.sub(r"^[-•*]\s+", "", line), style="List Bullet")
            else:
                doc.add_paragraph(line)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def make_pptx(text):
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)
    blocks = split_sections(text)
    # First slide is a title slide.
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Converted Document"
    if len(title_slide.placeholders) > 1:
        title_slide.placeholders[1].text = "Generated from extracted document text"
    # Keep slides readable: up to 5 bullets per slide.
    bullets = []
    for block in blocks:
        for line in block.splitlines():
            line = line.strip()
            if line and not re.match(r"^\[(Page|Slide)\s+\d+\]$", line, re.I):
                bullets.append(line)
    if not bullets:
        bullets = ["No readable text was extracted."]
    for start in range(0, len(bullets), 5):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Content {start // 5 + 1}"
        body = slide.placeholders[1].text_frame
        body.clear()
        for idx, bullet in enumerate(bullets[start:start+5]):
            p = body.paragraphs[0] if idx == 0 else body.add_paragraph()
            p.text = bullet
            p.level = 0
            p.font.size = PptPt(20)
    buffer = io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()

def make_pdf(text):
    buffer = io.BytesIO()
    pdf = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=0.65*inch, leftMargin=0.65*inch,
        topMargin=0.65*inch, bottomMargin=0.65*inch
    )
    styles = getSampleStyleSheet()
    story = []
    for block in split_sections(text):
        for line in block.splitlines():
            line = line.strip()
            if not line:
                continue
            safe = (line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
            if re.match(r"^\[(Page|Slide)\s+\d+\]$", line, re.I) or (line.isupper() and len(line) < 100):
                story.append(Paragraph(safe, styles["Heading2"]))
            else:
                story.append(Paragraph(safe, styles["BodyText"]))
                story.append(Spacer(1, 5))
        story.append(Spacer(1, 8))
    pdf.build(story)
    return buffer.getvalue()

def make_txt(text):
    return text.encode("utf-8")

def output_bytes(text, fmt):
    if fmt == "docx":
        return make_docx(text), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if fmt == "pptx":
        return make_pptx(text), "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    if fmt == "pdf":
        return make_pdf(text), "application/pdf"
    return make_txt(text), "text/plain; charset=utf-8"

def extract_input(ext, file_bytes, ocr_enabled):
    if ext == "pdf":
        return extract_pdf_text(file_bytes, ocr_enabled)
    if ext == "docx":
        return extract_docx_text(file_bytes)
    if ext == "pptx":
        return extract_pptx_text(file_bytes)
    if ext == "txt":
        return extract_text_file(file_bytes)
    if ext in {"png", "jpg", "jpeg", "webp", "tif", "tiff"}:
        return extract_image_text(file_bytes, ocr_enabled)
    raise ValueError(f"Unsupported input format: {ext}")

if uploaded:
    input_ext = Path(uploaded.name).suffix.lower().lstrip(".")
    if st.button("🚀 Convert file", type="primary", use_container_width=True):
        try:
            file_bytes = uploaded.getvalue()
            with st.spinner("Extracting text from your file..."):
                extracted = extract_input(input_ext, file_bytes, use_ocr)
            if not extracted.strip():
                st.warning("No text was detected. If this is a scan, enable OCR or try a clearer file.")
                st.stop()

            with st.expander("Preview extracted text", expanded=False):
                st.text(extracted[:12000])

            final_text = extracted
            if use_ai:
                with st.spinner("Sending text to Groq for few-shot AI cleanup..."):
                    final_text = groq_clean_text(extracted, groq_key, model_name)

            st.success("Text extraction and processing completed.")
            with st.expander("Preview final text", expanded=False):
                st.text(final_text[:12000])

            output, mime = output_bytes(final_text, output_format)
            original_stem = Path(uploaded.name).stem
            output_name = f"{original_stem}_converted.{output_format}"
            st.download_button(
                label=f"⬇️ Download {output_name}",
                data=output,
                file_name=output_name,
                mime=mime,
                use_container_width=True
            )
            st.caption(
                "This app converts the extracted text and creates a new file. "
                "It does not guarantee exact preservation of original layouts, images, charts, animations, or complex tables."
            )
        except Exception as exc:
            st.error(f"Conversion failed: {exc}")
            st.info("Check that the input file is valid. If AI cleanup is enabled, verify your Groq API key and model availability.")
else:
    st.subheader("How to use")
    st.markdown("""
    1. Choose a conversion from the **Settings** dropdown in the sidebar.
    2. Upload a supported file.
    3. Click **Convert file** and download the result.
    """)
    st.warning("Don't Leave the Page While Processing.")
