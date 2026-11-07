import os, re, uuid, time
from pathlib import Path
from typing import Optional
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from docx import Document
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pptx import Presentation
from pptx.util import Inches, Pt

APP_NAME = "LangMind File Generator API"
STORAGE_DIR = Path(os.getenv("FILE_STORAGE_DIR", "/tmp/generated_files"))
STORAGE_DIR.mkdir(parents=True, exist_ok=True)
MAX_CONTENT_CHARS = int(os.getenv("MAX_CONTENT_CHARS", "200000"))
FILE_TTL_SECONDS = int(os.getenv("FILE_TTL_SECONDS", "86400"))

app = FastAPI(
    title=APP_NAME,
    version="1.0.0",
    description="Create DOCX, XLSX, PDF and PPTX files from Markdown or JSON and return a download URL."
)

class CreateFileRequest(BaseModel):
    filename: str = Field(..., min_length=1, max_length=120, description="Output filename, with or without extension")
    format: str = Field(..., description="docx, xlsx, pdf or pptx")
    content: str = Field(..., min_length=1, description="Markdown content; PPTX also accepts a JSON presentation spec")

def safe_filename(name: str, ext: str) -> str:
    name = Path(name).name
    name = re.sub(r"[^A-Za-z0-9ก-๙._ -]", "_", name).strip(" .")
    if not name:
        name = "generated-document"
    if not name.lower().endswith("." + ext):
        name = f"{Path(name).stem}.{ext}"
    return name[:140]

def markdown_lines(content: str):
    return [line.rstrip() for line in content.replace("\r\n", "\n").split("\n")]

def create_docx(path: Path, content: str):
    doc = Document()
    for line in markdown_lines(content):
        s = line.strip()
        if not s:
            doc.add_paragraph("")
        elif s.startswith("### "):
            doc.add_heading(s[4:], level=3)
        elif s.startswith("## "):
            doc.add_heading(s[3:], level=2)
        elif s.startswith("# "):
            doc.add_heading(s[2:], level=1)
        elif s.startswith(("- ", "* ")):
            doc.add_paragraph(s[2:], style="List Bullet")
        elif re.match(r"^\d+\.\s", s):
            doc.add_paragraph(re.sub(r"^\d+\.\s", "", s), style="List Number")
        else:
            doc.add_paragraph(re.sub(r"\*\*(.*?)\*\*", r"\1", s))
    doc.save(path)

def extract_md_tables(content: str):
    lines = markdown_lines(content)
    tables, plain = [], []
    i = 0
    while i < len(lines):
        if "|" in lines[i] and i + 1 < len(lines) and re.match(r"^\s*\|?[\s:|-]+\|?\s*$", lines[i+1]):
            block = [lines[i]]
            i += 2
            while i < len(lines) and "|" in lines[i]:
                block.append(lines[i]); i += 1
            rows = []
            for row in block:
                cells = [c.strip() for c in row.strip().strip("|").split("|")]
                rows.append(cells)
            tables.append(rows)
        else:
            plain.append(lines[i]); i += 1
    return tables, plain

def create_xlsx(path: Path, content: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "ข้อมูล"
    tables, plain = extract_md_tables(content)
    if tables:
        for ti, rows in enumerate(tables):
            target = ws if ti == 0 else wb.create_sheet(f"ตาราง{ti+1}")
            for row in rows:
                if all(re.fullmatch(r":?-{3,}:?", c or "-") for c in row):
                    continue
                target.append(row)
            target.freeze_panes = "A2"
            for col in target.columns:
                letter = col[0].column_letter
                max_len = max((len(str(c.value or "")) for c in col), default=10)
                target.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
    else:
        for line in plain:
            if line.strip():
                ws.append([line.strip()])
        ws.column_dimensions["A"].width = 80
    wb.save(path)

def create_pdf(path: Path, content: str):
    # Embed Noto Sans Thai so Thai text is rendered correctly in generated PDFs.
    font_dir = Path(__file__).parent / "fonts"
    regular_font = font_dir / "NotoSansThai-Regular.ttf"
    bold_font = font_dir / "NotoSansThai-Bold.ttf"
    if regular_font.exists() and bold_font.exists():
        if "NotoThai" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("NotoThai", str(regular_font)))
            pdfmetrics.registerFont(TTFont("NotoThai-Bold", str(bold_font)))
            pdfmetrics.registerFontFamily("NotoThai", normal="NotoThai", bold="NotoThai-Bold",
                                          italic="NotoThai", boldItalic="NotoThai-Bold")
        styles = getSampleStyleSheet()
        for style_name in ("Normal", "BodyText", "Title", "Heading1", "Heading2", "Heading3"):
            styles[style_name].fontName = "NotoThai"
    else:
        styles = getSampleStyleSheet()
    story = []
    for line in markdown_lines(content):
        s = line.strip()
        if not s:
            story.append(Spacer(1, 4))
        elif s.startswith("# "):
            story.append(Paragraph(s[2:], styles["Title"]))
        elif s.startswith("## "):
            story.append(Paragraph(s[3:], styles["Heading2"]))
        elif s.startswith("### "):
            story.append(Paragraph(s[4:], styles["Heading3"]))
        elif s.startswith(("- ", "* ")):
            story.append(Paragraph("• " + s[2:], styles["BodyText"]))
        else:
            escaped = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            escaped = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", escaped)
            story.append(Paragraph(escaped, styles["BodyText"]))
            story.append(Spacer(1, 3))
    SimpleDocTemplate(str(path), pagesize=A4, rightMargin=18*mm, leftMargin=18*mm,
                      topMargin=18*mm, bottomMargin=18*mm).build(story)

def create_pptx(path: Path, content: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    try:
        import json
        spec = json.loads(content)
    except Exception:
        spec = None
    if isinstance(spec, dict) and isinstance(spec.get("slides"), list):
        slides = spec["slides"]
        for item in slides:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = str(item.get("title", ""))
            tf = slide.placeholders[1].text_frame
            tf.clear()
            body = item.get("content", "")
            bullets = item.get("bullets", [])
            if body:
                tf.text = str(body)
            for bullet in bullets:
                p = tf.add_paragraph()
                p.text = str(bullet)
                p.level = 0
    else:
        chunks, current = [], None
        for line in markdown_lines(content):
            if line.startswith("# "):
                if current: chunks.append(current)
                current = {"title": line[2:].strip(), "body": []}
            elif current is not None:
                if line.strip(): current["body"].append(line.strip().lstrip("-* "))
        if current: chunks.append(current)
        if not chunks:
            chunks = [{"title": "เอกสารนำเสนอ", "body": [x for x in markdown_lines(content) if x.strip()]}]
        for item in chunks:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = item["title"]
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for idx, text in enumerate(item["body"][:12]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.text = text
                p.level = 0
                p.font.size = Pt(22)
    prs.save(path)

@app.get("/health")
def health():
    return {"status": "ok", "service": APP_NAME, "supported_formats": ["docx", "xlsx", "pdf", "pptx"]}

@app.post("/create-file")
def create_file(payload: CreateFileRequest, request: Request):
    if len(payload.content) > MAX_CONTENT_CHARS:
        raise HTTPException(status_code=413, detail=f"content exceeds {MAX_CONTENT_CHARS} characters")
    fmt = payload.format.lower().strip().lstrip(".")
    creators = {"docx": create_docx, "xlsx": create_xlsx, "pdf": create_pdf, "pptx": create_pptx}
    if fmt not in creators:
        raise HTTPException(status_code=400, detail="Unsupported format. Use docx, xlsx, pdf or pptx.")
    file_id = uuid.uuid4().hex
    filename = safe_filename(payload.filename, fmt)
    stored_path = STORAGE_DIR / f"{file_id}.{fmt}"
    try:
        creators[fmt](stored_path, payload.content)
    except Exception as exc:
        if stored_path.exists():
            stored_path.unlink()
        raise HTTPException(status_code=500, detail=f"Could not create file: {str(exc)[:300]}")
    metadata = STORAGE_DIR / f"{file_id}.meta"
    metadata.write_text(f"{filename}\n{time.time()}", encoding="utf-8")
    base_url = os.getenv("PUBLIC_BASE_URL", str(request.base_url).rstrip("/"))
    return {
        "success": True,
        "filename": filename,
        "format": fmt,
        "download_url": f"{base_url}/download/{file_id}",
        "expires_in_seconds": FILE_TTL_SECONDS
    }

@app.get("/download/{file_id}")
def download_file(file_id: str):
    if not re.fullmatch(r"[a-f0-9]{32}", file_id):
        raise HTTPException(status_code=404, detail="File not found")
    matches = [STORAGE_DIR / f"{file_id}.{ext}" for ext in ("docx", "xlsx", "pdf", "pptx")]
    path = next((p for p in matches if p.exists()), None)
    meta = STORAGE_DIR / f"{file_id}.meta"
    if path is None or not meta.exists():
        raise HTTPException(status_code=404, detail="File not found or expired")
    try:
        lines = meta.read_text(encoding="utf-8").splitlines()
        filename = lines[0]
        created = float(lines[1])
        if time.time() - created > FILE_TTL_SECONDS:
            path.unlink(missing_ok=True); meta.unlink(missing_ok=True)
            raise HTTPException(status_code=410, detail="Download link expired")
    except HTTPException:
        raise
    except Exception:
        filename = path.name
    return FileResponse(path, filename=filename, media_type="application/octet-stream")
