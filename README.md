# LangMind File Generator API

Backend API สำหรับสร้างไฟล์ Word (.docx), Excel (.xlsx), PDF (.pdf) และ PowerPoint (.pptx) จาก Markdown หรือ JSON แล้วส่ง URL ดาวน์โหลดกลับมา

## 1) ทดสอบบนเครื่องตัวเอง

ต้องติดตั้ง Python 3.11 ขึ้นไป

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
uvicorn main:app --reload
```

เปิดเอกสาร API ที่ `http://127.0.0.1:8000/docs` และตรวจสุขภาพที่ `http://127.0.0.1:8000/health`

## 2) Deploy บน Render

1. สร้างบัญชี GitHub และ Render
2. สร้าง GitHub repository ใหม่ เช่น `langmind-file-generator`
3. อัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้ไปไว้ที่ root ของ repository
4. ใน Render เลือก **New + → Blueprint**
5. เชื่อม GitHub แล้วเลือก repository นี้
6. Render อ่าน `render.yaml` และสร้าง Web Service ให้ กดยืนยัน Deploy
7. เมื่อสถานะเป็น Live ให้คัดลอก URL เช่น `https://langmind-file-generator.onrender.com`
8. ทดสอบ `https://YOUR-SERVICE.onrender.com/health` ต้องได้ JSON ที่มี `"status":"ok"`

หรือเลือก **New + → Web Service** แล้วกำหนด:
- Build Command: `pip install -r requirements.txt`
- Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
- Health Check Path: `/health`

## 3) ตั้งค่า LangMind

- **Server URL**: `https://YOUR-SERVICE.onrender.com`
- **Spec URL**: `https://YOUR-SERVICE.onrender.com/openapi.json`
- หรือคัดลอก OpenAPI YAML จาก `openapi.yaml` ไปวางในช่อง OpenAPI JSON/YAML

## 4) เรียก API

POST `/create-file`

ตัวอย่าง DOCX:

```json
{
  "filename": "quotation-2026.docx",
  "format": "docx",
  "content": "# ใบเสนอราคา\n\n| รายการ | จำนวน | ราคา |\n|---|---:|---:|\n| สินค้า A | 2 | 500 |\n"
}
```

ตัวอย่าง XLSX: ใช้ Markdown table ใน `content` ระบบจะนำตารางไปใส่ใน Excel

ตัวอย่าง PDF: ใช้ Markdown headings, paragraphs และ bullet list

ตัวอย่าง PPTX แบบ Markdown: ใช้ `#` หนึ่งหัวข้อต่อหนึ่งสไลด์ และข้อความ/รายการ bullet ใต้หัวข้อนั้น

ตัวอย่าง PPTX แบบ JSON:

```json
{
  "slides": [
    {"title": "ภาพรวม", "content": "สรุปโครงการ"},
    {"title": "แผนงาน", "bullets": ["ออกแบบ", "พัฒนา", "ทดสอบ"]}
  ]
}
```

Response:

```json
{
  "success": true,
  "filename": "quotation-2026.docx",
  "format": "docx",
  "download_url": "https://YOUR-SERVICE.onrender.com/download/FILE_ID",
  "expires_in_seconds": 86400
}
```

## 5) ข้อจำกัดและการเก็บไฟล์

ค่าเริ่มต้นเก็บไฟล์ใน `/tmp/generated_files` และลิงก์มีอายุ 24 ชั่วโมง ไฟล์บน local/ephemeral disk อาจหายเมื่อ Render รีสตาร์ตหรือ deploy ใหม่ จึงเหมาะกับการทดลองและไฟล์ชั่วคราว

ถ้าต้องการเก็บไฟล์ให้คงอยู่ ให้ใช้ object storage (เช่น S3-compatible storage) หรือ Render Persistent Disk บนแพ็กเกจที่รองรับ แล้วตั้ง `FILE_STORAGE_DIR` เป็น mount path ที่กำหนด ทั้งนี้ Persistent Disk มีค่าใช้จ่ายและอาจมีข้อจำกัดตามประเภท service/plan

## 6) ความปลอดภัยก่อนใช้งานจริง

- API ตัวอย่างนี้ยังไม่มี API key authentication; ควรเพิ่มการยืนยันตัวตนก่อนเปิดให้ใช้งานสาธารณะ
- จำกัดขนาดเนื้อหาได้ด้วย `MAX_CONTENT_CHARS`
- ตั้ง `PUBLIC_BASE_URL` หากต้องการบังคับโดเมน URL ที่ส่งกลับ
- อย่าใส่ข้อมูลลับหรือข้อมูลส่วนบุคคลลงในไฟล์โดยไม่จำเป็น


## Thai language support

This version bundles Noto Sans Thai regular and bold fonts and embeds them in generated PDFs. Send request bodies as UTF-8 JSON. DOCX/XLSX/PPTX support Unicode text; the viewing device should have a Thai-capable font installed.
