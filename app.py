import os
import io
import uuid
import shutil
import base64
import subprocess
import cv2
import numpy as np
from PIL import Image, ImageEnhance
from flask import Flask, request, jsonify, send_file, make_response, send_from_directory, after_this_request
from werkzeug.utils import secure_filename
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import img2pdf
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from datetime import datetime
import fitz  # PyMuPDF
import pytesseract

app = Flask(__name__)

# Base working directories
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
BASE_TEMP_DIR = os.path.join(os.getcwd(), "editor_sessions")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(BASE_TEMP_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "docx", "xlsx", "pptx", "jpg", "jpeg", "png", "webp", "bmp"}
MAX_FILE_SIZE_MB = 35

# Resolve Windows Tesseract binary location automatically if present
if os.name == 'nt':
    standard_tesseract_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe")
    ]
    for t_path in standard_tesseract_paths:
        if os.path.exists(t_path):
            pytesseract.pytesseract.tesseract_cmd = t_path
            break

# Resolve Unicode-compatible font (Linux vs Windows fallback)
def get_unicode_font_path():
    possible_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
        "C:\\Windows\\Fonts\\segoeui.ttf"
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None

UNICODE_FONT_PATH = get_unicode_font_path()


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_render(template_name: str) -> str:
    template_path = os.path.join(app.root_path, 'templates', template_name)
    with open(template_path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()


def convert_to_pdf_libreoffice(input_path: str, output_dir: str) -> str:
    """Converts DOCX, XLSX, PPTX, or image formats to standard PDF via headless LibreOffice."""
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        input_path
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr.decode('utf-8', errors='ignore')}")

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    expected_pdf = os.path.join(output_dir, f"{base_name}.pdf")
    if not os.path.exists(expected_pdf):
        raise FileNotFoundError("Generated PDF was not produced by LibreOffice.")
    return expected_pdf


def is_scanned_page(page: fitz.Page) -> bool:
    """Checks if a page contains no extractable digital text."""
    return len(page.get_text("text").strip()) == 0


def run_ocr_on_page(page: fitz.Page) -> list:
    """Safely attempts OCR; returns an empty list without crashing if Tesseract is not installed."""
    try:
        zoom = 300 / 72
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        n_boxes = len(ocr_data['level'])
        text_blocks = []

        for i in range(n_boxes):
            text = ocr_data['text'][i].strip()
            if not text:
                continue

            x0 = float(ocr_data['left'][i]) / zoom
            y0 = float(ocr_data['top'][i]) / zoom
            w = float(ocr_data['width'][i]) / zoom
            h = float(ocr_data['height'][i]) / zoom

            text_blocks.append({
                "block_id": uuid.uuid4().hex[:12],
                "text": text,
                "bbox": [x0, y0, x0 + w, y0 + h],
                "font_size": max(h * 0.8, 8.0),
                "font": "OCR-Detected",
                "color": 0
            })
        return text_blocks
    except Exception as e:
        app.logger.warning(f"OCR bypass triggered (Tesseract unavailable): {e}")
        return []


# ==========================================
# FAVICON & SEO ROUTES
# ==========================================

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'), 'favicon.png', mimetype='image/png')


@app.route('/robots.txt')
def robots_txt():
    content = "User-agent: *\nAllow: /\n\nSitemap: https://likepdf.in/sitemap.xml"
    response = make_response(content)
    response.headers["Content-Type"] = "text/plain"
    return response


@app.route('/sitemap.xml')
def sitemap_xml():
    today = datetime.now().strftime('%Y-%m-%d')
    pages = [
        {"url": "", "priority": "1.0", "change": "daily"},
        {"url": "edit-pdf", "priority": "0.9", "change": "weekly"},
        {"url": "compress-pdf", "priority": "0.9", "change": "weekly"},
        {"url": "pdf-to-word", "priority": "0.9", "change": "weekly"},
        {"url": "merge-pdf", "priority": "0.9", "change": "weekly"},
        {"url": "split-pdf", "priority": "0.8", "change": "weekly"},
        {"url": "blur-to-clear", "priority": "0.9", "change": "weekly"},
        {"url": "full-hd-photo", "priority": "0.8", "change": "weekly"},
        {"url": "excel-to-pdf", "priority": "0.8", "change": "weekly"},
        {"url": "jpg-to-pdf", "priority": "0.8", "change": "weekly"},
        {"url": "unlock-pdf", "priority": "0.8", "change": "weekly"},
        {"url": "unlock-pdf-without-password", "priority": "0.8", "change": "weekly"},
        {"url": "protect-pdf", "priority": "0.8", "change": "weekly"},
        {"url": "watermark-remover", "priority": "0.7", "change": "weekly"},
        {"url": "rotate-pdf", "priority": "0.7", "change": "weekly"},
        {"url": "resize-image", "priority": "0.7", "change": "weekly"},
        {"url": "compress-image", "priority": "0.7", "change": "weekly"},
        {"url": "convert-image", "priority": "0.7", "change": "weekly"}
    ]
    xml_content = '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
    for page in pages:
        loc = f"https://likepdf.in/{page['url']}" if page['url'] else "https://likepdf.in/"
        xml_content += f'  <url>\n    <loc>{loc}</loc>\n    <lastmod>{today}</lastmod>\n    <changefreq>{page["change"]}</changefreq>\n    <priority>{page["priority"]}</priority>\n  </url>\n'
    xml_content += '</urlset>'
    response = make_response(xml_content)
    response.headers["Content-Type"] = "application/xml"
    return response


@app.route('/google53035f8b66856dae.html')
def google_verification():
    return send_from_directory(os.path.join(app.root_path, 'static'), 'google53035f8b66856dae.html')


# ==========================================
# PAGE ROUTERS (ALL TOOLS INCLUDED)
# ==========================================

@app.route('/')
def home(): return safe_render('home.html')

@app.route('/edit-pdf')
def edit_pdf_page(): return safe_render('edit_pdf.html')

@app.route('/unlock-pdf-without-password')
def unlock_without_password_page(): return safe_render('unlock_without_password.html')

@app.route('/unlock-pdf')
def unlock_pdf_page(): return safe_render('unlock_pdf.html')

@app.route('/pdf-to-word')
def pdf_to_word_page(): return safe_render('pdf_to_word.html')

@app.route('/compress-pdf')
def compress_pdf_page(): return safe_render('compress_pdf.html')

@app.route('/excel-to-pdf')
def excel_to_pdf_page(): return safe_render('excel_to_pdf.html')

@app.route('/jpg-to-pdf')
def jpg_to_pdf_page(): return safe_render('jpg_to_pdf.html')

@app.route('/merge-pdf')
def merge_pdf_page(): return safe_render('merge_pdf.html')

@app.route('/split-pdf')
def split_pdf_page(): return safe_render('split_pdf.html')

@app.route('/blur-to-clear')
def blur_to_clear_page(): return safe_render('blur_to_clear.html')

@app.route('/full-hd-photo')
def full_hd_photo_page(): return safe_render('full_hd_photo.html')

@app.route('/protect-pdf')
def protect_pdf_page(): return safe_render('protect_pdf.html')

@app.route('/watermark-remover')
def watermark_remover_page(): return safe_render('watermark_remover.html')

@app.route('/rotate-pdf')
def rotate_pdf_page(): return safe_render('rotate_pdf.html')

@app.route('/resize-image')
def resize_image_page(): return safe_render('resize_image.html')

@app.route('/compress-image')
def compress_image_page(): return safe_render('compress_image.html')

@app.route('/convert-image')
def convert_image_page(): return safe_render('convert_image.html')

@app.route('/privacy-policy')
def privacy_policy_page(): return safe_render('privacy_policy.html')

@app.route('/terms-of-service')
def terms_of_service_page(): return safe_render('terms_of_service.html')

@app.route('/contact')
def contact_page(): return safe_render('contact.html')


# ===========================================================================
# 1. PREPARE EDIT: Multi-Format Upload -> Normalization -> Spans
# ===========================================================================

@app.route("/api/pdf/prepare-edit", methods=["POST"])
def prepare_edit():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "" or not allowed_file(file.filename):
        return jsonify({"error": "Unsupported file format."}), 400

    session_id = uuid.uuid4().hex
    session_dir = os.path.join(BASE_TEMP_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    filename = secure_filename(file.filename)
    raw_path = os.path.join(session_dir, filename)
    file.save(raw_path)

    ext = filename.rsplit(".", 1)[1].lower()
    pdf_path = os.path.join(session_dir, "working_document.pdf")

    try:
        if ext == "pdf":
            shutil.copy(raw_path, pdf_path)
        else:
            converted_path = convert_to_pdf_libreoffice(raw_path, session_dir)
            shutil.move(converted_path, pdf_path)

        doc = fitz.open(pdf_path)
        pages_data = []

        for page_index in range(len(doc)):
            page = doc[page_index]

            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_b64 = base64.b64encode(pix.tobytes("png")).decode("utf-8")

            text_blocks = []

            if is_scanned_page(page):
                text_blocks = run_ocr_on_page(page)
            else:
                page_dict = page.get_text("dict")
                for block in page_dict.get("blocks", []):
                    if block.get("type") != 0:
                        continue
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            clean_text = span.get("text", "")
                            if not clean_text.strip():
                                continue
                            text_blocks.append({
                                "block_id": uuid.uuid4().hex[:12],
                                "text": clean_text,
                                "bbox": span["bbox"],
                                "font_size": span["size"],
                                "font": span.get("font", "Arial"),
                                "color": span.get("color", 0)
                            })

            pages_data.append({
                "page_number": page_index,
                "width": page.rect.width,
                "height": page.rect.height,
                "image_base64": img_b64,
                "text_blocks": text_blocks
            })

        doc.close()

        return jsonify({
            "session_id": session_id,
            "page_count": len(pages_data),
            "pages": pages_data
        })

    except Exception as e:
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({"error": f"Failed to prepare document: {str(e)}"}), 500


# ===========================================================================
# 2. ADD BLANK PAGE ROUTE
# ===========================================================================

@app.route("/api/pdf/add-page", methods=["POST"])
def add_new_page():
    data = request.get_json(silent=True) or {}
    session_id = secure_filename(str(data.get("session_id", "")))
    session_dir = os.path.join(BASE_TEMP_DIR, session_id)
    pdf_path = os.path.join(session_dir, "working_document.pdf")

    if not os.path.exists(pdf_path):
        return jsonify({"error": "Session expired"}), 404

    try:
        doc = fitz.open(pdf_path)
        new_page = doc.new_page(width=595, height=842)
        pix = new_page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_b64 = base64.b64encode(pix.tobytes("png")).decode("utf-8")
        page_num = len(doc) - 1
        doc.save(pdf_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        doc.close()

        return jsonify({
            "success": True,
            "page": {
                "page_number": page_num,
                "width": 595,
                "height": 842,
                "image_base64": img_b64,
                "text_blocks": []
            }
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===========================================================================
# 3. APPLY EDITS: Tight Boundary Math & Clean PDF Export
# ===========================================================================

@app.route("/api/pdf/apply-edit", methods=["POST"])
def apply_edit():
    data = request.get_json(silent=True)
    if not data or "session_id" not in data:
        return jsonify({"error": "Missing session_id or edit payload."}), 400

    session_id = secure_filename(str(data["session_id"]))
    session_dir = os.path.join(BASE_TEMP_DIR, session_id)
    pdf_path = os.path.join(session_dir, "working_document.pdf")

    if not os.path.exists(pdf_path):
        return jsonify({"error": "Session expired or not found. Please re-upload your document."}), 404

    def parse_hex_color(col):
        if isinstance(col, str) and col.startswith("#"):
            col = col.lstrip("#")
            return tuple(int(col[i:i+2], 16) / 255.0 for i in (0, 2, 4))
        elif isinstance(col, int):
            r = ((col >> 16) & 255) / 255.0
            g = ((col >> 8) & 255) / 255.0
            b = (col & 255) / 255.0
            return (r, g, b)
        return (0.0, 0.0, 0.0)

    try:
        doc = fitz.open(pdf_path)

        # 1. Apply Redactions with strict vertical bounds (prevents merging below)
        for edit in data.get("edits", []):
            page_idx = int(edit.get("page_number", 0))
            bbox = edit.get("bbox")
            if bbox and page_idx < len(doc):
                page = doc[page_idx]
                y0 = float(bbox[1])
                y1 = float(bbox[3])
                f_size = float(edit.get("font_size", 10))
                rect = fitz.Rect(float(bbox[0]), y0, float(bbox[2]), min(y1, y0 + f_size * 1.05))
                page.add_redact_annot(rect, fill=(1, 1, 1))

        # Commit redactions across pages
        for p in doc:
            p.apply_redactions()

        # 2. Draw Replacement Text (Strict Typographical Baseline)
        for edit in data.get("edits", []):
            page_idx = int(edit.get("page_number", 0))
            bbox = edit.get("bbox")
            new_text = str(edit.get("new_text", "")).strip()
            font_size = float(edit.get("font_size", 10.0))
            color = parse_hex_color(edit.get("color", 0))

            if not new_text or bbox is None or page_idx >= len(doc):
                continue

            page = doc[page_idx]
            x0 = float(bbox[0])
            y0 = float(bbox[1])
            baseline_y = y0 + (font_size * 0.82)

            font_kw = {"fontfile": UNICODE_FONT_PATH} if (UNICODE_FONT_PATH and os.path.exists(UNICODE_FONT_PATH)) else {"fontname": "helv"}
            page.insert_text(
                fitz.Point(x0, baseline_y),
                new_text,
                fontsize=font_size,
                color=color,
                **font_kw
            )

        # 3. Add Custom Free-Placed Text Additions
        for item in data.get("new_texts", []):
            page_idx = int(item.get("page_number", 0))
            if page_idx >= len(doc):
                continue

            page = doc[page_idx]
            x = float(item.get("x", 50))
            y = float(item.get("y", 50))
            txt = str(item.get("text", "")).strip()
            size = float(item.get("font_size", 10.0))
            color = parse_hex_color(item.get("color", "#000000"))

            baseline_y = y + (size * 0.82)

            font_kw = {"fontfile": UNICODE_FONT_PATH} if (UNICODE_FONT_PATH and os.path.exists(UNICODE_FONT_PATH)) else {"fontname": "helv"}
            page.insert_text(fitz.Point(x, baseline_y), txt, fontsize=size, color=color, **font_kw)

        # 4. Insert Images & Signatures
        combined_media = data.get("images", []) + data.get("signatures", [])
        for media in combined_media:
            page_idx = int(media.get("page_number", 0))
            bbox = media.get("bbox")
            b64_data = media.get("image_base64", "")

            if not bbox or not b64_data or page_idx >= len(doc):
                continue

            if "," in b64_data:
                b64_data = b64_data.split(",")[1]

            image_bytes = base64.b64decode(b64_data)
            page = doc[page_idx]
            rect = fitz.Rect(float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3]))
            page.insert_image(rect, stream=image_bytes)

        output_pdf_path = os.path.join(session_dir, "final_edited_document.pdf")
        doc.save(output_pdf_path, garbage=4, deflate=True)
        doc.close()

        @after_this_request
        def cleanup(response):
            try:
                shutil.rmtree(session_dir, ignore_errors=True)
            except Exception as ex:
                app.logger.error(f"Error purging session: {ex}")
            return response

        return send_file(
            output_pdf_path,
            as_attachment=True,
            download_name="edited_document.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        app.logger.error(f"Error compiling output PDF: {e}")
        return jsonify({"error": f"Failed to apply modifications: {str(e)}"}), 500


# ===========================================================================
# 4. ALL CONVERSION & UTILITY API ENDPOINTS
# ===========================================================================

@app.route('/api/unlock-without-password', methods=['POST'])
def api_unlock_without_password():
    file = request.files.get('file')
    if not file: 
        return jsonify({"error": "No file uploaded"}), 400
    
    try:
        file_bytes = file.read()
        reader = PdfReader(io.BytesIO(file_bytes))
        writer = PdfWriter()

        if reader.is_encrypted:
            decrypted = False
            # List of common automated fallback keys or empty passwords
            fallback_passwords = ["", " ", "1234", "0000"]
            
            for k in fallback_passwords:
                try:
                    if reader.decrypt(k) > 0: 
                        decrypted = True
                        break
                except Exception:
                    pass
            
            # If standard fallback keys fail, check if user sent a password via form request
            if not decrypted:
                form_password = request.form.get('password', '')
                if form_password:
                    try:
                        if reader.decrypt(form_password) > 0:
                            decrypted = True
                    except Exception:
                        pass

            if not decrypted:
                return jsonify({
                    "error": "This PDF is encrypted with a user password (such as a salary slip or bank statement security lock). Please use the standard 'Unlock PDF' tool with the correct password."
                }), 400

        for p in reader.pages: 
            writer.add_page(p)
            
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        
        return send_file(
            buf, 
            mimetype='application/pdf', 
            as_attachment=True, 
            download_name=f"unlocked_{secure_filename(file.filename)}"
        )
    except Exception as e: 
        return jsonify({"error": str(e)}), 500

@app.route('/api/unlock-pdf', methods=['POST'])
def api_unlock_pdf():
    file = request.files.get('file')
    pw = request.form.get('password')
    try:
        r = PdfReader(file)
        if r.is_encrypted: r.decrypt(pw)
        w = PdfWriter()
        for p in r.pages: w.add_page(p)
        buf = io.BytesIO()
        w.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name='unlocked.pdf')
    except Exception as e: return jsonify({"error": str(e)}), 400


@app.route('/api/protect-pdf', methods=['POST'])
def api_protect_pdf():
    file = request.files.get('file')
    pw = request.form.get('password')
    try:
        r = PdfReader(file)
        w = PdfWriter()
        for p in r.pages: w.add_page(p)
        w.encrypt(pw)
        buf = io.BytesIO()
        w.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name='protected.pdf')
    except Exception as e: return jsonify({"error": str(e)}), 500


@app.route('/api/pdf-to-word', methods=['POST'])
def api_pdf_to_word():
    file = request.files.get('pdf_file')
    if not file: return "No file uploaded", 400
    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    docx_path = os.path.join(OUTPUT_FOLDER, file.filename.rsplit('.', 1)[0] + '.docx')
    file.save(pdf_path)
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        return send_file(docx_path, as_attachment=True)
    except Exception as e: return str(e), 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)


@app.route('/api/compress-pdf', methods=['POST'])
def api_compress_pdf():
    file = request.files.get('file')
    level = request.form.get('compression_level', '2')
    try:
        reader = PdfReader(file)
        writer = PdfWriter()
        for p in reader.pages:
            if level == '3': p.compress_content_streams()
            writer.add_page(p)
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name=f"compressed_{file.filename}")
    except Exception as e: return str(e), 500


@app.route('/api/excel-to-pdf', methods=['POST'])
def api_excel_to_pdf():
    file = request.files.get('file')
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        data = [[str(c) if c is not None else '' for c in row] for row in ws.iter_rows(values_only=True)]
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=letter)
        table = Table(data)
        table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black)]))
        doc.build([table])
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name="converted.pdf")
    except Exception as e: return str(e), 500


@app.route('/api/jpg-to-pdf', methods=['POST'])
def api_jpg_to_pdf():
    uploaded_files = request.files.getlist('files')
    img_paths = []
    pdf_path = os.path.join(OUTPUT_FOLDER, 'images_combined.pdf')
    try:
        for f in uploaded_files:
            p = os.path.join(UPLOAD_FOLDER, f.filename)
            f.save(p)
            img_paths.append(p)
        with open(pdf_path, 'wb') as f: f.write(img2pdf.convert(img_paths))
        return send_file(pdf_path, as_attachment=True)
    except Exception as e: return str(e), 500
    finally:
        for p in img_paths:
            if os.path.exists(p): os.remove(p)


@app.route('/api/merge-pdf', methods=['POST'])
def api_merge_pdf():
    uploaded_files = request.files.getlist('files')
    writer = PdfWriter()
    try:
        for f in uploaded_files:
            reader = PdfReader(f)
            for p in reader.pages: writer.add_page(p)
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name="merged.pdf")
    except Exception as e: return str(e), 500


@app.route('/api/split-pdf', methods=['POST'])
def api_split_pdf():
    file = request.files.get('file')
    ranges = request.form.get('page_ranges', '')
    try:
        reader = PdfReader(file)
        writer = PdfWriter()
        total = len(reader.pages)
        if ranges:
            for part in ranges.split(','):
                part = part.strip()
                if '-' in part:
                    s, e = map(int, part.split('-'))
                    for i in range(s-1, min(e, total)): writer.add_page(reader.pages[i])
                else:
                    i = int(part) - 1
                    if 0 <= i < total: writer.add_page(reader.pages[i])
        else:
            writer.add_page(reader.pages[0])
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name="split.pdf")
    except Exception as e: return str(e), 500


@app.route('/api/blur-to-clear', methods=['POST'])
def api_blur_to_clear():
    file = request.files.get('file')
    try:
        nparr = np.frombuffer(file.read(), np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
        l, a, b = cv2.split(lab)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        l_clahe = clahe.apply(l)
        enhanced_bgr = cv2.cvtColor(cv2.merge((l_clahe, a, b)), cv2.COLOR_LAB2BGR)
        denoised = cv2.bilateralFilter(enhanced_bgr, 7, 50, 50)
        gaussian = cv2.GaussianBlur(denoised, (0, 0), 2.0)
        sharpened = cv2.addWeighted(denoised, 2.2, gaussian, -1.2, 0)
        pil_img = Image.fromarray(cv2.cvtColor(sharpened, cv2.COLOR_BGR2RGB))
        pil_img = ImageEnhance.Sharpness(pil_img).enhance(2.2)
        pil_img = ImageEnhance.Contrast(pil_img).enhance(1.2)
        buf = io.BytesIO()
        pil_img.save(buf, format='JPEG', quality=95)
        buf.seek(0)
        return send_file(buf, mimetype='image/jpeg', as_attachment=True, download_name='cleared.jpg')
    except Exception as e: return jsonify({"error": str(e)}), 500


@app.route('/api/full-hd-photo', methods=['POST'])
def api_full_hd_photo():
    file = request.files.get('file')
    try:
        with Image.open(file) as img:
            w, h = img.size
            hd_img = img.resize((w * 2, h * 2), Image.Resampling.LANCZOS)
            hd_img = ImageEnhance.Sharpness(hd_img).enhance(1.2)
            buf = io.BytesIO()
            hd_img.save(buf, format='JPEG', quality=95)
            buf.seek(0)
            return send_file(buf, mimetype='image/jpeg', as_attachment=True, download_name='hd.jpg')
    except Exception as e: return jsonify({"error": str(e)}), 500


@app.route('/api/watermark-remover', methods=['POST'])
def api_watermark_remover():
    file = request.files.get('file')
    try:
        reader = PdfReader(file)
        writer = PdfWriter()
        for page in reader.pages:
            if "/Annots" in page: del page["/Annots"]
            writer.add_page(page)
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name='clean.pdf')
    except Exception as e: return jsonify({"error": str(e)}), 500


@app.route('/api/rotate-pdf', methods=['POST'])
def api_rotate_pdf():
    file = request.files.get('file')
    angle = int(request.form.get('angle', 90))
    try:
        r = PdfReader(file)
        w = PdfWriter()
        for p in r.pages: w.add_page(p.rotate(angle))
        buf = io.BytesIO()
        w.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name='rotated.pdf')
    except Exception as e: return jsonify({"error": str(e)}), 500


@app.route('/api/resize-image', methods=['POST'])
def api_resize_image():
    file = request.files.get('file')
    w = int(request.form.get('width', 1920))
    h = int(request.form.get('height', 1080))
    try:
        with Image.open(file) as img:
            res = img.resize((w, h), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            res.save(buf, format='JPEG', quality=95)
            buf.seek(0)
            return send_file(buf, mimetype='image/jpeg', as_attachment=True, download_name='resized.jpg')
    except Exception as e: return jsonify({"error": str(e)}), 500


@app.route('/api/compress-image', methods=['POST'])
def api_compress_image():
    file = request.files.get('file')
    q = int(request.form.get('quality', 80))
    try:
        with Image.open(file) as img:
            buf = io.BytesIO()
            img.save(buf, format='JPEG', optimize=True, quality=q)
            buf.seek(0)
            return send_file(buf, mimetype='image/jpeg', as_attachment=True, download_name='compressed.jpg')
    except Exception as e: return jsonify({"error": str(e)}), 500


@app.route('/api/convert-image', methods=['POST'])
def api_convert_image():
    file = request.files.get('file')
    fmt = request.form.get('target_format', 'PNG')
    try:
        with Image.open(file) as img:
            if img.mode in ('RGBA', 'LA') and fmt == 'JPEG': img = img.convert('RGB')
            buf = io.BytesIO()
            img.save(buf, format=fmt, quality=95)
            buf.seek(0)
            return send_file(buf, mimetype=f'image/{fmt.lower()}', as_attachment=True, download_name=f'converted.{fmt.lower()}')
    except Exception as e: return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
