import os
import uuid
import cv2
import requests
import numpy as np
from PIL import Image, ImageEnhance
from flask import Flask, request, send_file, render_template, jsonify, make_response
from werkzeug.utils import secure_filename
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import img2pdf
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Helper configurations for Blur-to-Clear Route
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "bmp"}
MAX_FILE_SIZE_MB = 15

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def detect_blur_score(gray_img) -> float:
    """
    Laplacian variance: blur metric. Lower value = blurrier image.
    """
    return cv2.Laplacian(gray_img, cv2.CV_64F).var()

def adaptive_strength(blur_score: float):
    """
    Maps the blur score to sharpening/contrast parameters.
    """
    if blur_score < 15:          # very blurry
        sharpen_amount, unsharp_weight, contrast = 3.0, -1.8, 1.35
    elif blur_score < 50:        # moderately blurry
        sharpen_amount, unsharp_weight, contrast = 2.4, -1.4, 1.25
    elif blur_score < 150:       # mildly soft
        sharpen_amount, unsharp_weight, contrast = 1.8, -0.9, 1.15
    else:                        # already fairly sharp
        sharpen_amount, unsharp_weight, contrast = 1.3, -0.5, 1.05
    return sharpen_amount, unsharp_weight, contrast

# SAFE TEMPLATE RENDERER FUNCTION (Fixes UnicodeDecodeError on Windows)
def safe_render(template_name):
    template_path = os.path.join(app.root_path, 'templates', template_name)
    with open(template_path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()

# ==========================================
# TECHNICAL SEO: DYNAMIC ROBOTS.TXT & SITEMAP
# ==========================================

@app.route('/robots.txt')
def robots_txt():
    content = "User-agent: *\nAllow: /\n\nSitemap: https://likepdf.com/sitemap.xml"
    response = make_response(content)
    response.headers["Content-Type"] = "text/plain"
    return response

@app.route('/sitemap.xml')
def sitemap_xml():
    today = datetime.now().strftime('%Y-%m-%d')
    urls = [
        "", "pdf-to-word", "compress-pdf", "excel-to-pdf", "jpg-to-pdf",
        "merge-pdf", "split-pdf", "full-hd-photo", "ai-layout-fixer",
        "watermark-remover", "auto-sorter", "unlock-pdf", "protect-pdf",
        "rotate-pdf", "resize-image", "compress-image", "convert-image", "blur-to-clear"
    ]
    xml_content = '<?xml version="1.0" encoding="UTF-8"?>\n'
    xml_content += '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
    for url in urls:
        loc = f"https://likepdf.com/{url}" if url else "https://likepdf.com/"
        priority = "1.0" if not url else "0.8"
        xml_content += f'  <url>\n    <loc>{loc}</loc>\n    <lastmod>{today}</lastmod>\n    <changefreq>weekly</changefreq>\n    <priority>{priority}</priority>\n  </url>\n'
    xml_content += '</urlset>'
    response = make_response(xml_content)
    response.headers["Content-Type"] = "application/xml"
    return response


# ==========================================
# PAGE ROUTERS (USES SAFE UTF-8 RENDERER)
# ==========================================
@app.route('/')
def home():
    return safe_render('home.html')

@app.route('/pdf-to-word')
def pdf_to_word_page():
    return safe_render('pdf_to_word.html')

@app.route('/compress-pdf')
def compress_pdf_page():
    return safe_render('compress_pdf.html')

@app.route('/excel-to-pdf')
def excel_to_pdf_page():
    return safe_render('excel_to_pdf.html')

@app.route('/jpg-to-pdf')
def jpg_to_pdf_page():
    return safe_render('jpg_to_pdf.html')

@app.route('/merge-pdf')
def merge_pdf_page():
    return safe_render('merge_pdf.html')

@app.route('/split-pdf')
def split_pdf_page():
    return safe_render('split_pdf.html')

@app.route('/secret-image-tool')
def secret_image_tool_page():
    return safe_render('secret_image_tool.html')

@app.route('/full-hd-photo')
def full_hd_photo_page():
    return safe_render('full_hd_photo.html')

@app.route('/ai-layout-fixer')
def ai_layout_fixer_page():
    return safe_render('ai_layout_fixer.html')

@app.route('/watermark-remover')
def watermark_remover_page():
    return safe_render('watermark_remover.html')

@app.route('/auto-sorter')
def auto_sorter_page():
    return safe_render('auto_sorter.html')

@app.route('/unlock-pdf')
def unlock_pdf_page(): 
    return safe_render('unlock_pdf.html')

@app.route('/protect-pdf')
def protect_pdf_page(): 
    return safe_render('protect_pdf.html')

@app.route('/rotate-pdf')
def rotate_pdf_page():
    return safe_render('rotate_pdf.html')

@app.route('/resize-image')
def resize_image_page(): 
    return safe_render('resize_image.html')

@app.route('/compress-image')
def compress_image_page(): 
    return safe_render('compress_image.html')

@app.route('/convert-image')
def convert_image_page(): 
    return safe_render('convert_image.html')

@app.route('/blur-to-clear')
def blur_to_clear_page():
    return safe_render('blur_to_clear.html')


# ==========================================
# BACKEND API PROCESSING ROUTINES
# ==========================================

# 1. TOOL: PDF TO WORD
@app.route('/api/pdf-to-word', methods=['POST'])
def api_pdf_to_word():
    if 'pdf_file' not in request.files: return "No file uploaded", 400
    file = request.files['pdf_file']
    if file.filename == '': return "No file selected", 400
    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    docx_filename = file.filename.rsplit('.', 1)[0] + '.docx'
    docx_path = os.path.join(OUTPUT_FOLDER, docx_filename)
    file.save(pdf_path)
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        return send_file(docx_path, as_attachment=True)
    except Exception as e: return f"Error processing PDF: {str(e)}", 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

# 2. TOOL: COMPRESS PDF
@app.route('/api/compress-pdf', methods=['POST'])
def api_compress_pdf():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    level = request.form.get('compression_level', '2')
    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    out_path = os.path.join(OUTPUT_FOLDER, 'compressed_' + file.filename)
    file.save(pdf_path)
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        for page in reader.pages:
            if level == '3': page.compress_content_streams() 
            writer.add_page(page)
        with open(out_path, 'wb') as f: writer.write(f)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Compression error: {str(e)}", 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

# 3. TOOL: EXCEL TO PDF
@app.route('/api/excel-to-pdf', methods=['POST'])
def api_excel_to_pdf():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    xlsx_path = os.path.join(UPLOAD_FOLDER, file.filename)
    pdf_filename = file.filename.rsplit('.', 1)[0] + '.pdf'
    pdf_path = os.path.join(OUTPUT_FOLDER, pdf_filename)
    file.save(xlsx_path)
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append([str(cell) if cell is not None else '' for cell in row])
        doc = SimpleDocTemplate(pdf_path, pagesize=letter)
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('GRID', (0,0), (-1,-1), 1, colors.black)
        ]))
        doc.build([table])
        return send_file(pdf_path, as_attachment=True)
    except Exception as e: return f"Excel Conversion error: {str(e)}", 500
    finally:
        if os.path.exists(xlsx_path): os.remove(xlsx_path)

# 4. TOOL: JPG TO PDF
@app.route('/api/jpg-to-pdf', methods=['POST'])
def api_jpg_to_pdf():
    uploaded_files = request.files.getlist('files')
    if not uploaded_files or uploaded_files[0].filename == '': return "No files selected", 400
    img_paths = []
    pdf_path = os.path.join(OUTPUT_FOLDER, 'images_combined.pdf')
    try:
        for file in uploaded_files:
            p = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(p)
            img_paths.append(p)
        with open(pdf_path, 'wb') as f: f.write(img2pdf.convert(img_paths))
        return send_file(pdf_path, as_attachment=True)
    except Exception as e: return f"Image Conversion error: {str(e)}", 500
    finally:
        for p in img_paths:
            if os.path.exists(p): os.remove(p)

# 5. TOOL: MERGE PDF
@app.route('/api/merge-pdf', methods=['POST'])
def api_merge_pdf():
    uploaded_files = request.files.getlist('files')
    if not uploaded_files or uploaded_files[0].filename == '': return "No files selected", 400
    writer = PdfWriter()
    temp_paths = []
    out_path = os.path.join(OUTPUT_FOLDER, 'merged_document.pdf')
    try:
        for file in uploaded_files:
            p = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(p)
            temp_paths.append(p)
            writer.append(p)
        with open(out_path, 'wb') as f: writer.write(f)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Merging error: {str(e)}", 500
    finally:
        for p in temp_paths:
            if os.path.exists(p): os.remove(p)

# 6. TOOL: SPLIT PDF
@app.route('/api/split-pdf', methods=['POST'])
def api_split_pdf():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    method = request.form.get('split_method', 'custom')
    ranges = request.form.get('page_ranges', '')
    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    out_path = os.path.join(OUTPUT_FOLDER, 'split_' + file.filename)
    file.save(pdf_path)
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        total_pages = len(reader.pages)
        if method == 'custom' and ranges:
            for part in ranges.split(','):
                if Part := part.strip():
                    if '-' in Part:
                        start, end = map(int, Part.split('-'))
                        for idx in range(start-1, min(end, total_pages)): writer.add_page(reader.pages[idx])
                    else:
                        idx = int(Part) - 1
                        if 0 <= idx < total_pages: writer.add_page(reader.pages[idx])
        else:
            writer.add_page(reader.pages[0])
        with open(out_path, 'wb') as f: writer.write(f)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Splitting error: {str(e)}", 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

import io
from flask import send_file, jsonify

# 7. TOOL: BLUR TO CLEAR AI (MEMORY-BUFFER ENGINE - 0% PERMISSION ERRORS)
@app.route('/api/blur-to-clear', methods=['POST'])
def api_blur_to_clear():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    try:
        # 1. Read file directly into RAM memory buffer (No disk lock issues)
        file_bytes = file.read()
        nparr = np.frombuffer(file_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

        if img is None:
            return jsonify({"error": "Invalid image file format"}), 400

        # 2. Stage 1: CLAHE (Contrast Limited Adaptive Histogram Equalization)
        # This recovers hidden details from blurry regions
        lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
        l, a, b = cv2.split(lab)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        l_clahe = clahe.apply(l)
        enhanced_bgr = cv2.cvtColor(cv2.merge((l_clahe, a, b)), cv2.COLOR_LAB2BGR)

        # 3. Stage 2: Bilateral Denoising + Unsharp Sharpness Mask
        denoised = cv2.bilateralFilter(enhanced_bgr, 7, 50, 50)
        gaussian = cv2.GaussianBlur(denoised, (0, 0), 2.0)
        sharpened = cv2.addWeighted(denoised, 2.2, gaussian, -1.2, 0)

        # 4. Stage 3: PIL Sharpness & Contrast Boost
        img_rgb = cv2.cvtColor(sharpened, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(img_rgb)

        sharp_enhancer = ImageEnhance.Sharpness(pil_img)
        pil_img = sharp_enhancer.enhance(2.2)

        contrast_enhancer = ImageEnhance.Contrast(pil_img)
        pil_img = contrast_enhancer.enhance(1.2)

        # 5. Convert processed image back to Byte Stream (No file saving to disk needed)
        output_buffer = io.BytesIO()
        pil_img.save(output_buffer, format='JPEG', quality=95)
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='image/jpeg',
            as_attachment=True,
            download_name='cleared_image.jpg'
        )

    except Exception as e:
        app.logger.error(f"Blur to Clear Error: {str(e)}")
        return jsonify({"error": f"Image enhancement failed: {str(e)}"}), 500

# 8. TOOL: FULL HD PHOTO UPSCALER
@app.route('/api/full-hd-photo', methods=['POST'])
def api_full_hd_photo():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    img_path = os.path.join(UPLOAD_FOLDER, file.filename)
    out_path = os.path.join(OUTPUT_FOLDER, 'hd_' + file.filename)
    file.save(img_path)
    try:
        with Image.open(img_path) as img:
            width, height = img.size
            hd_img = img.resize((width * 2, height * 2), Image.Resampling.LANCZOS)
            enhancer = ImageEnhance.Sharpness(hd_img)
            hd_img = enhancer.enhance(1.2)
            color_enhancer = ImageEnhance.Color(hd_img)
            hd_img = color_enhancer.enhance(1.05)
            hd_img.save(out_path)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Upscaling error: {str(e)}", 500
    finally:
        if os.path.exists(img_path): os.remove(img_path)

# 9. TOOL: AI RESUME LAYOUT FIXER
@app.route('/api/ai-layout-fixer', methods=['POST'])
def api_ai_layout_fixer():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    out_path = os.path.join(OUTPUT_FOLDER, 'fixed_' + file.filename)
    file.save(pdf_path)
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        for page in reader.pages:
            page.compress_content_streams() 
            writer.add_page(page)
        with open(out_path, 'wb') as f: writer.write(f)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Error: {str(e)}", 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

# 10. TOOL: WATERMARK REMOVER
@app.route('/api/watermark-remover', methods=['POST'])
def api_watermark_remover():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    out_path = os.path.join(OUTPUT_FOLDER, 'clean_' + file.filename)
    file.save(pdf_path)
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        for page in reader.pages:
            if "/Annots" in page: del page["/Annots"]
            writer.add_page(page)
        with open(out_path, 'wb') as f: writer.write(f)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Error: {str(e)}", 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

# 11. TOOL: SMART PAGE SORTER
@app.route('/api/auto-sorter', methods=['POST'])
def api_auto_sorter():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    out_path = os.path.join(OUTPUT_FOLDER, 'sorted_' + file.filename)
    file.save(pdf_path)
    try:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        for page in reversed(reader.pages): writer.add_page(page)
        with open(out_path, 'wb') as f: writer.write(f)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Error: {str(e)}", 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

# 12. API: PROTECT PDF
@app.route('/api/protect-pdf', methods=['POST'])
def api_protect_pdf():
    file = request.files['file']
    password = request.form.get('password')
    in_p = os.path.join(UPLOAD_FOLDER, file.filename)
    out_p = os.path.join(OUTPUT_FOLDER, 'protected_' + file.filename)
    file.save(in_p)
    try:
        r = PdfReader(in_p)
        w = PdfWriter()
        for p in r.pages: w.add_page(p)
        w.encrypt(password)
        with open(out_p, 'wb') as f: w.write(f)
        return send_file(out_p, as_attachment=True)
    except: return "Security processing failed", 500
    finally: os.remove(in_p)

# 13. API: UNLOCK PDF
@app.route('/api/unlock-pdf', methods=['POST'])
def api_unlock_pdf():
    file = request.files['file']
    password = request.form.get('password')
    in_p = os.path.join(UPLOAD_FOLDER, file.filename)
    out_p = os.path.join(OUTPUT_FOLDER, 'unlocked_' + file.filename)
    file.save(in_p)
    try:
        r = PdfReader(in_p)
        if r.is_encrypted: r.decrypt(password)
        w = PdfWriter()
        for p in r.pages: w.add_page(p)
        with open(out_p, 'wb') as f: w.write(f)
        return send_file(out_p, as_attachment=True)
    except: return "Decryption failed. Verify key.", 400
    finally: os.remove(in_p)

# 14. API: ROTATE PDF
@app.route('/api/rotate-pdf', methods=['POST'])
def api_rotate_pdf():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    angle = int(request.form.get('angle', 90))
    in_p = os.path.join(UPLOAD_FOLDER, file.filename)
    out_p = os.path.join(OUTPUT_FOLDER, 'rotated_' + file.filename)
    file.save(in_p)
    try:
        r = PdfReader(in_p)
        w = PdfWriter()
        for p in r.pages: w.add_page(p.rotate(angle))
        with open(out_p, 'wb') as f: w.write(f)
        return send_file(out_p, mimetype='application/pdf', as_attachment=True)
    except: return "Matrix rotation layout anomaly", 500
    finally: os.remove(in_p)

# 15. API: RESIZE IMAGE
@app.route('/api/resize-image', methods=['POST'])
def api_resize_image():
    file = request.files['file']
    w = int(request.form.get('width', 1920))
    h = int(request.form.get('height', 1080))
    in_p = os.path.join(UPLOAD_FOLDER, file.filename)
    out_p = os.path.join(OUTPUT_FOLDER, 'resized_' + file.filename)
    file.save(in_p)
    try:
        with Image.open(in_p) as img:
            resized = img.resize((w, h), Image.Resampling.LANCZOS)
            resized.save(out_p, quality=95)
        return send_file(out_p, as_attachment=True)
    except: return "Spatial adjustments failed", 500
    finally: os.remove(in_p)

# 16. API: COMPRESS IMAGE
@app.route('/api/compress-image', methods=['POST'])
def api_compress_image():
    file = request.files['file']
    q = int(request.form.get('quality', 80))
    in_p = os.path.join(UPLOAD_FOLDER, file.filename)
    out_p = os.path.join(OUTPUT_FOLDER, 'compressed_' + file.filename)
    file.save(in_p)
    try:
        with Image.open(in_p) as img: img.save(out_p, optimize=True, quality=q)
        return send_file(out_p, as_attachment=True)
    except: return "Footprint reduction dropped", 500
    finally: os.remove(in_p)

# 17. API: CONVERT IMAGE
@app.route('/api/convert-image', methods=['POST'])
def api_convert_image():
    file = request.files['file']
    fmt = request.form.get('target_format', 'PNG')
    in_p = os.path.join(UPLOAD_FOLDER, file.filename)
    out_filename = file.filename.rsplit('.', 1)[0] + f'.{fmt.lower()}'
    out_p = os.path.join(OUTPUT_FOLDER, out_filename)
    file.save(in_p)
    try:
        with Image.open(in_p) as img:
            if img.mode in ('RGBA', 'LA') and fmt == 'JPEG': img = img.convert('RGB')
            img.save(out_p, fmt, quality=95)
        return send_file(out_p, as_attachment=True)
    except: return "Transcoding stream dropped", 500
    finally: os.remove(in_p)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
