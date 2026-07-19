import os
import cv2
import numpy as np
from PIL import Image, ImageEnhance
from flask import Flask, request, send_file, render_template, jsonify, make_response
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
# PAGE ROUTERS (FRONTEND RENDERING)
# ==========================================
@app.route('/')
def home():
    return render_template('home.html')

@app.route('/pdf-to-word')
def pdf_to_word_page():
    return render_template('pdf_to_word.html')

@app.route('/compress-pdf')
def compress_pdf_page():
    return render_template('compress_pdf.html')

@app.route('/excel-to-pdf')
def excel_to_pdf_page():
    return render_template('excel_to_pdf.html')

@app.route('/jpg-to-pdf')
def jpg_to_pdf_page():
    return render_template('jpg_to_pdf.html')

@app.route('/merge-pdf')
def merge_pdf_page():
    return render_template('merge_pdf.html')

@app.route('/split-pdf')
def split_pdf_page():
    return render_template('split_pdf.html')

@app.route('/secret-image-tool')
def secret_image_tool_page():
    return render_template('secret_image_tool.html')

@app.route('/full-hd-photo')
def full_hd_photo_page():
    return render_template('full_hd_photo.html')

@app.route('/ai-layout-fixer')
def ai_layout_fixer_page():
    return render_template('ai_layout_fixer.html')

@app.route('/watermark-remover')
def watermark_remover_page():
    return render_template('watermark_remover.html')

@app.route('/auto-sorter')
def auto_sorter_page():
    return render_template('auto_sorter.html')

@app.route('/unlock-pdf')
def unlock_pdf_page(): 
    return render_template('unlock_pdf.html')

@app.route('/protect-pdf')
def protect_pdf_page(): 
    return render_template('protect_pdf.html')

# FIXED: ???? ?? ???? ???? ???? ??????? ???? ???????? ????? ?? ??????? ?? ???? ??
@app.route('/rotate-pdf')
def rotate_pdf_page():
    return render_template('rotate_pdf.html')

@app.route('/resize-image')
def resize_image_page(): 
    return render_template('resize_image.html')

@app.route('/compress-image')
def compress_image_page(): 
    return render_template('compress_image.html')

@app.route('/convert-image')
def convert_image_page(): 
    return render_template('convert_image.html')

@app.route('/blur-to-clear')
def blur_to_clear_page():
    return render_template('blur_to_clear.html')


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
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    for idx in range(start-1, min(end, total_pages)): writer.add_page(reader.pages[idx])
                else:
                    idx = int(part.strip()) - 1
                    if 0 <= idx < total_pages: writer.add_page(reader.pages[idx])
        else:
            writer.add_page(reader.pages[0])
        with open(out_path, 'wb') as f: writer.write(f)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Splitting error: {str(e)}", 500
    finally:
        if os.path.exists(pdf_path): os.remove(pdf_path)

# 7. TOOL: BLUR TO CLEAR AI
@app.route('/api/blur-to-clear', methods=['POST'])
def api_blur_to_clear():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    img_path = os.path.join(UPLOAD_FOLDER, file.filename)
    out_path = os.path.join(OUTPUT_FOLDER, 'cleared_' + file.filename)
    file.save(img_path)
    try:
        img = cv2.imread(img_path)
        gaussian_blur = cv2.GaussianBlur(img, (0, 0), 3)
        sharpened = cv2.addWeighted(img, 1.5, gaussian_blur, -0.5, 0)
        cv2.imwrite(out_path, sharpened)
        return send_file(out_path, as_attachment=True)
    except Exception as e: return f"Image processing error: {str(e)}", 500
    finally:
        if os.path.exists(img_path): os.remove(img_path)

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
