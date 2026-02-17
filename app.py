import pikepdf
import pandas as pd
from xhtml2pdf import pisa
import os
from flask import Flask, render_template, request, send_file, after_this_request, send_from_directory
from werkzeug.utils import secure_filename
from pdf2docx import Converter
from PyPDF2 import PdfMerger, PdfReader, PdfWriter

app = Flask(__name__)

# --- CONFIGURATION ---
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

# --- ROUTE: HOME ---
@app.route('/')
def home():
    return render_template('home.html')

# ==========================================
# TOOL 1: PDF TO WORD (Existing)
# ==========================================
@app.route('/pdf-to-word')
def pdf_to_word_page():
    return render_template('pdf_to_word.html')

@app.route('/api/convert-pdf', methods=['POST'])
def convert_pdf_api():
    file = request.files['file']
    filename = secure_filename(file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)
    
    docx_filename = filename.rsplit('.', 1)[0] + '.docx'
    docx_path = os.path.join(app.config['CONVERTED_FOLDER'], docx_filename)
    
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None, ocr=0)
    cv.close()
    
    return send_file(docx_path, as_attachment=True)

# ==========================================
# TOOL 2: MERGE PDF (New!)
# ==========================================
@app.route('/merge-pdf')
def merge_pdf_page():
    return render_template('merge_pdf.html')

@app.route('/api/merge-pdf', methods=['POST'])
def merge_pdf_api():
    files = request.files.getlist('files')
    if len(files) < 2:
        return "Upload at least 2 files", 400

    merger = PdfMerger()
    
    # Process files
    for file in files:
        filename = secure_filename(file.filename)
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(path)
        merger.append(path)

    output_name = "Merged_Document.pdf"
    output_path = os.path.join(app.config['CONVERTED_FOLDER'], output_name)
    
    merger.write(output_path)
    merger.close()

    return send_file(output_path, as_attachment=True)

# ==========================================
# TOOL 3: SPLIT PDF (New!)
# ==========================================
@app.route('/split-pdf')
def split_pdf_page():
    return render_template('split_pdf.html')

@app.route('/api/split-pdf', methods=['POST'])
def split_pdf_api():
    file = request.files['file']
    # Default to extracting page 1 if not specified (Simple version)
    # In a full version, you'd accept a form input for page numbers
    
    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_path)

    reader = PdfReader(input_path)
    writer = PdfWriter()

    # LOGIC: Extract all pages into separate files, but for now let's just zip them 
    # OR for simplicity, let's just extract the First Page as a demo
    writer.add_page(reader.pages[0])
    
    output_name = f"Split_Page_1_{filename}"
    output_path = os.path.join(app.config['CONVERTED_FOLDER'], output_name)
    
    with open(output_path, "wb") as f:
        writer.write(f)

    return send_file(output_path, as_attachment=True)

# --- Don't forget to import this at the top of app.py! ---
import img2pdf 

# ==========================================
# TOOL 4: JPG TO PDF (New!)
# ==========================================
@app.route('/jpg-to-pdf')
def jpg_to_pdf_page():
    return render_template('jpg_to_pdf.html')

@app.route('/api/jpg-to-pdf', methods=['POST'])
def jpg_to_pdf_api():
    files = request.files.getlist('files')
    
    if len(files) == 0:
        return "No files selected", 400

    # 1. Save all images temporarily
    img_paths = []
    for file in files:
        filename = secure_filename(file.filename)
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(path)
        img_paths.append(path)

    # 2. Convert to PDF
    output_pdf_name = "Converted_Images.pdf"
    output_path = os.path.join(app.config['CONVERTED_FOLDER'], output_pdf_name)
    
    # img2pdf writes the PDF bytes directly to the file
    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(img_paths))

    # 3. Send file to user
    return send_file(output_path, as_attachment=True)

# ==========================================
# TOOL 5: COMPRESS PDF
# ==========================================
@app.route('/compress-pdf')
def compress_pdf_page():
    return render_template('compress_pdf.html')

@app.route('/api/compress-pdf', methods=['POST'])
def compress_pdf_api():
    file = request.files['file']
    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_path)
    
    # OUTPUT PATH
    output_name = f"Compressed_{filename}"
    output_path = os.path.join(app.config['CONVERTED_FOLDER'], output_name)
    
    # 1. Open with PikePDF
    with pikepdf.open(input_path) as pdf:
        # 2. Save with compression enabled
        # This removes unused objects and compresses streams
        pdf.save(output_path, compress_streams=True)
    
    return send_file(output_path, as_attachment=True)

# ==========================================
# TOOL 6: EXCEL TO PDF
# ==========================================
@app.route('/excel-to-pdf')
def excel_to_pdf_page():
    return render_template('excel_to_pdf.html')

@app.route('/api/excel-to-pdf', methods=['POST'])
def excel_to_pdf_api():
    file = request.files['file']
    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_path)
    
    # 1. Read Excel using Pandas
    # This reads the first sheet
    df = pd.read_excel(input_path)
    
    # 2. Convert Data to HTML (Simple Table)
    html_string = df.to_html(classes='table table-striped')
    
    # Add some basic styling so the PDF doesn't look ugly
    styled_html = f"""
    <html>
    <head>
        <style>
            table {{ width: 100%; border-collapse: collapse; }}
            th, td {{ border: 1px solid black; padding: 8px; text-align: left; font-size: 10px; }}
            th {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>
        <h2>{filename}</h2>
        {html_string}
    </body>
    </html>
    """
    
    # 3. Convert HTML to PDF
    output_name = f"{filename.rsplit('.', 1)[0]}.pdf"
    output_path = os.path.join(app.config['CONVERTED_FOLDER'], output_name)
    
    with open(output_path, "wb") as pdf_file:
        pisa_status = pisa.CreatePDF(styled_html, dest=pdf_file)
        
    if pisa_status.err:
        return "Error converting file", 500

    return send_file(output_path, as_attachment=True)

# ... (All your previous tool codes are up here) ...

# ==========================================
# SEO ROUTES (Robots & Sitemap)
# ==========================================
@app.route('/robots.txt')
def robots():
    return send_from_directory('static', 'robots.txt')

@app.route('/sitemap.xml')
def sitemap():
    return send_from_directory('static', 'sitemap.xml')

# --- PASTE THE CODE ABOVE THIS LINE ---
if __name__ == '__main__':
    app.run(debug=True, port=5000)

if __name__ == '__main__':
    app.run(debug=True, port=5000)