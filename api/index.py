from flask import Flask, request, render_template, send_file
import os
import tempfile
import pandas as pd
import barcode
from barcode.writer import ImageWriter
from fpdf import FPDF
from werkzeug.utils import secure_filename
import shutil
import logging

logging.basicConfig(level=logging.DEBUG)
app = Flask(__name__, template_folder='templates')

def sanitize_upc(upc):
    try:
        if pd.isna(upc):
            return ''
        upc_str = str(int(float(upc)))
        return ''.join(filter(str.isdigit, upc_str))
    except (ValueError, TypeError):
        return ''

def pad_upc(upc):
    if not upc:
        return ''
    if len(upc) < 12:
        return upc.zfill(12)
    elif 12 < len(upc) < 13:
        return upc.zfill(13)
    return upc

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/generate', methods=['POST'])
def generate():
    temp_dir = None
    try:
        if 'file' not in request.files:
            return "Error: No file uploaded.", 400
        file = request.files['file']
        if file.filename == '':
            return "Error: No file selected.", 400
        if not file.filename.endswith('.xlsx'):
            return "Error: Only .xlsx files are supported.", 400

        temp_dir = tempfile.mkdtemp()
        file_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(file_path)

        df = pd.read_excel(file_path, sheet_name=0, skiprows=10, engine="openpyxl")
        if "Unnamed: 13" not in df.columns:
            return "Error: 'Unnamed: 13' column not found.", 400
        df = df[["Unnamed: 13"]].dropna()
        df.columns = ['UPC']

        upc_codes = df['UPC'].apply(sanitize_upc).apply(pad_upc)
        upc_codes = upc_codes[upc_codes != '000000000000']

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        images_per_row, images_per_col = 2, 3
        image_width, image_height, margin = 75, 50, 5

        for i, upc in enumerate(upc_codes):
            if not upc:
                continue
            barcode_class = barcode.get_barcode_class('upca' if len(upc) == 12 else 'ean13')
            upc_barcode = barcode_class(upc, writer=ImageWriter())
            barcode_path = os.path.join(temp_dir, f"{upc}.png")
            upc_barcode.save(barcode_path.rsplit('.png', 1)[0])

            if i % (images_per_row * images_per_col) == 0:
                pdf.add_page()
            row, col = divmod(i % (images_per_row * images_per_col), images_per_row)
            x = margin + col * (image_width + margin)
            y = margin + row * (image_height + margin)
            pdf.image(barcode_path, x, y, w=image_width, h=image_height)

        pdf_path = os.path.join(temp_dir, "barcodes.pdf")
        pdf.output(pdf_path)
        return send_file(pdf_path, as_attachment=True, download_name="barcodes.pdf")

    except Exception as e:
        logging.error(f"Server Error: {str(e)}", exc_info=True)
        return f"Error: {str(e)}", 500
    finally:
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    app.run(debug=True)
