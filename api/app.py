from flask import Flask, request, render_template, send_file
import os
import tempfile
import pandas as pd
import barcode
from barcode.writer import ImageWriter
from fpdf import FPDF
from werkzeug.utils import secure_filename
import shutil

app = Flask(__name__)

# Function to sanitize and process UPC codes
def sanitize_upc(upc):
    try:
        if pd.isna(upc):
            return ''
        upc_str = str(int(float(upc)))
        upc_str = ''.join(filter(str.isdigit, upc_str))
        return upc_str.zfill(12) if len(upc_str) < 12 else upc_str
    except ValueError:
        return ''

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/generate', methods=['POST'])
def generate_barcodes():
    if 'file' not in request.files:
        return "Error: No file uploaded.", 400

    file = request.files['file']
    if file.filename == '':
        return "Error: No file selected.", 400

    temp_dir = tempfile.mkdtemp()
    try:
        file_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(file_path)

        df = pd.read_excel(file_path)
        if "UPC" not in df.columns:
            return "Error: 'UPC' column not found.", 400

        upc_codes = df['UPC'].apply(sanitize_upc).dropna().unique()
        if len(upc_codes) == 0:
            return "Error: No valid UPC codes found.", 400

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        for upc in upc_codes:
            try:
                barcode_class = barcode.get_barcode_class('upca')
                upc_barcode = barcode_class(upc, writer=ImageWriter())
                barcode_path = os.path.join(temp_dir, f"{upc}.png")
                upc_barcode.save(barcode_path.split('.png')[0])
                pdf.add_page()
                pdf.image(barcode_path, x=10, y=10, w=100)
            except Exception as e:
                print(f"Error generating barcode for {upc}: {e}")

        pdf_path = os.path.join(temp_dir, "barcodes.pdf")
        pdf.output(pdf_path)

        return send_file(pdf_path, as_attachment=True, download_name="barcodes.pdf")
    except Exception as e:
        return f"Error processing file: {str(e)}", 500
    finally:
        shutil.rmtree(temp_dir)

# Export the Flask app for Vercel
app = app
