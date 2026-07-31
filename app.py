"""
Invoice Generator - Web Application
------------------------------------
A user-friendly, browser-based Web UI for generating PDF invoices.
No command-line needed for non-technical users.

Run with:
    python app.py
"""

import os
import io
import zipfile
import threading
import webbrowser
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
import pandas as pd
import generate_invoices

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/sample-data', methods=['GET'])
def get_sample_data():
    """Returns sample invoice data for 1-click non-tech testing."""
    sample_path = os.path.join(BASE_DIR, "sample_invoice_input.xlsx")
    if not os.path.exists(sample_path):
        return jsonify({"error": "Sample file not found"}), 44

    try:
        raw_df = pd.read_excel(sample_path, dtype=str)
        first_col_vals = [str(x).strip() for x in raw_df.iloc[:, 0].dropna()]
        if "Invoice No" in first_col_vals or "Buyer Name" in first_col_vals:
            header_col = raw_df.columns[0]
            raw_df = raw_df.dropna(subset=[header_col])
            raw_df[header_col] = raw_df[header_col].astype(str).str.strip()
            df = raw_df.set_index(header_col).T.reset_index(drop=True)
        else:
            df = raw_df

        data = df.fillna("").to_dict(orient="records")
        columns = list(df.columns)
        return jsonify({"columns": columns, "data": data, "filename": "sample_invoice_input.xlsx"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Uploads and parses an Excel spreadsheet."""
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        return jsonify({"error": "Please upload a valid Excel file (.xlsx or .xls)"}), 400

    save_path = os.path.join(UPLOAD_DIR, file.filename)
    file.save(save_path)

    try:
        raw_df = pd.read_excel(save_path, dtype=str)
        first_col_vals = [str(x).strip() for x in raw_df.iloc[:, 0].dropna()]
        if "Invoice No" in first_col_vals or "Buyer Name" in first_col_vals:
            header_col = raw_df.columns[0]
            raw_df = raw_df.dropna(subset=[header_col])
            raw_df[header_col] = raw_df[header_col].astype(str).str.strip()
            df = raw_df.set_index(header_col).T.reset_index(drop=True)
        else:
            df = raw_df

        data = df.fillna("").to_dict(orient="records")
        columns = list(df.columns)
        return jsonify({
            "columns": columns,
            "data": data,
            "filename": file.filename,
            "filepath": save_path
        })
    except Exception as e:
        return jsonify({"error": f"Failed to parse Excel file: {str(e)}"}), 500


@app.route('/api/generate', methods=['POST'])
def generate():
    """Generates PDF invoices from current data."""
    req_json = request.get_json(silent=True) or {}
    data = req_json.get('data', [])
    filepath = req_json.get('filepath', '')

    if not data and not filepath:
        filepath = os.path.join(BASE_DIR, "sample_invoice_input.xlsx")

    # If data grid was modified on screen, save to temporary Excel sheet
    if data:
        temp_excel = os.path.join(UPLOAD_DIR, "_active_grid.xlsx")
        df = pd.DataFrame(data)
        df.to_excel(temp_excel, index=False, engine="openpyxl")
        filepath = temp_excel

    template_path = os.path.join(BASE_DIR, "invoice_template.docx")

    try:
        logs, generated_files = generate_invoices.generate_invoices(filepath, OUTPUT_DIR)
        
        pdf_list = []
        for f in generated_files:
            bname = os.path.basename(f)
            pdf_list.append({
                "name": bname,
                "size_kb": round(os.path.getsize(f) / 1024, 1),
                "url": f"/api/download/{bname}"
            })

        return jsonify({
            "success": True,
            "logs": logs,
            "files": pdf_list,
            "count": len(pdf_list)
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/download/<filename>')
def download_pdf(filename):
    """Serves an individual PDF invoice file."""
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.route('/api/download-all')
def download_all_zip():
    """Bundles all generated PDFs into a single downloadable ZIP file."""
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in os.listdir(OUTPUT_DIR):
            if fname.endswith('.pdf'):
                fpath = os.path.join(OUTPUT_DIR, fname)
                zf.write(fpath, fname)
    memory_file.seek(0)
    return send_file(
        memory_file,
        mimetype='application/zip',
        as_attachment=True,
        download_name='Invoices_Batch.zip'
    )


def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000")


if __name__ == '__main__':
    print("\n=======================================================")
    print("  EXCEL TO PDF INVOICE GENERATOR  ")
    print("  Automated Excel-to-PDF Professional Invoice Creator By Soubhik  ")
    print("  Opening browser at http://127.0.0.1:5000 ...  ")
    print("=======================================================\n")
    
    # Auto-open browser 1.2s after starting
    threading.Timer(1.2, open_browser).start()
    app.run(host='127.0.0.1', port=5000, debug=False)
