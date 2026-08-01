"""
Invoice Generator - Web Application (Smart Edition)
--------------------------------------------------
A user-friendly, browser-based Web UI for generating PDF invoices.
Includes Direct Data Entry Grid, Smart Auto-Mapping, Forex Live Rates, and Mobile Optimization.

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
    """Returns sample invoice data with smart preprocessing."""
    sample_path = os.path.join(BASE_DIR, "sample_invoice_input.xlsx")
    if not os.path.exists(sample_path):
        return jsonify({"error": "Sample file not found"}), 404

    try:
        raw_df = pd.read_excel(sample_path, dtype=str)
        cleaned_df, suggestions, errors = generate_invoices.smart_preprocess_dataframe(raw_df)
        data = cleaned_df.fillna("").to_dict(orient="records")
        columns = list(cleaned_df.columns)
        return jsonify({
            "columns": columns,
            "data": data,
            "filename": "sample_invoice_input.xlsx",
            "suggestions": suggestions
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Uploads and parses an Excel spreadsheet with smart preprocessing."""
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    fname_lower = file.filename.lower()
    if not (fname_lower.endswith('.xlsx') or fname_lower.endswith('.xls') or fname_lower.endswith('.xlsm')):
        return jsonify({"error": "Please upload a valid Excel file (.xlsx, .xls, or .xlsm)"}), 400

    safe_fname = "".join(c for c in file.filename if c.isalnum() or c in (" ", "_", ".", "-")).strip()
    if not safe_fname:
        safe_fname = "uploaded_invoice.xlsx"

    save_path = os.path.join(UPLOAD_DIR, safe_fname)
    file.save(save_path)

    try:
        raw_df = pd.read_excel(save_path, dtype=str)
        cleaned_df, suggestions, errors = generate_invoices.smart_preprocess_dataframe(raw_df)
        
        if errors:
            return jsonify({"error": "<br>".join(errors), "suggestions": suggestions}), 400

        data = cleaned_df.fillna("").to_dict(orient="records")
        columns = list(cleaned_df.columns)
        return jsonify({
            "columns": columns,
            "data": data,
            "filename": safe_fname,
            "filepath": save_path,
            "suggestions": suggestions
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

    # Save active UI grid state to openpyxl sheet for rendering
    if data:
        temp_excel = os.path.join(UPLOAD_DIR, "_active_grid.xlsx")
        df = pd.DataFrame(data)
        df.to_excel(temp_excel, index=False, engine="openpyxl")
        filepath = temp_excel

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


@app.route('/api/export-excel', methods=['POST'])
def export_excel():
    """Converts active grid data into a downloadable Excel spreadsheet with standardized columns & auto-spacing."""
    req_json = request.get_json(silent=True) or {}
    data = req_json.get('data', [])
    if not data:
        return jsonify({"error": "No data to export"}), 400

    try:
        df = pd.DataFrame(data)
        
        # Standard upload column order
        standard_cols = [
            "Invoice No", "Invoice Date", "Buyer Name", "Buyer Address Line1",
            "Buyer Address Line2", "Country", "Particulars", "Period Description",
            "HSN/SAC", "Amount", "Currency", "LUT Bond No", "LUT From", "LUT To",
            "Exchange Rate", "PO Number", "PO Date", "Project/Order Number",
            "Payment Terms", "Project Cost"
        ]
        
        ordered = [c for c in standard_cols if c in df.columns] + [c for c in df.columns if c not in standard_cols]
        df = df[ordered]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Invoice_Data')
            ws = writer.sheets['Invoice_Data']
            
            # Auto-space column widths so all text/data is visible without cutoffs
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    val = str(cell.value or '')
                    if len(val) > max_len:
                        max_len = len(val)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Standard_Invoice_Input.xlsx'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


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
    print("  EXCEL TO PDF INVOICE GENERATOR (SMART & MOBILE READY)  ")
    print("  Automated Excel-to-PDF Professional Invoice Creator  ")
    print("  Opening browser at http://127.0.0.1:5000 ...  ")
    print("=======================================================\n")
    
    threading.Timer(1.2, open_browser).start()
    app.run(host='0.0.0.0', port=5000, debug=False)
