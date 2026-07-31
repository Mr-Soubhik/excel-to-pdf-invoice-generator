"""
Invoice Generator Engine
------------------------
Reads an Excel file (one row = one line item, grouped by Invoice No),
validates the data, fills the Word template, and exports one PDF per invoice.

Usage:
    python generate_invoices.py <input_excel_path> <output_folder>
"""

import sys
import os
import subprocess
import shutil
from datetime import datetime

import pandas as pd
from docxtpl import DocxTemplate
from num2words import num2words

# ---------- CONFIG ----------
REQUIRED_COLUMNS = [
    "Invoice No", "Invoice Date", "Buyer Name", "Buyer Address Line1",
    "Buyer Address Line2", "Country", "Particulars", "Period Description",
    "HSN/SAC", "Amount", "Currency",
    "LUT Bond No", "LUT From", "LUT To",
]
# Optional columns: fine to be missing from the sheet entirely, or blank per-row.
# If missing from the sheet, treated as blank for every row.
OPTIONAL_COLUMNS = [
    "Exchange Rate",       # only needed when Currency != INR
    "PO Number",
    "PO Date",
    "Project/Order Number",
    "Payment Terms",
    "Project Cost",
]

CURRENCY_NAMES = {
    "USD": "US Dollars", "EUR": "Euros", "GBP": "Pounds Sterling",
    "INR": "Rupees", "AUD": "Australian Dollars", "CAD": "Canadian Dollars",
    "SGD": "Singapore Dollars", "AED": "UAE Dirhams", "JPY": "Japanese Yen",
}
CURRENCY_SYMBOLS = {
    "USD": "$", "EUR": "\u20ac", "GBP": "\u00a3", "INR": "\u20b9",
    "AUD": "A$", "CAD": "C$", "SGD": "S$", "AED": "AED ", "JPY": "\u00a5",
}
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "invoice_template.docx")


# ---------- VALIDATION ----------
def validate_dataframe(df):
    errors = []

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        errors.append(f"Missing required column(s): {', '.join(missing_cols)}")
        return errors  # can't validate further without the columns

    # Ensure optional columns exist (as all-blank) so downstream code doesn't KeyError
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    for idx, row in df.iterrows():
        row_num = idx + 2  # +2 because Excel is 1-indexed and has a header row
        if pd.isna(row["Invoice No"]) or str(row["Invoice No"]).strip() == "":
            errors.append(f"Row {row_num}: Invoice No is empty")
        if pd.isna(row["Buyer Name"]) or str(row["Buyer Name"]).strip() == "":
            errors.append(f"Row {row_num}: Buyer Name is empty")
        if pd.isna(row["Particulars"]) or str(row["Particulars"]).strip() == "":
            errors.append(f"Row {row_num}: Particulars is empty")

        amt = row["Amount"]
        if pd.isna(amt):
            errors.append(f"Row {row_num}: Amount is empty")
        else:
            try:
                float(amt)
            except (ValueError, TypeError):
                errors.append(f"Row {row_num}: Amount is not a valid number ('{amt}')")

        currency = str(row["Currency"]).strip().upper() if not pd.isna(row["Currency"]) else ""
        if not currency:
            errors.append(f"Row {row_num}: Currency is empty")
        elif currency != "INR":
            rate = row.get("Exchange Rate", None)
            if rate is None or pd.isna(rate):
                errors.append(f"Row {row_num}: Exchange Rate is required when Currency is {currency} (not INR)")
            else:
                try:
                    float(rate)
                except (ValueError, TypeError):
                    errors.append(f"Row {row_num}: Exchange Rate is not a valid number ('{rate}')")

    return errors


# ---------- HELPERS ----------
def fmt_amount(value, currency="USD"):
    """Format a number as '$ 12,200' style using the right symbol for the currency."""
    symbol = CURRENCY_SYMBOLS.get(currency.upper(), currency.upper() + " ")
    if value == int(value):
        return f"{symbol}{value:,.0f}"
    return f"{symbol}{value:,.2f}"


def amount_to_words(value, currency="USD"):
    """12200, 'USD' -> 'Twelve Thousand Two Hundred US Dollars'"""
    words = num2words(int(round(value)), lang="en").replace(",", "")
    currency_name = CURRENCY_NAMES.get(currency.upper(), currency.upper())
    return f"{words.title()} {currency_name}"


def safe_invoice_filename(invoice_no, buyer_name):
    """Turn 'WTEXI/26-27/077' + 'ABC CORP' into a safe filename."""
    inv = invoice_no.replace("/", "-").strip()
    buyer = "".join(c for c in buyer_name if c.isalnum() or c in (" ", "_")).strip().replace(" ", "_")
    return f"Invoice_{inv}_{buyer}.pdf"


def convert_docx_to_pdf(docx_path, output_dir):
    """Uses LibreOffice headless to convert docx -> pdf."""
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
        check=True,
        capture_output=True,
    )


# ---------- MAIN GENERATION LOGIC ----------
def generate_invoices(input_excel_path, output_folder):
    log_lines = []
    log_lines.append(f"Invoice generation started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log_lines.append(f"Input file: {input_excel_path}")

    if not os.path.exists(input_excel_path):
        log_lines.append(f"ERROR: Input file not found: {input_excel_path}")
        return log_lines, []

    raw_df = pd.read_excel(input_excel_path, dtype=str)

    # Check if attributes are listed vertically (Column A contains field names like "Invoice No")
    first_col_vals = [str(x).strip() for x in raw_df.iloc[:, 0].dropna()]
    if "Invoice No" in first_col_vals or "Buyer Name" in first_col_vals:
        # Transpose vertical format: rows become columns, columns become rows
        header_col = raw_df.columns[0]
        raw_df = raw_df.dropna(subset=[header_col])
        raw_df[header_col] = raw_df[header_col].astype(str).str.strip()
        transposed = raw_df.set_index(header_col).T
        transposed.columns.name = None
        df = transposed.reset_index(drop=True)
    else:
        df = raw_df

    # Validate
    errors = validate_dataframe(df)
    if errors:
        log_lines.append("\nVALIDATION FAILED. Fix the following issues and re-run:")
        for e in errors:
            log_lines.append(f"  - {e}")
        return log_lines, []

    log_lines.append("Validation passed. No missing or invalid fields found.\n")

    os.makedirs(output_folder, exist_ok=True)
    temp_dir = os.path.join(output_folder, "_temp_docx")
    os.makedirs(temp_dir, exist_ok=True)

    generated_files = []

    # Group rows by Invoice No -> each group is one invoice
    grouped = df.groupby("Invoice No", sort=False)

    for invoice_no, group in grouped:
        first = group.iloc[0]
        currency = str(first["Currency"]).strip().upper()

        # Build line items
        items = []
        total_amount = 0.0
        for i, (_, row) in enumerate(group.iterrows(), start=1):
            amt = float(row["Amount"])
            total_amount += amt
            items.append({
                "sl_no": i,
                "particulars": str(row["Particulars"]),
                "period": str(row["Period Description"]) if not pd.isna(row["Period Description"]) else "",
                "hsn_sac": str(row["HSN/SAC"]),
                "amount_fmt": fmt_amount(amt, currency),
            })

        # INR conversion remark only applies when invoice currency isn't already INR
        remarks = ""
        if currency != "INR":
            rate_val = first.get("Exchange Rate", None)
            if rate_val is not None and not pd.isna(rate_val):
                exchange_rate = float(rate_val)
                inr_value = total_amount * exchange_rate
                remarks = f"{currency} {total_amount:,.0f} @ {exchange_rate} = INR {inr_value:,.0f}/-"

        def opt(col_name):
            """Return a stripped string for an optional column, or '' if blank/missing."""
            val = first.get(col_name, None)
            if val is None or pd.isna(val):
                return ""
            return str(val).strip()

        context = {
            "invoice_no": str(invoice_no),
            "invoice_date": str(first["Invoice Date"]),
            "buyer_name": str(first["Buyer Name"]),
            "buyer_address_line1": str(first["Buyer Address Line1"]),
            "buyer_address_line2": str(first["Buyer Address Line2"]),
            "country": str(first["Country"]),
            "lut_bond_no": str(first["LUT Bond No"]),
            "lut_from": str(first["LUT From"]),
            "lut_to": str(first["LUT To"]),
            "items": items,  # docxtpl repeats the line-item table row once per entry
            "total_amount_fmt": fmt_amount(total_amount, currency),
            "currency": currency,
            "amount_in_words": amount_to_words(total_amount, currency),
            "hsn_sac_summary": items[0]["hsn_sac"] if items else "",
            "igst_rate": "0%",
            "igst_amount": "",
            "total_tax_amount": "",
            "tax_amount_words": "NIL",
            "remarks": remarks,
            "po_number": opt("PO Number"),
            "po_date": opt("PO Date"),
            "project_order_number": opt("Project/Order Number"),
            "payment_terms": opt("Payment Terms"),
            "project_cost": opt("Project Cost"),
        }

        # Render template — `items` list drives the repeating line-item table row
        doc = DocxTemplate(TEMPLATE_PATH)
        doc.render(context)

        safe_name = safe_invoice_filename(str(invoice_no), str(first["Buyer Name"]))
        docx_out_path = os.path.join(temp_dir, safe_name.replace(".pdf", ".docx"))
        doc.save(docx_out_path)

        try:
            convert_docx_to_pdf(docx_out_path, output_folder)
            pdf_generated_path = os.path.join(
                output_folder, os.path.basename(docx_out_path).replace(".docx", ".pdf")
            )
            generated_files.append(pdf_generated_path)
            log_lines.append(f"Generated: {os.path.basename(pdf_generated_path)}  (Total: {fmt_amount(total_amount, currency)}, {len(items)} line item(s))")
        except subprocess.CalledProcessError as e:
            log_lines.append(f"ERROR converting {safe_name} to PDF: {e.stderr.decode(errors='ignore')}")

    shutil.rmtree(temp_dir, ignore_errors=True)

    log_lines.append(f"\nDone. {len(generated_files)} invoice(s) generated in: {output_folder}")
    return log_lines, generated_files


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generate_invoices.py <input_excel_path> <output_folder>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_dir = sys.argv[2]

    logs, files = generate_invoices(input_path, output_dir)
    print("\n".join(logs))
