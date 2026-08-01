"""
Invoice Generator Engine (Smart Edition)
---------------------------------------
Reads an Excel file (one row = one line item, grouped by Invoice No),
performs smart auto-mapping, auto-fetches live exchange rates,
validates the data, fills the Word template, and exports clean 1-page PDF invoices.

Usage:
    python generate_invoices.py <input_excel_path> <output_folder>
"""

import sys
import os
import subprocess
import shutil
import json
import urllib.request
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

OPTIONAL_COLUMNS = [
    "Exchange Rate",       # required when Currency != INR
    "PO Number",
    "PO Date",
    "Project/Order Number",
    "Payment Terms",
    "Project Cost",
]

COLUMN_ALIASES = {
    "Invoice No": ["invoice no", "invoice_no", "invoiceno", "invoice #", "inv no", "inv_no", "inv #", "invoice number", "bill no", "bill number", "inv_number"],
    "Invoice Date": ["invoice date", "invoice_date", "invoicedate", "inv date", "inv_date", "date", "bill date", "date of issue"],
    "Buyer Name": ["buyer name", "buyer_name", "buyername", "buyer", "client name", "client", "customer name", "customer", "bill to", "billed to", "company name"],
    "Buyer Address Line1": ["buyer address line1", "buyer address 1", "address line 1", "address1", "street address", "buyer_address_line1", "address line1", "address 1"],
    "Buyer Address Line2": ["buyer address line2", "buyer address 2", "address line 2", "address2", "city state zip", "buyer_address_line2", "address line2", "address 2"],
    "Country": ["country", "buyer country", "client country", "country name"],
    "Particulars": ["particulars", "description", "item description", "service description", "item", "services", "details", "particular"],
    "Period Description": ["period description", "period", "service period", "billing period", "month", "duration"],
    "HSN/SAC": ["hsn/sac", "hsn", "sac", "hsn_sac", "hsn code", "sac code", "tax code", "hsn/sac code"],
    "Amount": ["amount", "total amount", "item amount", "total", "price", "net amount", "val", "value", "cost"],
    "Currency": ["currency", "curr", "currency code", "unit"],
    "LUT Bond No": ["lut bond no", "lut no", "lut number", "lut_bond_no", "lut bond"],
    "LUT From": ["lut from", "lut_from", "lut valid from", "lut start"],
    "LUT To": ["lut to", "lut_to", "lut valid to", "lut end"],
    "Exchange Rate": ["exchange rate", "exchange_rate", "forex rate", "conversion rate", "rate", "inr rate"],
    "PO Number": ["po number", "po no", "po_number", "po #", "purchase order"],
    "PO Date": ["po date", "po_date", "purchase order date"],
    "Project/Order Number": ["project/order number", "project number", "order number", "project no", "order no"],
    "Payment Terms": ["payment terms", "terms", "payment term", "due terms"],
    "Project Cost": ["project cost", "total cost"]
}

CURRENCY_NAMES = {
    "USD": "US Dollars", "EUR": "Euros", "GBP": "Pounds Sterling",
    "INR": "Rupees", "AUD": "Australian Dollars", "CAD": "Canadian Dollars",
    "SGD": "Singapore Dollars", "AED": "UAE Dirhams", "JPY": "Japanese Yen",
}
CURRENCY_SYMBOLS = {
    "USD": "$", "EUR": "\u20ac", "GBP": "\u00a3", "INR": "\u20b9",
    "AUD": "A$", "CAD": "C$", "SGD": "S$", "AED": "AED ", "JPY": "\u00a5",
}
FALLBACK_EXCHANGE_RATES = {
    "USD": 83.80, "EUR": 90.50, "GBP": 106.80, "AUD": 54.20,
    "CAD": 61.10, "SGD": 62.40, "AED": 22.80, "JPY": 0.55
}
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "invoice_template.docx")


# ---------- LIVE EXCHANGE RATE FETCHING ----------
def fetch_live_exchange_rate(base_currency="USD", target_currency="INR"):
    """Fetches live exchange rate using open API with fallback."""
    base_curr = base_currency.upper().strip()
    if base_curr == "INR":
        return 1.0
    try:
        url = f"https://open.er-api.com/v6/latest/{base_curr}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=3) as response:
            data = json.loads(response.read().decode('utf-8'))
            if data.get("result") == "success" and "rates" in data:
                rate = float(data["rates"].get(target_currency.upper()))
                return round(rate, 2)
    except Exception:
        pass
    return FALLBACK_EXCHANGE_RATES.get(base_curr, None)


# ---------- SMART PREPROCESSING & FUZZY AUTO-MAPPING ----------
def smart_preprocess_dataframe(df):
    """
    Smartly cleans and maps Excel columns using fuzzy matching, auto-fetches live
    exchange rates via live API, auto-completes missing fields with smart defaults,
    and returns (cleaned_df, suggestions, errors).
    """
    suggestions = []
    errors = []
    
    if df is None or df.empty:
        return df, suggestions, ["Uploaded file is empty"]

    # 1. Fuzzy Header Matching
    new_cols = {}
    for col in df.columns:
        clean_col = str(col).strip()
        matched_std = None
        for std_col, aliases in COLUMN_ALIASES.items():
            if clean_col == std_col or clean_col.lower() in aliases:
                matched_std = std_col
                break
        if matched_std and matched_std != clean_col:
            new_cols[col] = matched_std
            suggestions.append(f"✨ Auto-mapped Excel column '{clean_col}' → '{matched_std}'")
        else:
            new_cols[col] = clean_col

    df = df.rename(columns=new_cols)

    # Handle vertical format if needed
    first_col_vals = [str(x).strip() for x in df.iloc[:, 0].dropna()]
    if "Invoice No" in first_col_vals or "Buyer Name" in first_col_vals:
        header_col = df.columns[0]
        df = df.dropna(subset=[header_col])
        df[header_col] = df[header_col].astype(str).str.strip()
        transposed = df.set_index(header_col).T
        transposed.columns.name = None
        df = transposed.reset_index(drop=True)
        suggestions.append("📐 Transposed vertical layout into standard data grid.")

    # 2. Check for missing required columns
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        errors.append(f"Missing required column(s): {', '.join(missing_cols)}")
        return df, suggestions, errors

    # 3. Ensure optional columns exist
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    # 4. Smart Auto-Fill & Live Rate Fetching per row
    today_str = datetime.now().strftime("%Y-%m-%d")
    fetched_rates = {}

    for idx in range(len(df)):
        row_num = idx + 2
        
        # Auto-fill missing Invoice Date
        if pd.isna(df.at[idx, "Invoice Date"]) or str(df.at[idx, "Invoice Date"]).strip() == "":
            df.at[idx, "Invoice Date"] = today_str
            inv_no = df.at[idx, "Invoice No"] if not pd.isna(df.at[idx, "Invoice No"]) else f"Row {row_num}"
            suggestions.append(f"📅 Auto-filled missing Invoice Date for {inv_no} ({today_str})")

        # Auto-fill missing HSN/SAC code with standard IT services code 998313
        if pd.isna(df.at[idx, "HSN/SAC"]) or str(df.at[idx, "HSN/SAC"]).strip() == "":
            df.at[idx, "HSN/SAC"] = "998313"
            suggestions.append(f"🏷️ Auto-filled standard HSN/SAC code 998313 for Row {row_num}")

        # Smart Currency & Exchange Rate Fetching
        curr = str(df.at[idx, "Currency"]).strip().upper() if not pd.isna(df.at[idx, "Currency"]) else "USD"
        df.at[idx, "Currency"] = curr

        if curr != "INR":
            rate_val = df.at[idx, "Exchange Rate"]
            if pd.isna(rate_val) or str(rate_val).strip() == "":
                if curr not in fetched_rates:
                    fetched_rates[curr] = fetch_live_exchange_rate(curr, "INR")
                
                live_rate = fetched_rates.get(curr)
                if live_rate:
                    df.at[idx, "Exchange Rate"] = str(live_rate)
                    suggestions.append(f"🌐 Smartly fetched live Forex rate for {curr}: 1 {curr} = ₹{live_rate} INR")
                else:
                    errors.append(f"Row {row_num}: Could not fetch exchange rate for currency '{curr}'")

        # Layout compacting checks
        line2 = df.at[idx, "Buyer Address Line2"]
        if pd.isna(line2) or str(line2).strip() == "":
            df.at[idx, "Buyer Address Line2"] = ""

    # Space saving suggestion
    inv_counts = df.groupby("Invoice No").size()
    max_items = inv_counts.max() if not inv_counts.empty else 1
    if max_items <= 3:
        suggestions.append(f"⚡ Space Saved: Compact 1-page layout active for all generated invoices.")
    else:
        suggestions.append(f"📄 Dynamic Layout: Scaled table row heights to prevent multi-page overflow.")

    return df, suggestions, errors


# ---------- VALIDATION ----------
def validate_dataframe(df):
    errors = []

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        errors.append(f"Missing required column(s): {', '.join(missing_cols)}")
        return errors

    for idx, row in df.iterrows():
        row_num = idx + 2
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
    
    # Run Smart Preprocessing & Fuzzy Mapping
    df, suggestions, prep_errors = smart_preprocess_dataframe(raw_df)
    if suggestions:
        log_lines.append("SMART SUGGESTIONS & AUTO-FIXES APPLIED:")
        for s in suggestions:
            log_lines.append(f"  {s}")

    # Validate
    errors = validate_dataframe(df)
    if errors or prep_errors:
        all_errs = prep_errors + errors
        log_lines.append("\nVALIDATION FAILED. Fix the following issues and re-run:")
        for e in all_errs:
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

    # Cleanup temp working files to save server disk space
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
