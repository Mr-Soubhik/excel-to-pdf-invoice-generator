# 📑 EXCEL TO PDF INVOICE GENERATOR

> **Automated Excel-to-PDF Professional Invoice Creator By Soubhik**

A zero-commandline, non-tech friendly Web Application to generate professional PDF invoices from Excel spreadsheets.

---

## 🚀 How to Run (1-Step)

Simply run:

```bash
python app.py
```

Your web browser will **automatically open** to:
👉 **[http://127.0.0.1:5000](http://127.0.0.1:5000)**

---

## ✨ Features for Non-Tech Users

1. **Drag & Drop Upload**: Simply drop your `.xlsx` invoice file into the browser.
2. **1-Click Demo Data**: Click **"⚡ Load Sample Data"** to generate test invoices instantly without needing any files prepared.
3. **In-Browser Spreadsheet Editor**: View and edit invoice numbers, dates, buyer names, amounts, and descriptions directly inside your web browser before generating.
4. **Instant PDF Downloads**: Download individual PDFs or click **"📦 Download All as ZIP"** to get all invoices in one zip file.

---

## 📁 Project Structure

```
invoice-generator/
│
├── app.py                     ← MAIN ENTRY POINT (Run this!)
├── generate_invoices.py       ← Core PDF engine
├── templates/
│   └── index.html             ← Web UI Interface
├── invoice_template.docx      ← Word template read at runtime
├── sample_invoice_input.xlsx  ← Built-in test spreadsheet
├── output/                    ← Generated PDF invoices land here
├── requirements.txt           ← Dependencies
└── README.md
```
