# 📑 EXCEL TO PDF INVOICE GENERATOR (SMART & MOBILE EDITION)

> **Automated Excel-to-PDF Professional Invoice Creator By Soubhik**

A zero-commandline, non-tech friendly Web Application to generate professional PDF invoices from Excel spreadsheets with Smart Auto-Mapping, Live Forex Rates, 1-Page Layout Optimizer, and Mobile responsiveness.

---

## 🚀 How to Run (Local)

Run:
```bash
python app.py
```
Your web browser will **automatically open** to:
👉 **[http://127.0.0.1:5000](http://127.0.0.1:5000)**

---

## 🌐 How to Run with FREE Cloudflare Public URL (For Client & Mobile)

To give your client or access from your smartphone via a secure, encrypted **https://** web link:

```bash
./run_with_cloudflare.sh
```

This will automatically start the server and print a free Cloudflare HTTPS link like:
👉 **`https://your-tunnel-name.trycloudflare.com`**

---

## ✨ Smart & Mobile Features

1. **🤖 Smart Fuzzy Column Auto-Mapping**: Automatically recognizes column variations (`Invoice #`, `Bill To`, `Client`, `Date`, `Total`) and maps them to standard invoice fields.
2. **🌐 Live Forex Exchange Rates**: Auto-fetches real-time exchange rates (USD, EUR, GBP, AUD, CAD, SGD, AED, JPY to INR) via live API for foreign currency invoices.
3. **💡 Smart AI Suggestions Banner**: Displays actionable pre-flight feedback, auto-fixes, and warnings before generating PDFs.
4. **📱 Fully Mobile Responsive**: Re-designed for smartphones and tablets with touch-friendly controls and responsive data grid.
5. **📦 1-Page Layout & Space Saver**: Dynamically scales spacing so invoices fit on 1 clean page, and automatically cleans up temporary files to save server disk space.
6. **1-Click Demo & Downloads**: Single-click demo loading, in-browser spreadsheet editor, and 1-click batch ZIP download.

---

## 📁 Project Structure

```
invoice-generator/
│
├── app.py                     ← MAIN ENTRY POINT (Run this!)
├── generate_invoices.py       ← Core Smart PDF engine & Forex API
├── run_with_cloudflare.sh     ← Cloudflare 1-click launcher
├── cloudflared                ← Cloudflare tunnel executable
├── templates/
│   └── index.html             ← Smart & Mobile Web UI Interface
├── invoice_template.docx      ← Word template read at runtime
├── sample_invoice_input.xlsx  ← Built-in test spreadsheet
├── output/                    ← Generated PDF invoices land here
├── requirements.txt           ← Dependencies
└── README.md
```
