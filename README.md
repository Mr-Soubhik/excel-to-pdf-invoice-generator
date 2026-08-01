# 📑 EXCEL TO PDF INVOICE GENERATOR (DUAL MODE: UPLOAD & DIRECT ENTRY)

> **Automated Excel-to-PDF Professional Invoice Creator By Soubhik**

A zero-commandline, non-tech friendly Web Application to generate professional PDF invoices via **Excel Upload** OR **Direct In-Browser Data Entry** with Smart Auto-Mapping, Live Forex Rates, 1-Page Layout Optimizer, and Mobile responsiveness.

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

## ✨ Features

1. **📁 Excel Upload Mode**: Drag & drop any `.xlsx` invoice file into the browser.
2. **✨ ➕ Direct Data Entry Mode**: Click **"Direct Entry (New Invoice)"** to build an invoice grid from scratch in your browser without needing Excel installed!
3. **➕ Add Row & 📥 Export Excel**: Add new rows on the fly, edit fields in real-time, and export your manual data back into a clean `.xlsx` file.
4. **💾 Browser Draft Memory (`localStorage`)**: Saves your typed grid state automatically so your work is never lost on refresh.
5. **🤖 Smart Fuzzy Column Auto-Mapping**: Automatically recognizes column variations (`Invoice #`, `Bill To`, `Client`, `Date`, `Total`).
6. **🌐 Live Forex Exchange Rates**: Auto-fetches real-time exchange rates (USD, EUR, GBP, AUD, CAD, SGD, AED, JPY to INR) via live API.
7. **💡 Smart AI Suggestions Banner**: Displays pre-flight feedback and warnings before generating PDFs.
8. **📱 Fully Mobile Responsive**: Designed for smartphones and tablets with touch-friendly controls.
9. **📦 1-Page Layout & Space Saver**: Auto-fits multi-item invoices to 1 clean page.

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
