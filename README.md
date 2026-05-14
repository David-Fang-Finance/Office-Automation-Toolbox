# Office Automation Toolbox

A feature-rich desktop application built with **PyQt5** that streamlines enterprise workflows by integrating SAP automation, intelligent document processing, and productivity tools into a single, unified interface.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![PyQt5](https://img.shields.io/badge/PyQt5-Desktop-green.svg)
![Windows](https://img.shields.io/badge/Windows-10/11-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-orange.svg)

---

## ✨ Highlights

- 🚀 **SAP Automation** — Automate F-03 clearing, FBL1N vendor line-item exports, and Open AP queries directly from your desktop via SAP GUI Scripting
- 🧠 **AI-Powered Invoice Extraction** — Leverage Google Gemini or Baidu OCR to extract structured data from invoices and customs declarations in batch
- 📊 **Excel Batch Processing** — Merge multiple workbooks with precision, preserving data integrity across large datasets
- 📧 **Email Automation** — Generate Outlook drafts programmatically with customizable templates and attachments
- 🎨 **Modern UI** — Clean, professional interface with real-time progress tracking and interactive data previews

---

## 📸 Feature Overview

| Module | Description |
|--------|-------------|
| ⏰ **Reminder Assistant** | Lightweight date-based task reminder with color-coded priorities |
| 🎯 **SAP Toolkit** | SAP GUI login, Open AP query, F-03 auto-clearing, and vendor detail export |
| 📄 **PDF Processor** | Split PDF files by page ranges or extract individual pages |
| 📊 **Excel Merger** | Batch-merge multiple Excel files with sheet-level or cross-sheet consolidation |
| ⚡ **Quick Access** | One-click shortcuts to frequently used files and applications |
| 🧾 **Invoice Extractor** | AI-driven extraction of invoice fields (number, date, amount, tax, etc.) from PDFs |
| 🔍 **OCR Invoice Extractor** | Baidu OCR-powered invoice recognition with high-precision text detection |
| ✨ **Gemini Invoice Extractor** | Google Gemini-powered multimodal invoice parsing with configurable prompts |
| 🗃️ **Access DB Query** | Connect to Microsoft Access databases, browse tables, and execute SQL queries |
| 📧 **Email Automation** | Create Outlook email drafts with HTML body, attachments, and recipient management |

---

## 🛠️ Prerequisites

| Dependency | Version | Purpose |
|------------|---------|---------|
| Python | ≥ 3.8 | Runtime |
| PyQt5 | ≥ 5.15 | GUI framework |
| pywin32 | ≥ 305 | Windows COM / SAP GUI scripting |
| openpyxl | ≥ 3.1 | Excel file I/O |
| pandas | ≥ 2.0 | Data manipulation |
| pdfplumber | ≥ 0.10 | PDF text extraction |
| PyPDF2 | ≥ 3.0 | PDF splitting |
| Pillow | ≥ 9.0 | Image processing |
| pyodbc | ≥ 5.0 | Access database connectivity |
| google-genai | ≥ 1.0 | Google Gemini API client |

> **Note:** SAP-related features require SAP GUI for Windows installed locally with scripting enabled.

---

## 🚀 Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/David-Fang-Finance/office-automation-toolbox.git
cd office-automation-toolbox
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure

Copy the example configuration and fill in your credentials:

```bash
cp config.example.json config.json
```

Edit `config.json` with your actual settings. Refer to [`input argument.txt`](input argument.txt) for a full list of configurable parameters.

> ⚠️ **`config.json` is git-ignored** — it will never be committed or pushed. Only `config.example.json` (a blank template) is tracked.

### 4. Run

```bash
python main.py
```

---

## 📁 Project Structure

```
office-automation-toolbox/
│
├── main.py                           # Application entry point
├── config.example.json               # Configuration template (tracked)
├── input argument.txt                # Full parameter reference
├── .gitignore                        # Git ignore rules
├── README.md                         # This file
│
├── features/                         # Core feature modules
│   ├── sap_integrated.py             #   SAP login, Open AP, F-03 clearing
│   ├── sap_fbl1n_export.py           #   SAP FBL1N vendor export
│   ├── pdf_split.py                  #   PDF split by page
│   ├── excel_merger.py               #   Multi-file Excel merge
│   ├── quick_open.py                 #   Quick file launcher
│   ├── invoice_extractor.py          #   PDF invoice & customs extraction
│   ├── ocr_invoice_extractor.py      #   Baidu OCR invoice recognition
│   ├── gemini_invoice_extractor.py   #   Gemini AI invoice parsing
│   ├── access_database_query.py      #   Access DB browser & query
│   ├── date_reminder.py              #   Date-based reminders
│   └── email_automation.py           #   Outlook draft generator
│
└── utils/
    └── config_manager.py             #   Centralized config read/write
```

---

## ⚙️ Configuration

All sensitive credentials and environment-specific settings are stored in `config.json` (not tracked by Git). Key configuration groups include:

| Section | Parameters |
|---------|-----------|
| **SAP** | `sap_path`, `client`, `user`, `password`, `connection_name`, `user_id` |
| **Baidu OCR** | `api_key`, `secret_key` |
| **Gemini** | `api_key`, `model_name` |
| **Email** | `smtp_server`, `smtp_port`, `username`, `password` |
| **Database** | `access_db_path` |
| **Paths** | `pdf_folder`, `excel_output`, `export_folder` |

For a complete parameter reference, see [`input argument.txt`](input argument.txt).

---

## 🔒 Security

This project follows security best practices for credential management:

- ✅ All credentials are stored in `config.json`, which is **git-ignored**
- ✅ Only a blank `config.example.json` template is committed to the repository
- ✅ No hardcoded secrets, API keys, user IDs, or internal paths in source code
- ✅ API keys are entered via the GUI at runtime and saved locally

---

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -m 'Add your feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
