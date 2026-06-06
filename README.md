<div align="center">

<img src="assets/logo.png" alt="VNSGU Result Analyzer Logo" width="120" />

# 🎓 Student Result Analyzer
### SASCMA – STERS | VNSGU

**A powerful desktop application for analyzing, visualizing, and exporting student results from VNSGU university PDF mark-sheets.**

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776ab?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![PySide6](https://img.shields.io/badge/PySide6-6.10%2B-41cd52?style=for-the-badge&logo=qt&logoColor=white)](https://doc.qt.io/qtforpython/)
[![SQLite](https://img.shields.io/badge/SQLite-3%20Tables-003b57?style=for-the-badge&logo=sqlite&logoColor=white)](https://sqlite.org)
[![License](https://img.shields.io/badge/License-MIT-f97316?style=for-the-badge)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078d4?style=for-the-badge&logo=windows&logoColor=white)](https://windows.com)

---

</div>

## 📖 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [Screenshots](#-screenshots)
- [Project Structure](#-project-structure)
- [Database Architecture](#-database-architecture)
- [Tech Stack](#-tech-stack)
- [Installation](#-installation)
- [Usage](#-usage)
- [PDF Format Support](#-pdf-format-support)
- [Keyboard Shortcuts](#-keyboard-shortcuts)
- [Export Options](#-export-options)
- [Contributing](#-contributing)
- [License](#-license)

---

## 🌟 Overview

**Student Result Analyzer** is a feature-rich, offline desktop application built for **SASCMA (School of Applied Sciences, Computer, and Management Applications)** affiliated with **VNSGU (Veer Narmad South Gujarat University)**.

It automatically **parses official VNSGU result PDFs**, extracts per-student and per-subject marks, stores them in a normalized SQLite database, and provides:

- 📊 Interactive charts and KPI dashboards
- 🏫 College-wise comparative analysis
- 📄 Excel & PDF export with professional formatting
- 🔍 Search, filter, and rank students
- 📈 Subject-wise performance breakdown

> No internet required. 100% offline. All data stays on your machine.

---

## ✨ Features

### 📥 Smart PDF Import
- Drag-and-drop or browse to upload VNSGU result PDFs
- Auto-detects **course name**, **semester**, and **academic year** from PDF headers
- Parses **7 subjects** with separate **External (EXT)** and **Internal (INT)** marks
- Handles `AB` (absent), `ZR`, grace marks (`13+5`), and edge cases
- Detects **college name** from each page header automatically
- Accurate **PASS/FAIL** detection per student and per subject using official passing thresholds

### 📊 Dashboard & Analytics
- **5 KPI Cards**: Total Students, Pass Count, Fail Count, Pass %, Avg SGPA
- **Pass/Fail Pie Chart** with interactive color coding
- **Grade Distribution Bar Chart** (O / A+ / A / B+ / B / C / F)
- **Topper Highlight** with name and marks

### 🏫 College-wise Report
| Mode | What you see |
|------|-------------|
| **All Colleges (Overview)** | Comparison bar chart, overall pass/fail pie chart, summary table with Pass%, Avg SGPA, ATKT count per college |
| **Individual College** | KPIs + student-level table sorted by SGPA |

### 📋 Result Table
- Full sortable student table with Roll No, Name, Sub 1–7, Total, SGPA, Status, ATKT Count
- Color-coded PASS (green) / FAIL (red) rows
- Real-time search by name, roll number, or SP ID

### ❌ ATKT / Failed Students
- Dedicated page listing students with pending subjects
- Shows subject name, EXT marks, INT marks, and total for each failed subject

### 📈 Analysis & Charts
- Subject-wise pass percentage comparison
- Per-subject average EXT, INT, and Total
- Percentage range distribution (Below 50 / 50–59 / 60–69 / 70–79 / 80–89 / 90–100)

### 📤 Export
| Format | Content |
|--------|---------|
| **Excel (.xlsx)** | Styled table with subject columns, conditional formatting, charts, college sheet |
| **PDF Report** | Professional A4 report with university header image, summary stats, student table |

### ⌨️ Keyboard Shortcuts
| Shortcut | Action |
|----------|--------|
| `Ctrl+U` | Upload PDF |
| `Ctrl+F` | Search students |
| `Ctrl+E` | Export Excel |
| `Ctrl+P` | Export PDF |
| `Ctrl+R` | Refresh dashboard |

---

## 🖼️ Screenshots

> *(Add screenshots here after running the app)*

| Dashboard | College Overview | ATKT Report |
|-----------|-----------------|-------------|
| ![Dashboard](assets/screenshot_dashboard.png) | ![College](assets/screenshot_college.png) | ![ATKT](assets/screenshot_atkt.png) |

---

## 📁 Project Structure

```
Student Result Analyzer/
│
├── main.py                     # Main application — UI, routing, all pages
│
├── core/
│   ├── pdf_parser.py           # PDF parsing engine (pdfplumber-based)
│   └── database.py             # SQLite database layer (3-table schema)
│
├── assets/
│   ├── logo.png                # App logo (splash + sidebar)
│   ├── icon.ico                # Window icon
│   ├── university_pdf_header.png   # PDF export header (VNSGU)
│   └── sascma_pdf_header.png       # PDF export header (SASCMA)
│
├── Result/                     # Auto-generated export output folder
├── Analysis/                   # Analysis exports
│
└── requirements.txt            # Python dependencies
```

---

## 🗄️ Database Architecture

The application uses a **3-table normalized SQLite schema** (in-memory by default):

```sql
┌─────────────────────┐
│   import_sessions   │  ← One row per PDF upload
│─────────────────────│
│ id, pdf_filename    │
│ course, semester    │
│ academic_year       │
│ college_name        │
│ max_marks (700)     │
│ imported_at         │
└────────┬────────────┘
         │ 1:N
┌────────▼────────────┐
│      students       │  ← One row per student
│─────────────────────│
│ id, session_id      │
│ seat_no, sp_id      │
│ name, gender        │
│ total_marks, sgpa   │
│ overall_grade       │
│ overall_status      │
│ atkt_count, college │
└────────┬────────────┘
         │ 1:7
┌────────▼────────────┐
│   subject_marks     │  ← 7 rows per student
│─────────────────────│
│ id, student_id      │
│ subject_index (0–6) │
│ subject_name        │
│ ext_mark, int_mark  │
│ total_mark, grade   │
│ status (PASS/FAIL)  │
└─────────────────────┘
```

A **SQL VIEW** (`v_students_flat`) joins all 3 tables into a flat DataFrame for the UI — keeping the UI code backward-compatible while storing complete normalized data.

---

## 🛠️ Tech Stack

| Component | Technology | Purpose |
|-----------|-----------|---------|
| **UI Framework** | [PySide6](https://doc.qt.io/qtforpython/) | Desktop GUI (Qt6 bindings) |
| **PDF Parsing** | [pdfplumber](https://github.com/jsvine/pdfplumber) | Text + word extraction from PDFs |
| **Database** | SQLite3 (stdlib) | Normalized result storage |
| **Data Processing** | [pandas](https://pandas.pydata.org/) | DataFrame operations, stats |
| **Charts** | [matplotlib](https://matplotlib.org/) | Pie, bar, horizontal bar charts |
| **Excel Export** | [openpyxl](https://openpyxl.readthedocs.io/) | Styled .xlsx reports |
| **PDF Export** | [reportlab](https://www.reportlab.com/) | Professional PDF generation |
| **Language** | Python 3.10+ | Core application |

---

## ⚙️ Installation

### Prerequisites
- **Python 3.10 or higher** — [Download](https://python.org/downloads)
- **Windows 10/11** (tested; may work on Linux/macOS with minor adjustments)

### Step 1 — Clone the repository
```bash
git clone https://github.com/YOUR_USERNAME/student-result-analyzer.git
cd student-result-analyzer
```

### Step 2 — Create a virtual environment (recommended)
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

### Step 3 — Install dependencies
```bash
pip install -r requirements.txt
```

### Step 4 — Run the application
```bash
python main.py
```

> **Note:** On first run, the app may take a few seconds to load while importing Qt and matplotlib libraries.

---

## 🚀 Usage

### 1. Upload a PDF
- Click **"Upload PDF"** or press `Ctrl+U`
- Select a VNSGU result PDF from your system
- The parser will auto-detect subjects, course, semester, and academic year

### 2. View Dashboard
- All KPIs, charts, and student table are populated automatically after upload
- Navigate pages using the **left sidebar**

### 3. Analyze by College
- Go to **"College Wise Report"**
- Select **"All Colleges (Overview)"** for a comparative chart and summary table
- Select an individual college to drill down into student-level details

### 4. Find ATKT Students
- Go to **"ATKT / Failed"** page
- Lists all students with 1 or more pending subjects with subject-wise details

### 5. Export
- Use **"Export Excel"** or **"Export PDF"** buttons on any page
- Exports include all currently filtered/visible data

---

## 📄 PDF Format Support

This parser is specifically designed for **VNSGU result PDFs** with the following structure:

```
College Name: SASCMA
                     Subject1  Subject2  Subject3  Subject4  Subject5  Subject6  Subject7
EXT            18/28  18/28     18/28     18/28      9/14      9/14      9/14
INT             7/12   7/12      7/12      7/12       6/11      6/11      6/11

001  1234567890  M  STUDENT NAME  B.SC. IT  SEM-1
EXT  25  30  22  28  15  12  8  138
INT  10  12   9  11   0   0  0   42
TOTAL  35  42  31  39  15  12  8  182
GL GP CR  A+  O  A  A+  A  A  O
1-001-8.50      PASS YR
```

**Supported features:**
- ✅ Absent marks (`AB`, `ZR`, `-`)
- ✅ Grace marks (`13+5` format)
- ✅ Multi-page PDFs
- ✅ Multiple colleges in one PDF
- ✅ Auto semester detection (First, Second... Tenth)
- ✅ Auto academic year detection (NOV-2025 → 2025-2026)

---

## 📤 Export Options

### Excel Report
- **Sheet 1:** Full student data with styled headers and conditional formatting
- Subject columns with EXT/INT breakdown
- Color-coded PASS (green) / FAIL (red) rows
- Auto-sized columns

### PDF Report
- University/college header image
- Summary statistics block
- Formatted student table (A4 landscape)
- Professional typography with VNSGU branding

---

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. **Fork** the repository
2. **Create** a feature branch: `git checkout -b feature/your-feature`
3. **Commit** your changes: `git commit -m "Add: your feature description"`
4. **Push** to your branch: `git push origin feature/your-feature`
5. **Open a Pull Request**

### Development Guidelines
- Follow existing code style (PySide6 signals/slots pattern)
- Add docstrings to new functions
- Test with actual VNSGU PDFs before submitting
- Do not commit large binary files (PDFs, result sheets)

### Reporting Issues
Please include:
- Python version (`python --version`)
- OS version
- Error traceback (full)
- Sample PDF structure (if possible, anonymized)

---

## 📋 Roadmap

- [ ] 🌐 Multi-university PDF format support
- [ ] 💾 Persistent SQLite file storage (save between sessions)
- [ ] 📊 Year-over-year comparison charts
- [ ] 🔔 Notification system for ATKT deadlines
- [ ] 🖨️ Direct print support
- [ ] 🌍 Multi-language UI (Gujarati / Hindi)
- [ ] 📱 Mobile-friendly web export

---

## 📜 License

```
MIT License

Copyright (c) 2025 SASCMA – STERS | VNSGU

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software...
```

See [LICENSE](LICENSE) file for full text.

---

## 👨‍💻 Author

**Developed for SASCMA – STERS | VNSGU**

> School of Applied Sciences, Computer & Management Applications  
> Veer Narmad South Gujarat University, Surat

---

<div align="center">

**⭐ If this project helped you, please give it a star on GitHub! ⭐**

Made with ❤️ for VNSGU students

</div>
