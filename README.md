---
title: PDF to Excel Converter - Pakistani Banks
emoji: 📊
colorFrom: blue
colorTo: green
sdk: docker
app_port: 7860
pinned: false
license: mit
---

# 📊 PDF to Excel Converter - Pakistani Banks

A complete web application that converts Pakistani bank statement PDFs into structured Excel files.

## 🏦 Supported Banks

- ✅ **Habib Bank Limited (HBL)**
- ✅ **Bank AL Habib**
- ✅ **Meezan Bank**

## ✨ Features

- **Auto Bank Detection** - Automatically detects bank from PDF content
- **OCR Fallback** - Advanced OCR for image-based/scanned PDFs
- **Smart Number Extraction** - Extracts deposit slip numbers, branch codes
- **Clean Excel Output** - Structured columns with proper formatting
- **Modern UI** - Bootstrap 5 responsive design

## 🚀 How to Use

1. **Upload PDF** - Drag & drop or browse your bank statement PDF
2. **Click Convert** - Application will auto-detect the bank
3. **Download Excel** - Get structured Excel file instantly

## 📁 Excel Output Format

### Meezan Bank (Numbers Only)
| Date | Particulars1 | Particulars2 | Particulars3 | Particulars4 | Credit | Debit |
|------|--------------|--------------|--------------|--------------|--------|-------|
| 2025-12-02 | 1042 | 13D | 2 | 4171966 | 2250000 | 0 |

### HBL / Bank AL Habib (Text)
| Date | Particulars1 | Particulars2 | Particulars3 | Particulars4 | Credit | Debit |
|------|--------------|--------------|--------------|--------------|--------|-------|
| 2025-12-01 | Online | Deposit | 25687849 | MUHAMMAD | 158000 | 0 |

## 💻 Local Installation

```bash
# Clone the repository
git clone https://huggingface.co/spaces/MSK9218/MSK-Pdf-to-Excel-v1
cd MSK-Pdf-to-Excel-v1

# Install dependencies
pip install -r requirements.txt

# Run locally
python app.py
```

Access: **http://127.0.0.1:5000**

## 🐳 Docker

```bash
docker build -t pdf-excel-converter .
docker run -p 5000:5000 pdf-excel-converter
```

## 📋 Requirements

- Python 3.10+
- Tesseract OCR
- Poppler-utils

## 🛠️ Technologies

- **Backend:** Flask
- **Frontend:** Bootstrap 5
- **PDF Processing:** pdfplumber, pytesseract
- **Excel:** pandas, openpyxl

## 📝 License

MIT License

## 👨‍💻 Author

**Muhammad Saad Ahmed**

GitHub: [@Muhammad-Saad-Ahmed](https://github.com/Muhammad-Saad-Ahmed)

---

**Built with ❤️ for Pakistani Banking Community**
