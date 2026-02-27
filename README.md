# PDF to Excel Converter - Pakistani Banks

A production-ready web application that converts Pakistani bank statement PDFs into structured Excel files.

## Supported Banks

- **Habib Bank Limited (HBL)**
- **Bank AL Habib**
- **Meezan Bank**

## Features

- ✅ Auto bank detection from PDF content
- ✅ Drag & Drop PDF upload
- ✅ OCR fallback for image-based PDFs
- ✅ Deposit slip number splitting (up to 4 columns)
- ✅ Clean Excel export with structured columns
- ✅ Modern Bootstrap 5 UI
- ✅ Secure file handling
- ✅ Docker support

## Installation

### Prerequisites

- Python 3.10 or higher
- pip (Python package manager)

### 1. Install Tesseract OCR

**Windows:**
1. Download installer from: https://github.com/UB-Mannheim/tesseract/wiki
2. Run the installer (default location: `C:\Program Files\Tesseract-OCR`)
3. Add Tesseract to system PATH or set environment variable:
   ```
   setx TESSDATA_PREFIX "C:\Program Files\Tesseract-OCR\tessdata"
   ```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr poppler-utils
```

**macOS:**
```bash
brew install tesseract poppler
```

### 2. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the Application

```bash
python app.py
```

The application will start at: http://localhost:5000

---

## Docker Deployment

### Build the Docker Image

```bash
docker build -t pdf-to-excel-converter .
```

### Run the Container

```bash
docker run -p 5000:5000 pdf-to-excel-converter
```

### Using Docker Compose

Create `docker-compose.yml`:
```yaml
version: '3.8'
services:
  pdf-converter:
    build: .
    ports:
      - "5000:5000"
    volumes:
      - ./uploads:/app/uploads
    restart: unless-stopped
```

Run with:
```bash
docker-compose up -d
```

---

## Project Structure

```
project/
│
├── app.py                 # Flask backend application
├── requirements.txt       # Python dependencies
├── Dockerfile            # Docker configuration
├── templates/
│   └── index.html        # Frontend UI (Bootstrap 5)
├── static/
│   └── style.css         # Custom styles
└── uploads/              # Temporary file storage
```

---

## Excel Output Format

| Date | Slip1 | Slip2 | Slip3 | Slip4 | Credit | Debit | Bank |
|------|-------|-------|-------|-------|--------|-------|------|
| 2024-01-15 | ABC123 | XYZ789 | | | 50000 | 0 | Habib Bank Limited |

---

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/` | Main UI page |
| POST | `/convert` | Upload and convert PDF |
| GET | `/download/<filename>` | Download converted Excel |

---

## Security Features

- ✅ File type validation (PDF only)
- ✅ Secure filename generation (prevents injection)
- ✅ File size limit (16MB max)
- ✅ Automatic file cleanup after processing
- ✅ Non-root user in Docker container

---

## Troubleshooting

### OCR Not Working

Ensure Tesseract is installed and in your PATH:
```bash
tesseract --version
```

### PDF Extraction Fails

- Ensure pdfplumber is installed: `pip install pdfplumber`
- For image-based PDFs, OCR fallback will be used automatically

### Port Already in Use

Change the port in `app.py`:
```python
app.run(debug=True, host='0.0.0.0', port=8080)
```

---

## License

MIT License
