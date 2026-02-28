FROM python:3.11-slim

# Install system dependencies (poppler for pdf2image, tesseract for OCR)
RUN apt-get update && apt-get install -y \
    poppler-utils \
    tesseract-ocr \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# HuggingFace Spaces requires a non-root user with UID 1000
RUN useradd -m -u 1000 user
USER user

ENV HOME=/home/user \
    PATH=/home/user/.local/bin:$PATH

WORKDIR /home/user/app

# Install Python dependencies
COPY --chown=user requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY --chown=user app.py .
COPY --chown=user templates/ templates/
COPY --chown=user static/ static/

# Create uploads directory (writable by user)
RUN mkdir -p uploads

# HuggingFace Spaces default port
EXPOSE 7860

CMD ["gunicorn", "--bind", "0.0.0.0:7860", "--workers", "1", "--timeout", "600", "--threads", "4", "app:app"]
