# Render.com Dockerfile
# PDF to Excel Converter - Pakistani Banks

FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PIP_NO_CACHE_DIR=1

# Set working directory
WORKDIR /app

# Install system dependencies (updated package names)
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    tesseract-ocr \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

# Copy requirements
COPY requirements.txt .

# Install Python packages
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY app.py .
COPY templates/ templates/
COPY static/ static/

# Create uploads
RUN mkdir -p uploads

# Expose port (Render uses PORT env variable)
EXPOSE $PORT

# Run with Gunicorn
CMD ["gunicorn", "--bind", f"0.0.0.0:${PORT:-5000}", "--workers", "2", "--timeout", "120", "app:app"]
