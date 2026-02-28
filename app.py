"""
Pakistani Bank Statement PDF to Excel Converter
Supports: HBL, Bank AL Habib, Meezan Bank
"""

import gc
import os
import re
import threading
import uuid
import zipfile
from pathlib import Path
from io import BytesIO

import pandas as pd
import pdfplumber
import pytesseract
from flask import Flask, render_template, request, send_file, jsonify
from PIL import Image
from pdf2image import convert_from_path
from werkzeug.utils import secure_filename

# =============================================================================
# Configuration
# =============================================================================

app = Flask(__name__)
app.config['SECRET_KEY'] = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_FOLDER = BASE_DIR / 'uploads'
UPLOAD_FOLDER.mkdir(exist_ok=True)

app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
ALLOWED_EXTENSIONS = {'pdf'}

SUPPORTED_BANKS = {
    'hbl': 'Habib Bank Limited',
    'alhabib': 'Bank AL Habib',
    'meezan': 'Meezan Bank'
}

# In-memory job store — safe with 1 gunicorn worker (single process)
jobs: dict = {}


# =============================================================================
# Utility Functions
# =============================================================================

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def generate_safe_filename(original_filename: str) -> str:
    secure_name = secure_filename(original_filename)
    if not secure_name:
        secure_name = 'uploaded_file'
    unique_id = str(uuid.uuid4())[:8]
    name, ext = os.path.splitext(secure_name)
    return f"{name}_{unique_id}{ext}"


def clean_numeric_value(value) -> float:
    if value is None or value == '':
        return 0.0
    value = str(value).strip().replace(',', '')
    match = re.search(r'-?[\d.]+', value)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return 0.0
    return 0.0


def detect_bank(text: str) -> str:
    """Auto-detect bank from PDF text content."""
    if not text:
        return 'unsupported'
    text_lower = text.lower()
    
    if 'bank al habib' in text_lower or 'bank alhabib' in text_lower:
        return 'alhabib'
    elif 'meezan bank' in text_lower or 'meezanbank' in text_lower or 'mbl treasur' in text_lower:
        return 'meezan'
    elif 'habib bank limited' in text_lower or 'hbl tower' in text_lower:
        return 'hbl'
    
    return 'unsupported'


# =============================================================================
# PDF Extraction
# =============================================================================

CHUNK_SIZE = 35  # pages per Excel part


def get_pdf_page_count(pdf_path: str) -> int:
    with pdfplumber.open(pdf_path) as pdf:
        return len(pdf.pages)


def _ocr_single_page(pdf_path: str, page_1idx: int) -> str:
    """OCR fallback for a single page (1-indexed)."""
    try:
        images = convert_from_path(pdf_path, dpi=150, first_page=page_1idx, last_page=page_1idx)
        if images:
            text = pytesseract.image_to_string(images[0], lang='eng')
            del images
            gc.collect()
            return text or ''
    except Exception:
        pass
    return ''


def extract_text_with_pdfplumber(pdf_path: str, start: int = 0, end: int = None) -> str:
    """Extract text from pages[start:end] (0-indexed).
    On any page error, OCR is used for that page so nothing is lost.
    No signals used — safe to call from background threads.
    """
    text_content = []
    pdf = None
    try:
        pdf = pdfplumber.open(pdf_path)
        end_idx = end if end is not None else len(pdf.pages)

        for rel_idx, page in enumerate(pdf.pages[start:end_idx]):
            abs_page_1idx = start + rel_idx + 1  # 1-indexed for pdf2image
            page_text = None
            try:
                page_text = page.extract_text()
            except Exception:
                # pdfplumber/pdfminer failed — OCR this page
                page_text = _ocr_single_page(pdf_path, abs_page_1idx)
            finally:
                del page
                gc.collect()

            if page_text:
                text_content.append(page_text)
    except Exception as e:
        raise Exception(f"pdfplumber extraction failed: {str(e)}")
    finally:
        if pdf is not None:
            pdf.close()
    return '\n'.join(text_content)


def extract_text_with_ocr(pdf_path: str, first_page: int, last_page: int) -> str:
    """OCR pages first_page..last_page (1-indexed, inclusive). Processes chunk at once."""
    text_content = []
    try:
        # dpi=150 uses ~4x less memory than dpi=300; still readable for printed statements
        images = convert_from_path(pdf_path, dpi=150, first_page=first_page, last_page=last_page)
        for image in images:
            try:
                page_text = pytesseract.image_to_string(image, lang='eng')
                if page_text:
                    text_content.append(page_text)
            except Exception:
                continue
            finally:
                del image  # release memory immediately
    except Exception as e:
        raise Exception(f"OCR extraction failed: {str(e)}")
    return '\n'.join(text_content)


def extract_chunk_text(pdf_path: str, start: int, end: int) -> str:
    """Try pdfplumber first, fall back to OCR for pages[start:end] (0-indexed, end exclusive)."""
    try:
        text = extract_text_with_pdfplumber(pdf_path, start, end)
        if text and len(text.strip()) > 50:
            return text
    except Exception:
        pass
    try:
        # pdf2image uses 1-indexed pages
        return extract_text_with_ocr(pdf_path, start + 1, end)
    except Exception as e:
        raise Exception(f"All extraction methods failed for pages {start+1}-{end}: {str(e)}")


# =============================================================================
# Bank-Specific Parsers
# =============================================================================

def parse_hbl(text: str) -> list:
    """
    Parse Habib Bank Limited (HBL) statement.
    Format: |DATE|VALUE|PARTICULARS|DEBIT|CREDIT|BALANCE|
    """
    transactions = []
    lines = text.split('\n')
    date_pattern = r'(\d{1,2}(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\d{2})'
    current_txn = None

    for line in lines:
        if not line.strip():
            continue
        if 'BROUGHT FORWARD' in line.upper():
            continue
        if 'DATE' in line.upper() and 'PARTICULARS' in line.upper():
            continue
        if 'ACCOUNT STATEMENT' in line.upper():
            continue
        if '---' in line or 'Continue on next page' in line:
            continue
        if 'System Generated' in line or 'does not require' in line:
            continue

        date_match = re.search(date_pattern, line, re.IGNORECASE)

        if date_match:
            if current_txn:
                transactions.append(current_txn)

            cols = [c.strip() for c in line.split('|')]
            current_txn = {
                'date': parse_hbl_date(date_match.group(1)),
                'particulars_list': [],
                'credit': '',
                'debit': '',
                'bank': 'Habib Bank Limited'
            }

            # cols[3] = PARTICULARS, cols[4] = DEBIT, cols[5] = CREDIT
            particulars = cols[3] if len(cols) > 3 else ''
            debit_col = cols[4] if len(cols) > 4 else ''
            credit_col = cols[5] if len(cols) > 5 else ''

            if particulars:
                current_txn['particulars_list'].append(particulars)

            # Extract amounts
            amount_pattern = r'[\d,]+\.?\d*'
            if credit_col:
                amt_match = re.search(amount_pattern, credit_col)
                if amt_match:
                    current_txn['credit'] = amt_match.group()

            if debit_col:
                amt_match = re.search(amount_pattern, debit_col)
                if amt_match:
                    current_txn['debit'] = amt_match.group()

        elif current_txn and line.strip():
            cols = [c.strip() for c in line.split('|')]
            if len(cols) > 3 and cols[3]:
                current_txn['particulars_list'].append(cols[3])

    if current_txn:
        transactions.append(current_txn)

    return transactions


def parse_hbl_date(date_str: str) -> str:
    """Parse HBL date format (DDMMMYY) to YYYY-MM-DD."""
    if not date_str:
        return ''
    date_str = date_str.strip().upper()
    hbl_pattern = r'(\d{1,2})(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)(\d{2})'
    match = re.match(hbl_pattern, date_str)
    if match:
        day = int(match.group(1))
        month_str = match.group(2)
        year = int(match.group(3))
        months = {
            'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
            'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
        }
        month = months.get(month_str, 1)
        year += 2000
        return f"{year:04d}-{month:02d}-{day:02d}"
    return normalize_date_fallback(date_str)


def normalize_date_fallback(date_str: str) -> str:
    """Normalize date to YYYY-MM-DD format (fallback)."""
    if not date_str:
        return ''
    date_str = date_str.strip()
    date_patterns = [
        (r'(\d{1,2})/(\d{1,2})/(\d{2,4})', '%d/%m/%Y'),
        (r'(\d{1,2})-(\d{1,2})-(\d{2,4})', '%d-%m-%Y'),
    ]
    for pattern, fmt in date_patterns:
        match = re.search(pattern, date_str)
        if match:
            try:
                parts = re.findall(r'\d+', match.group())
                if len(parts) >= 3:
                    day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
                    if year < 100:
                        year += 2000 if year < 50 else 1900
                    if 1 <= day <= 31 and 1 <= month <= 12:
                        return f"{year:04d}-{month:02d}-{day:02d}"
            except (ValueError, IndexError):
                continue
    return date_str


def parse_meezan(text: str) -> list:
    """
    Parse Meezan Bank statement.
    Format: DD/MM/YY DD/MM/YY Doc.No Particulars Amount
    Negative amounts = Debit, Positive = Credit
    
    Extracts ONLY numbers or text+number combinations from particulars.
    Pure text words are skipped.
    """
    transactions = []
    lines = text.split('\n')
    date_pattern = r'(\d{2}/\d{2}/\d{2,4})'
    amount_pattern = r'-?[\d,]+\.?\d*'
    current_txn = None
    
    skip_patterns = [
        'OPENING BALANCE', 'CLOSING BALANCE', 'Generated By', 'Print Date',
        'STATEMENT OF', 'IBAN', 'Account No', 'Total No', 'Total Credit',
        'Total Debit', 'Turnover', '___', '==', 'page', 'Generated'
    ]

    def is_number_or_alphanumeric(token):
        """Check if token is a number or contains numbers (like 13D, 1042, etc.)"""
        # Pure numbers
        if re.match(r'^\d+$', token):
            return True
        # Alphanumeric (contains both letters and numbers like 13D, A123, etc.)
        if re.match(r'^[A-Z]*\d+[A-Z]*$', token, re.IGNORECASE):
            return True
        # Numbers in parentheses like (1042)
        if re.match(r'^\(\d+\)$', token):
            return True
        return False

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue
        
        if any(skip in line_stripped.upper() for skip in skip_patterns):
            continue
        if 'https://' in line_stripped or 'servlet' in line_stripped:
            continue
        if line_stripped.startswith('=') or line_stripped.startswith('_'):
            continue
        if '<=' in line_stripped and 'BALANCE' in line_stripped:
            continue
        
        date_match = re.match(date_pattern, line_stripped)
        
        if date_match:
            if current_txn:
                transactions.append(current_txn)
            
            rest_of_line = line_stripped[date_match.end():].strip()
            date_str = date_match.group(1)
            parsed_date = parse_meezan_date(date_str)
            
            current_txn = {
                'date': parsed_date,
                'particulars_numbers': [],  # Store only numbers/alphanumeric
                'credit': '',
                'debit': '',
                'bank': 'Meezan Bank'
            }
            
            # Extract amount
            amounts = re.findall(amount_pattern, rest_of_line)
            if amounts:
                amount = amounts[-1]
                try:
                    amt_val = float(amount.replace(',', ''))
                    if amt_val < 0:
                        current_txn['debit'] = str(abs(amt_val))
                    else:
                        current_txn['credit'] = str(amt_val)
                except ValueError:
                    pass
            
            # Clean particulars - remove dates and amounts
            particulars_text = rest_of_line
            second_date_match = re.match(date_pattern, particulars_text)
            if second_date_match:
                particulars_text = particulars_text[second_date_match.end():].strip()
            
            if amounts:
                for amt in reversed(amounts):
                    if amt in particulars_text:
                        idx = particulars_text.rfind(amt)
                        if idx > 0:
                            particulars_text = particulars_text[:idx].strip()
                            break
            
            # Extract tokens from particulars
            if particulars_text:
                tokens = particulars_text.split()
                for token in tokens:
                    token_clean = token.strip('(),.')
                    if token_clean and is_number_or_alphanumeric(token_clean):
                        # Remove parentheses from stored value
                        if token_clean.startswith('(') and token_clean.endswith(')'):
                            token_clean = token_clean[1:-1]
                        current_txn['particulars_numbers'].append(token_clean)
        
        elif current_txn and line_stripped:
            if not any(skip in line_stripped.upper() for skip in skip_patterns):
                if 'https://' not in line_stripped and not line_stripped.startswith('='):
                    if not re.match(r'^\d{2}/\d{2}/\d{2}', line_stripped):
                        if not line_stripped.startswith('_'):
                            if 'AM' not in line_stripped and 'PM' not in line_stripped:
                                if 'Account Statement' not in line_stripped:
                                    # Extract numbers/alphanumeric from continuation line
                                    tokens = line_stripped.split()
                                    for token in tokens:
                                        token_clean = token.strip('(),.')
                                        if token_clean and is_number_or_alphanumeric(token_clean):
                                            if token_clean.startswith('(') and token_clean.endswith(')'):
                                                token_clean = token_clean[1:-1]
                                            current_txn['particulars_numbers'].append(token_clean)

    if current_txn:
        transactions.append(current_txn)

    return transactions


def parse_meezan_date(date_str: str) -> str:
    """Parse Meezan date format (DD/MM/YY) to YYYY-MM-DD."""
    if not date_str:
        return ''
    date_str = date_str.strip()
    
    match = re.match(r'(\d{2})/(\d{2})/(\d{2})', date_str)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        year += 2000 if year < 50 else 1900
        return f"{year:04d}-{month:02d}-{day:02d}"
    
    match = re.match(r'(\d{2})/(\d{2})/(\d{4})', date_str)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        return f"{year:04d}-{month:02d}-{day:02d}"
    
    return date_str


def parse_alhabib(text: str) -> list:
    """Parse Bank AL Habib statement."""
    transactions = []
    lines = text.split('\n')
    current_transaction = None
    slip_buffer = []
    date_pattern = r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})'
    amount_pattern = r'[\d,]+\.?\d*'

    for line in lines:
        line = line.strip()
        if not line:
            continue
        date_match = re.search(date_pattern, line)
        if date_match:
            if current_transaction:
                current_transaction['particulars_list'] = slip_buffer.copy()
                transactions.append(current_transaction)
                slip_buffer = []
            current_transaction = {
                'date': normalize_date_fallback(date_match.group(1)),
                'particulars_list': [],
                'credit': '',
                'debit': '',
                'bank': 'Bank AL Habib'
            }
            amounts = re.findall(amount_pattern, line)
            for amount in amounts:
                if 'cr' in line.lower() or '+' in line:
                    current_transaction['credit'] = amount
                elif 'dr' in line.lower() or '-' in line:
                    current_transaction['debit'] = amount
                else:
                    if not current_transaction['credit']:
                        current_transaction['credit'] = amount
            slip_buffer.append(line)
        elif current_transaction:
            slip_buffer.append(line)

    if current_transaction:
        current_transaction['particulars_list'] = slip_buffer.copy()
        transactions.append(current_transaction)

    return transactions


def parse_statement(text: str, bank_code: str) -> list:
    """Route to appropriate bank parser."""
    parsers = {'hbl': parse_hbl, 'alhabib': parse_alhabib, 'meezan': parse_meezan}
    parser_func = parsers.get(bank_code)
    if not parser_func:
        raise ValueError(f"No parser available for bank: {bank_code}")
    return parser_func(text)


# =============================================================================
# Excel Generation
# =============================================================================

def generate_excel(transactions: list, bank_name: str) -> bytes:
    """
    Generate Excel file from parsed transactions.
    
    For Meezan Bank: Date, Particulars1-4 (numbers/alphanumeric only), Credit, Debit, Bank
    For Other Banks: Date, Particulars1-4 (all text), Credit, Debit, Bank
    """
    data = []
    is_meezan = 'meezan' in bank_name.lower()
    max_particulars = 4

    for txn in transactions:
        if is_meezan:
            # Meezan: Use only numbers/alphanumeric from particulars_numbers
            numbers = txn.get('particulars_numbers', [])[:max_particulars]
            while len(numbers) < max_particulars:
                numbers.append('')
            
            row = {
                'Date': txn.get('date', ''),
                'Particulars1': numbers[0],
                'Particulars2': numbers[1],
                'Particulars3': numbers[2],
                'Particulars4': numbers[3],
                'Credit': clean_numeric_value(txn.get('credit', '')),
                'Debit': clean_numeric_value(txn.get('debit', '')),
                'Bank': bank_name
            }
        else:
            # Other banks: Flatten all particulars text
            particulars_list = txn.get('particulars_list', [])
            all_parts = []
            for p in particulars_list:
                parts = [part.strip() for part in p.split() if part.strip()]
                all_parts.extend(parts)
            
            all_parts = all_parts[:max_particulars]
            while len(all_parts) < max_particulars:
                all_parts.append('')
            
            row = {
                'Date': txn.get('date', ''),
                'Particulars1': all_parts[0],
                'Particulars2': all_parts[1],
                'Particulars3': all_parts[2],
                'Particulars4': all_parts[3],
                'Credit': clean_numeric_value(txn.get('credit', '')),
                'Debit': clean_numeric_value(txn.get('debit', '')),
                'Bank': bank_name
            }
        
        data.append(row)

    df = pd.DataFrame(data)
    columns = ['Date', 'Particulars1', 'Particulars2', 'Particulars3', 'Particulars4', 'Credit', 'Debit', 'Bank']
    df = df[columns]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Transactions', index=False)
    output.seek(0)
    return output.getvalue()


# =============================================================================
# Flask Routes
# =============================================================================

@app.route('/')
def index():
    return render_template('index.html')


def process_pdf_job(job_id: str, file_path: str):
    """Runs in a background thread. Processes PDF and updates jobs[job_id]."""
    try:
        jobs[job_id]['status'] = 'processing'

        try:
            total_pages = get_pdf_page_count(file_path)
        except Exception as e:
            jobs[job_id] = {'status': 'error', 'error': f'Could not read PDF: {str(e)}'}
            return

        jobs[job_id]['total_pages'] = total_pages

        bank_code = None
        bank_name = None
        all_excel_parts = []
        total_transactions = 0

        for part_num, start in enumerate(range(0, total_pages, CHUNK_SIZE), 1):
            end = min(start + CHUNK_SIZE, total_pages)
            jobs[job_id]['progress'] = f'Processing pages {start + 1}–{end} of {total_pages}...'

            try:
                chunk_text = extract_chunk_text(file_path, start, end)
            except Exception:
                gc.collect()
                continue

            if not chunk_text or len(chunk_text.strip()) < 10:
                continue

            if bank_code is None:
                bank_code = detect_bank(chunk_text)
                if bank_code == 'unsupported':
                    jobs[job_id] = {'status': 'error', 'error': 'Unsupported bank. Supports HBL, Bank AL Habib, Meezan Bank.'}
                    return
                bank_name = SUPPORTED_BANKS.get(bank_code, bank_code)

            try:
                transactions = parse_statement(chunk_text, bank_code)
            except Exception:
                continue
            finally:
                del chunk_text
                gc.collect()

            if not transactions:
                continue

            try:
                excel_bytes = generate_excel(transactions, bank_name)
                safe_bank = bank_name.replace(' ', '_')
                all_excel_parts.append((f"Part{part_num}_{safe_bank}.xlsx", excel_bytes))
                total_transactions += len(transactions)
            except Exception:
                continue
            finally:
                del transactions
                gc.collect()

        if not all_excel_parts:
            jobs[job_id] = {'status': 'error', 'error': 'No transactions found in the PDF.'}
            return

        uid = uuid.uuid4().hex[:8]

        if len(all_excel_parts) == 1:
            out_filename = f"bank_statement_{uid}.xlsx"
            out_path = os.path.join(app.config['UPLOAD_FOLDER'], out_filename)
            with open(out_path, 'wb') as f:
                f.write(all_excel_parts[0][1])
            jobs[job_id] = {
                'status': 'done',
                'bank_detected': bank_name,
                'transactions_count': total_transactions,
                'parts_count': 1,
                'download_filename': out_filename,
                'is_zip': False
            }
        else:
            zip_filename = f"bank_statement_{uid}.zip"
            zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for name, data in all_excel_parts:
                    zf.writestr(name, data)
            jobs[job_id] = {
                'status': 'done',
                'bank_detected': bank_name,
                'transactions_count': total_transactions,
                'parts_count': len(all_excel_parts),
                'download_filename': zip_filename,
                'is_zip': True
            }

    except Exception as e:
        jobs[job_id] = {'status': 'error', 'error': f'Unexpected error: {str(e)}'}
    finally:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass
        gc.collect()


@app.route('/convert', methods=['POST'])
def convert_pdf():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file selected'}), 400
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': 'Invalid file type. Only PDF files are allowed.'}), 400

    try:
        safe_filename = generate_safe_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
        file.save(file_path)

        job_id = uuid.uuid4().hex
        jobs[job_id] = {'status': 'queued'}

        thread = threading.Thread(target=process_pdf_job, args=(job_id, file_path), daemon=True)
        thread.start()

        # Return immediately — Render's 30s proxy timeout is no longer an issue
        return jsonify({'success': True, 'job_id': job_id})

    except Exception as e:
        return jsonify({'success': False, 'error': f'Upload failed: {str(e)}'}), 500


@app.route('/status/<job_id>')
def job_status(job_id: str):
    job = jobs.get(job_id)
    if job is None:
        return jsonify({'status': 'not_found'}), 404
    return jsonify(job)


@app.route('/download/<filename>')
def download_file(filename: str):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if not os.path.exists(file_path):
            return jsonify({'success': False, 'error': 'File not found'}), 404
        if filename.endswith('.zip'):
            mimetype = 'application/zip'
        else:
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return send_file(file_path, as_attachment=True, download_name=filename, mimetype=mimetype)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.errorhandler(413)
def too_large(e):
    return jsonify({'success': False, 'error': 'File too large. Max 16MB.'}), 413


@app.errorhandler(500)
def server_error(e):
    return jsonify({'success': False, 'error': 'Internal server error.'}), 500


if __name__ == '__main__':
    UPLOAD_FOLDER.mkdir(exist_ok=True)
    # Production mode - disable debug and reloader
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)
