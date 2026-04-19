from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import os
import uuid
from werkzeug.utils import secure_filename
import PyPDF2
import pdfplumber
import pandas as pd
import pytesseract
from PIL import Image
import io
import tempfile
import shutil

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = os.path.abspath('uploads')
app.config['PROCESSED_FOLDER'] = os.path.abspath('processed')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Ensure upload directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

# Configure Tesseract path (adjust for your system)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(pdf_path):
    """Extract text from PDF using PyPDF2"""
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                text += page.extract_text() or ""
    except Exception as e:
        print(f"Error extracting text with PyPDF2: {e}")
        # Fallback to pdfplumber
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text += page.extract_text() or ""
        except Exception as e2:
            print(f"Error extracting text with pdfplumber: {e2}")
            return ""

    return text

def extract_tables_from_pdf(pdf_path):
    """Extract tables from PDF using pdfplumber"""
    tables_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table_idx, table in enumerate(tables):
                    if table:
                        # Convert to DataFrame for easier handling
                        df = pd.DataFrame(table)
                        tables_data.append({
                            'page': page_num + 1,
                            'table_index': table_idx + 1,
                            'data': df
                        })
    except Exception as e:
        print(f"Error extracting tables: {e}")

    return tables_data

def perform_ocr_on_image(image):
    """Perform OCR on PIL Image"""
    try:
        text = pytesseract.image_to_string(image)
        return text
    except Exception as e:
        print(f"OCR error: {e}")
        return ""

def extract_text_with_ocr(pdf_path):
    """Extract text from scanned PDF using OCR"""
    text = ""
    try:
        # Use pdfplumber to convert pages to images for OCR
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # Convert page to image
                img = page.to_image(resolution=300).original
                page_text = perform_ocr_on_image(img)
                text += f"\n--- Page {page_num + 1} ---\n{page_text}"
    except Exception as e:
        print(f"OCR extraction error: {e}")

    return text

def parse_invoice_data(text, pdf_path=None):
    """Parse common invoice fields from extracted text and tables"""
    import re

    def extract_hsn_from_row(row, preferred_idx=None):
        """Find an HSN/SAC-like value anywhere in a table row."""
        candidates = []

        def add_candidate(value):
            if value is None:
                return
            text_value = str(value).strip()
            if not text_value:
                return
            lowered = text_value.lower()
            if any(keyword in lowered for keyword in ['description', 'qty', 'quantity', 'rate', 'amount', 'cgst', 'sgst', 'discount', 'total']):
                return

            # Prefer standalone 2-8 digit values, which matches most HSN/SAC codes.
            for match in re.findall(r'\b\d{2,8}\b', text_value.replace(',', '')):
                candidates.append(match)

        if preferred_idx is not None and isinstance(preferred_idx, int) and 0 <= preferred_idx < len(row):
            add_candidate(row[preferred_idx])

        for cell in row:
            add_candidate(cell)

        for candidate in candidates:
            cleaned = re.sub(r'\D', '', str(candidate))
            if 2 <= len(cleaned) <= 8:
                return cleaned

        return ''

    def is_summary_row(row):
        """Check if a row is a summary/total row and should be skipped."""
        row_str = ' '.join(str(cell).lower() for cell in row if cell)
        
        # Check for summary keywords
        summary_keywords = ['total', 'subtotal', 'balance', 'grand total', 'net total', 'sub total', 
                          'rounding', 'amount due', 'amount payable']
        if any(keyword in row_str for keyword in summary_keywords):
            return True
        
        # Check if row has minimal item data (no quantity or all zeros/empty)
        # This filters out rows where qty is empty/zero and amount looks like a total
        qty_empty = True
        for i, cell in enumerate(row):
            cell_str = str(cell).strip().lower()
            # If any cell looks like a quantity value (non-zero number), it's likely an item row
            if cell and any(keyword in cell_str for keyword in ['qty', 'quantity']):
                continue
            if cell_str and re.search(r'^\d+(\.\d+)?$', cell_str):
                try:
                    val = float(cell_str)
                    if val > 0 and val < 100:  # Likely a quantity or rate
                        qty_empty = False
                        break
                except:
                    pass
        
        # Row with all zeros or mostly empty values across item columns is likely summary
        empty_count = sum(1 for cell in row if not str(cell).strip() or str(cell).strip() == '0')
        if empty_count > len(row) * 0.6:  # More than 60% empty/zero
            return True
        
        return False

    # First extract common fields that apply to the entire invoice
    common_data = {}

    # Common patterns for invoice header data
    header_patterns = {
        'invoice_no': r'(?:#\s*:?\s*|Invoice\s*#?\s*:?\s*|INVOICE\s*NO\.?\s*:?\s*)([A-Z0-9\-/]+)',
        'date': r'(?:Invoice\s*Date\s*:?\s*|DATE\s*:?\s*|Date\s*:?\s*)([\d/.\-]+)',
        # 'receiver_name': r'(?:Bill\s*To|Receiver|Billed\s*to)\s*[:.-]\s*([A-Za-z0-9\s.,]+)', # Commented out
        # 'receiver_gst': r'(?:GSTIN\s*:?\s*|GST\s*#?\s*:?\s*)([A-Z0-9]{15})', # Commented out - too generic, picks up sender GST
        'taxrate': r'(?:CGST|SGST)\s*(\d+(?:\.\d+)?)\s*%',
        'cgst': r'CGST\d*\s*\(\d+%\)\s*([\d,]+\.?\d*)',
        'sgst': r'SGST\d*\s*\(\d+%\)\s*([\d,]+\.?\d*)',
        'invoice_value': r'(?:Total\s*:?\s*|TOTAL\s*:?\s*)(?:Rs\.?|INR|₹)?\s*([\d,]+\.?\d*)',
        'total_invoice_value': r'(?:Balance\s*Due\s*:?\s*|GRAND\s*TOTAL\s*:?\s*|FINAL\s*TOTAL\s*:?\s*)(?:Rs\.?|INR|₹)?\s*([\d,]+\.?\d*)'
    }

    # Extract header data using regex patterns
    for field, pattern in header_patterns.items():
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            value = match.group(1).strip()
            # Clean up the value
            value = re.sub(r'[^\w\s.,/-]', '', value)  # Remove special characters except common ones
            common_data[field] = value

    # Specialized extraction for Receiver Name
    # Matches "Bill To" followed optionally by : or . and surrounding whitespace/newlines, then captures the next meaningful lines (up to 3)
    # This handles cases where the next line is "Ship To" (interleaved headers) or empty
    rec_name_pattern = r'(?:Bill\s*To|Receiver|Billed\s*to)\s*[:\.]?\s*((?:[^\r\n]+[\r\n]*){1,4})'
    rec_name_match = re.search(rec_name_pattern, text, re.IGNORECASE)
    if rec_name_match:
        # Get the block of text following "Bill To"
        candidate_lines = rec_name_match.group(1).splitlines()
        for line in candidate_lines:
            clean_name = line.strip()
            # Skip if it's too short, or is a known header keyword like "Ship To", "GSTIN", "Date"
            # converting to lower for checks
            check = clean_name.lower()
            if (len(clean_name) > 2 and 
                not any(x in check for x in ['gstin', 'invoice', 'date', 'ship to', 'shipment', 'place of supply', 'terms'])):
                 # Further cleanup: sometimes "Receiver: Name" is captured, we want just "Name"
                 clean_name = re.sub(r'^(?:M/s|Mr\.|Mrs\.|Dr\.)\s*', '', clean_name, flags=re.IGNORECASE)
                 common_data['receiver_name'] = clean_name
                 break # Found a valid name line, stop looking

    # Specialized extraction for Receiver GST (prioritize Bill To section)
    # Strategy: Look for "Bill To" followed by "GSTIN" within reasonable distance
    receiver_gst_match = re.search(r'(?:Bill\s*To|Receiver|Billed\s*to)[\s\S]{0,500}?(?:GSTIN|GST\s*#?)[\s.:]*([A-Z0-9]{15})', text, re.IGNORECASE | re.MULTILINE)
    if receiver_gst_match:
         common_data['receiver_gst'] = receiver_gst_match.group(1).strip()
    else:
        # Fallback: Find ALL GSTs and if there are multiple, assume the second one is receiver
        # If there's only one, it might be the only one present.
        all_gsts = re.findall(r'(?:GSTIN\s*:?\s*|GST\s*#?\s*:?\s*)([A-Z0-9]{15})', text, re.IGNORECASE)
        # 09BBDPY4789B1Z5 is the sender's GST (hardcoded check only if strictly necessary, strictly context is better)
        if all_gsts:
             if len(all_gsts) > 1:
                 # Usually first is sender, second is receiver
                 common_data['receiver_gst'] = all_gsts[1]
             else:
                 common_data['receiver_gst'] = all_gsts[0]

    # Extract table data to get all line items
    all_items = []
    if pdf_path:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table:
                            continue

                        # Try to identify column indices from a header row
                        # Scan first few rows to find a likely header row
                        header_row = None
                        header_row_idx = -1
                        
                        for idx, row in enumerate(table[:10]): # Look at first 10 rows
                            row_str = str(row).lower()
                            # Check for key columns that signify a header
                            if 'description' in row_str or 'qty' in row_str or ('hsn' in row_str and 'sac' in row_str):
                                header_row = row
                                header_row_idx = idx
                                break
                        
                        if header_row is None:
                             header_row = table[0]
                             header_row_idx = 0

                        col_map = {}
                        
                        # Helper to find column index from keywords
                        def find_col_idx(headers, keywords):
                            for idx, h in enumerate(headers):
                                if h and any(k in str(h).lower() for k in keywords):
                                    return idx
                            return None

                        # Map columns dynamically based on headers
                        col_map['sno'] = find_col_idx(header_row, ['s.no', 'serial', 'sr.', 'no.'])
                        col_map['hsn'] = find_col_idx(header_row, ['hsn', 'sac', 'code'])
                        col_map['qty'] = find_col_idx(header_row, ['qty', 'quantity', 'unit'])
                        col_map['rate'] = find_col_idx(header_row, ['rate', 'taxable', 'price'])
                        col_map['discount'] = find_col_idx(header_row, ['disc', 'less'])
                        
                        # Special handling for Tax columns: prefer "Amount" or "Amt" specific columns if available
                        # But often they are just labeled "CGST", "SGST" and split into Rate/Amt
                        col_map['cgst'] = find_col_idx(header_row, ['cgst'])
                        col_map['sgst'] = find_col_idx(header_row, ['sgst'])
                        col_map['amount'] = find_col_idx(header_row, ['amount', 'total', 'value'])

                        # Check if sub-headers exist (row after header) for Tax Amount
                        if header_row_idx + 1 < len(table):
                             sub_header = table[header_row_idx + 1]
                             # If we found a CGST column, check if 'Amt' is near it in sub-header
                             if col_map['cgst'] is not None:
                                 # Look at columns starting from cgst idx
                                 for offset in range(3): # Check current and next 2 cols
                                     chk_idx = col_map['cgst'] + offset
                                     if chk_idx < len(sub_header):
                                         val = str(sub_header[chk_idx]).lower()
                                         if 'amt' in val or 'amount' in val:
                                             col_map['cgst_amt_idx'] = chk_idx
                                             break
                             
                             if col_map['sgst'] is not None:
                                 for offset in range(3):
                                     chk_idx = col_map['sgst'] + offset
                                     if chk_idx < len(sub_header):
                                         val = str(sub_header[chk_idx]).lower()
                                         if 'amt' in val or 'amount' in val:
                                             col_map['sgst_amt_idx'] = chk_idx
                                             break
                        
                        # If we didn't find clear headers (especially HSN), likely the specific original format
                        # or a format with complex headers. Fallback to original hardcoded indices for backward compatibility
                        # but check length effectively.
                        if col_map.get('hsn') is None:
                            col_map = {
                                'sno': 1, 'hsn': 3, 'qty': 4, 'rate': 5, 
                                'discount': 6, 'cgst': 10, 'sgst': 12, 'amount': 13
                            }

                        for row_idx, row in enumerate(table):
                            # Skip header rows
                            if row_idx <= header_row_idx:
                                continue
                            # Skip sub-header row if it looks like one (contains 'Amt', '%')
                            if row_idx == header_row_idx + 1 and any(x in str(row).lower() for x in ['amt', '%']):
                                continue
                            # Skip summary/total rows
                            if is_summary_row(row):
                                continue

                            # Safe extraction helper
                            def get_val(idx):
                                if idx is not None and isinstance(idx, int) and 0 <= idx < len(row):
                                    return row[idx]
                                return ''

                            hsn_val = get_val(col_map.get('hsn'))
                            if not hsn_val or not re.search(r'\d{2,8}', str(hsn_val)):
                                hsn_val = extract_hsn_from_row(row, col_map.get('hsn'))

                            qty_val = get_val(col_map.get('qty'))
                            rate_val = get_val(col_map.get('rate'))
                            discount_val = get_val(col_map.get('discount'))
                            
                            # Use found amount indices if available, else fallback to standard
                            cgst_val = get_val(col_map.get('cgst_amt_idx', col_map.get('cgst')))
                            sgst_val = get_val(col_map.get('sgst_amt_idx', col_map.get('sgst')))
                            
                            # Heuristic fix: If extracted GST value looks like a Rate (e.g. "9", "9%") and not an Amount,
                            # and we didn't find an explicit 'Amt' column, try to look ahead for the amount.
                            if not col_map.get('cgst_amt_idx') and cgst_val and re.match(r'^\d+(\.\d+)?\s*%?$', str(cgst_val).strip()):
                                # Current value is likely a rate. Check next 2 columns for an amount.
                                base_idx = col_map.get('cgst')
                                if base_idx is not None:
                                    for offset in [1, 2]:
                                        next_val = get_val(base_idx + offset)
                                        if next_val and re.match(r'[\d,]+\.\d{2}', str(next_val).strip()):
                                            cgst_val = next_val
                                            break

                            if not col_map.get('sgst_amt_idx') and sgst_val and re.match(r'^\d+(\.\d+)?\s*%?$', str(sgst_val).strip()):
                                base_idx = col_map.get('sgst')
                                if base_idx is not None:
                                    for offset in [1, 2]:
                                        next_val = get_val(base_idx + offset)
                                        if next_val and re.match(r'[\d,]+\.\d{2}', str(next_val).strip()):
                                            sgst_val = next_val
                                            break
                            
                            invoice_val = get_val(col_map.get('amount'))
                            sno_val = get_val(col_map.get('sno'))

                            # Check if this row contains valid item data
                            hsn_clean = str(hsn_val).strip().replace('\n', '').replace(' ', '')
                            
                            # Valid item row: has serial number (digits) and HSN (digits)
                            # Relaxed check: if we have HSN and (SNO or some value), consider it valid
                            if (hsn_clean and re.search(r'\d{2,8}', hsn_clean)):
                                item_data = {
                                    'sno': str(sno_val).strip() if sno_val else str(len(all_items) + 1),
                                    'hsn': str(hsn_val).strip(),
                                    'quantity': '',
                                    'taxable_amount': '',
                                    'discount': '',
                                    'taxrate': '18%',  # Default GST rate
                                    'cgst': '',
                                    'sgst': '',
                                    'invoice_value': ''
                                }

                                # Extract quantity
                                if qty_val:
                                    qty_match = re.search(r'([\d,]+\.?\d*)', str(qty_val).strip())
                                    if qty_match:
                                        item_data['quantity'] = qty_match.group(1)

                                # Extract taxable amount (rate)
                                if rate_val:
                                    rate_match = re.search(r'([\d,]+\.?\d*)', str(rate_val).strip())
                                    if rate_match:
                                        item_data['taxable_amount'] = rate_match.group(1)

                                # Extract discount
                                if discount_val:
                                    item_data['discount'] = str(discount_val).strip()

                                # Extract CGST
                                if cgst_val:
                                    cgst_match = re.search(r'([\d,]+\.?\d*)', str(cgst_val).strip())
                                    if cgst_match:
                                        item_data['cgst'] = cgst_match.group(1)

                                # Extract SGST
                                if sgst_val:
                                    sgst_match = re.search(r'([\d,]+\.?\d*)', str(sgst_val).strip())
                                    if sgst_match:
                                        item_data['sgst'] = sgst_match.group(1)

                                # Extract invoice value (amount for this line)
                                if invoice_val:
                                    invoice_match = re.search(r'([\d,]+\.?\d*)', str(invoice_val).strip())
                                    if invoice_match:
                                        item_data['invoice_value'] = invoice_match.group(1)

                                all_items.append(item_data)
        except Exception as e:
            print(f"Error extracting table data: {e}")

    # Aggregation Logic
    if all_items:
        aggregated_items = {}
        for item in all_items:
            hsn = item['hsn']
            if not hsn: # Skip if no HSN
                continue

            # Keep rows separate unless they are effectively identical.
            item_key = (
                hsn,
                str(item.get('quantity', '')).strip(),
                str(item.get('taxable_amount', '')).strip(),
                str(item.get('cgst', '')).strip(),
                str(item.get('sgst', '')).strip(),
                str(item.get('invoice_value', '')).strip()
            )

            if item_key not in aggregated_items:
                aggregated_items[item_key] = {
                    'sno': item['sno'], # Keep first SNO
                    'hsn': hsn,
                    'quantity': 0.0,
                    'taxable_amount': 0.0,
                    'taxrate': item['taxrate'], # Assume same for same HSN
                    'cgst': 0.0,
                    'sgst': 0.0,
                    'invoice_value': 0.0,
                    'count': 0
                }
            
            # Helper to safely parse float
            def parse_float(val):
                try:
                    return float(str(val).replace(',', ''))
                except (ValueError, TypeError):
                    return 0.0

            # Aggregate values
            agg = aggregated_items[item_key]
            agg['quantity'] += parse_float(item['quantity'])
            
            # For taxable_amount aggregation, use the Amount column (invoice_value) instead of Rate
            agg['taxable_amount'] += parse_float(item['invoice_value'])
            
            agg['cgst'] += parse_float(item['cgst'])
            agg['sgst'] += parse_float(item['sgst'])
            agg['invoice_value'] += parse_float(item['invoice_value'])
            agg['count'] += 1


        # Replace all_items with aggregated list
        if aggregated_items:
             new_all_items = []
             calculated_total_value = 0.0
             for idx, (item_key, agg) in enumerate(aggregated_items.items(), 1):
                 
                 # Calculate Invoice Value = Taxable + CGST + SGST (per user request)
                 # Note: taxable_amount here is already summing the 'Amount' column from original data
                 # If original Amount column included tax, adding tax again would double count.
                 # Assuming 'Amount' column was Taxable Value.
                 # User specific request: "in the invoice value cell write the sum of taxable amount cgst sgst"
                 
                 # Since we updated taxable_amount aggregation to sum 'invoice_value' (Amount column),
                 # we should assume that was the Taxable Value. 
                 
                 item_invoice_value = agg['taxable_amount'] + agg['cgst'] + agg['sgst']
                 calculated_total_value += item_invoice_value
                 
                 new_all_items.append({
                     'sno': str(idx),
                     'hsn': item_key[0],
                     'quantity': f"{agg['quantity']:.2f}",
                     'taxable_amount': f"{agg['taxable_amount']:.2f}",
                     # 'discount': '', # Discount removed
                     'taxrate': agg['taxrate'],
                     'cgst': f"{agg['cgst']:.2f}",
                     'sgst': f"{agg['sgst']:.2f}",
                     'invoice_value': f"{item_invoice_value:.2f}"
                 })
             all_items = new_all_items
             
             # Store the calculated total
             common_data['calculated_total_invoice_value'] = f"{calculated_total_value:.2f}"

    # If no table items found, create a single item with empty values
    if not all_items:
        all_items = [{
            'sno': '1',
            'hsn': '',
            'quantity': '',
            'taxable_amount': '',
            'discount': '', # Kept for structure but unused
            'taxrate': '',
            'cgst': '',
            'sgst': '',
            'invoice_value': ''
        }]

    # Create final data list - one entry per item with merged cells logic
    final_data = []
    for i, item in enumerate(all_items):
        if i == 0:
            # First item: include all invoice-level fields
            row_data = {
                'sno': item['sno'],  # Use actual SNO from table
                'invoice_no': common_data.get('invoice_no', ''),
                'date': common_data.get('date', ''),
                'receiver_name': common_data.get('receiver_name', ''),
                'receiver_gst': common_data.get('receiver_gst', ''),
                'hsn': item['hsn'],
                'quantity': item['quantity'],
                'taxable_amount': item['taxable_amount'],
                # 'discount': item.get('discount', ''), # Removed
                'taxrate': item['taxrate'],
                'cgst': item['cgst'],
                'sgst': item['sgst'],
                'invoice_value': item['invoice_value'],
                'total_invoice_value': ''  # Empty for all item rows
            }
        else:
            # Subsequent items: empty invoice fields, filled item fields
            row_data = {
                'sno': item['sno'],  # Use actual SNO from table
                'invoice_no': '',  # Empty
                'date': '',  # Empty
                'receiver_name': '',  # Empty
                'receiver_gst': '',  # Empty
                'hsn': item['hsn'],
                'quantity': item['quantity'],
                'taxable_amount': item['taxable_amount'],
                # 'discount': item.get('discount', ''), # Removed
                'taxrate': item['taxrate'],
                'cgst': item['cgst'],
                'sgst': item['sgst'],
                'invoice_value': item['invoice_value'],
                'total_invoice_value': ''  # Empty for all item rows
            }
        final_data.append(row_data)

    # Add final row with total invoice value
    final_data.append({
        'sno': '',  # Empty
        'invoice_no': '',  # Empty
        'date': '',  # Empty
        'receiver_name': '',  # Empty
        'receiver_gst': '',  # Empty
        'hsn': '',  # Empty
        'quantity': '',  # Empty
        'taxable_amount': '',  # Empty
        # 'discount': '',  # Removed
        'taxrate': '',  # Empty
        'cgst': '',  # Empty
        'sgst': '',  # Empty
        'invoice_value': '',  # Empty
        'total_invoice_value': common_data.get('calculated_total_invoice_value', common_data.get('total_invoice_value', ''))  # Use calculated total if available
    })


    return final_data

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        flash('No file part')
        return redirect(url_for('index'))

    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        flash('No selected files')
        return redirect(url_for('index'))

    processed_files = []
    all_data = []

    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            unique_id = str(uuid.uuid4())
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{filename}")
            file.save(file_path)

            # Process the PDF
            extracted_data = process_pdf(file_path, filename, unique_id)
            processed_files.append(extracted_data)
            all_data.extend(extracted_data.get('parsed_data', []))

    # Always create consolidated Excel file
    excel_filename = f"consolidated_{uuid.uuid4()}.xlsx"
    excel_path = os.path.join(app.config['PROCESSED_FOLDER'], excel_filename)
    create_consolidated_excel(all_data, excel_path)

    return render_template('results.html',
                         processed_files=processed_files,
                         excel_download=excel_filename,
                         consolidated=True,
                         all_data=all_data)

def process_pdf(file_path, original_filename, unique_id):
    """Process a single PDF file, which may contain multiple invoices.
       It splits the PDF if multiple invoices are found."""
    import re
    # Check if there are multiple invoices in this PDF and split them
    split_pdfs = []
    header_pattern = r'(?:#\s*:?\s*|Invoice\s*#?\s*:?\s*|INVOICE\s*NO\.?\s*:?\s*)([A-Z0-9\-/]+)'
    
    try:
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            num_pages = len(reader.pages)
            if num_pages > 1:
                current_writer = PyPDF2.PdfWriter()
                current_invoice_no = None
                current_pages = 0

                for i in range(num_pages):
                    page = reader.pages[i]
                    text = page.extract_text() or ""
                    
                    match = re.search(header_pattern, text, re.IGNORECASE)
                    inv_no = match.group(1).strip() if match else None
                    
                    is_new_invoice = False
                    if inv_no:
                        if current_invoice_no is None:
                            is_new_invoice = True
                        elif inv_no != current_invoice_no:
                            is_new_invoice = True

                    if is_new_invoice and current_pages > 0:
                        temp_fd, temp_path = tempfile.mkstemp(suffix='.pdf', prefix='split_', dir=app.config['UPLOAD_FOLDER'])
                        os.close(temp_fd)
                        with open(temp_path, 'wb') as out_f:
                            current_writer.write(out_f)
                        split_pdfs.append(temp_path)
                        
                        current_writer = PyPDF2.PdfWriter()
                        current_pages = 0
                        current_invoice_no = inv_no
                    elif is_new_invoice:
                        current_invoice_no = inv_no

                    current_writer.add_page(page)
                    current_pages += 1

                if current_pages > 0:
                    temp_fd, temp_path = tempfile.mkstemp(suffix='.pdf', prefix='split_', dir=app.config['UPLOAD_FOLDER'])
                    os.close(temp_fd)
                    with open(temp_path, 'wb') as out_f:
                        current_writer.write(out_f)
                    split_pdfs.append(temp_path)
    except Exception as e:
        print(f"Error checking/splitting PDF: {e}")

    # Process all split files or the original one
    files_to_process = split_pdfs if len(split_pdfs) > 1 else [file_path]
    
    all_parsed_data = []
    all_tables = []
    combined_text = ""
    
    for i, current_file in enumerate(files_to_process):
        # Extract text
        text = extract_text_from_pdf(current_file)

        # If little text found, try OCR
        if len(text.strip()) < 100:
            print(f"Little text found in {current_file}, attempting OCR...")
            text = extract_text_with_ocr(current_file)

        # Extract tables
        tables = extract_tables_from_pdf(current_file)

        # Parse structured data
        parsed_data_list = parse_invoice_data(text, current_file)
        
        all_parsed_data.extend(parsed_data_list)
        all_tables.extend(tables)
        combined_text += f"\n--- Invoice {i+1} ---\n{text}"

    # Create individual Excel file for the whole upload
    excel_filename = f"{unique_id}_extracted.xlsx"
    excel_path = os.path.join(app.config['PROCESSED_FOLDER'], excel_filename)

    create_excel_file(all_parsed_data, all_tables, combined_text, excel_path)

    return {
        'filename': original_filename,
        'unique_id': unique_id,
        'text': combined_text[:1000] + "..." if len(combined_text) > 1000 else combined_text,
        'tables_count': len(all_tables),
        'parsed_data': all_parsed_data,
        'excel_path': excel_filename
    }

def create_excel_file(parsed_data, tables, raw_text, output_path):
    """Create Excel file with extracted data"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        sheets_created = 0

        # Parsed data sheet
        if parsed_data and len(parsed_data) > 0:
            try:
                # parsed_data is now a list of dictionaries (one per item)
                df_parsed = pd.DataFrame(parsed_data)
                df_parsed.to_excel(writer, sheet_name='Parsed_Data', index=False)
                sheets_created += 1
            except Exception as e:
                print(f"Error creating parsed data sheet: {e}")
                # If DataFrame creation fails, try with original data
                try:
                    if parsed_data:
                        df_parsed = pd.DataFrame(parsed_data)
                        df_parsed.to_excel(writer, sheet_name='Parsed_Data', index=False)
                        sheets_created += 1
                except Exception as e2:
                    print(f"Error creating parsed data sheet with original data: {e2}")

        # Tables sheet
        if tables:
            for i, table_info in enumerate(tables):
                try:
                    sheet_name = f'Table_{table_info["page"]}_{i+1}'
                    table_info['data'].to_excel(writer, sheet_name=sheet_name, index=False)
                    sheets_created += 1
                except Exception as e:
                    print(f"Error creating table sheet: {e}")

        # Raw text sheet
        if raw_text and raw_text.strip():
            try:
                df_text = pd.DataFrame({'Raw_Text': [raw_text]})
                df_text.to_excel(writer, sheet_name='Raw_Text', index=False)
                sheets_created += 1
            except Exception as e:
                print(f"Error creating raw text sheet: {e}")

        # If no sheets were created, create a default sheet
        if sheets_created == 0:
            try:
                df_empty = pd.DataFrame({'Message': ['No data could be extracted from this PDF']})
                df_empty.to_excel(writer, sheet_name='No_Data', index=False)
            except Exception as e:
                print(f"Error creating default sheet: {e}")

def create_consolidated_excel(all_data, output_path):
    """Create consolidated Excel for multiple files with all invoice data"""
    # Define the column order as requested
    columns = ['sno', 'invoice_no', 'date', 'receiver_name', 'receiver_gst', 'hsn',
               'quantity', 'taxable_amount', 'taxrate', 'cgst', 'sgst',
               'invoice_value', 'total_invoice_value']

    # Create DataFrame with specified columns
    if all_data:
        df = pd.DataFrame(all_data)

        # Ensure all required columns exist
        for col in columns:
            if col not in df.columns:
                df[col] = ''

        # Reorder columns as specified
        df = df[columns]
    else:
        # Create empty DataFrame with proper columns
        df = pd.DataFrame(columns=columns)

    # Rename columns to match user's requirements
    column_mapping = {
        'sno': 'SNO',
        'invoice_no': 'INVOICE NO',
        'date': 'DATE',
        'receiver_name': 'RECEIVER NAME',
        'receiver_gst': 'RECEIVER GST',
        'hsn': 'HSN',
        'quantity': 'QUANTITY',
        'taxable_amount': 'TAXABLE AMOUNT',
        # 'discount': 'DISCOUNT', # Removed
        'taxrate': 'TAXRATE',
        'cgst': 'CGST',
        'sgst': 'SGST',
        'invoice_value': 'INVOICE VALUE',
        'total_invoice_value': 'TOTAL INVOICE VALUE'
    }

    df = df.rename(columns=column_mapping)

    # Save to Excel
    df.to_excel(output_path, index=False)

    return output_path

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['PROCESSED_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash('File not found')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)