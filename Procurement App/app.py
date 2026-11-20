import os
import csv
import io
import re
import sqlite3
from datetime import datetime, date
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, send_file, Response
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_
from sqlalchemy.exc import IntegrityError
from werkzeug.utils import secure_filename
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd

# Optional OCR libraries for scanned PDFs
try:
    from pdf2image import convert_from_path
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

app = Flask(__name__)
app.config['SECRET_KEY'] = 'akiraz-procurement-secret-key-change-me'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///procurement.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = os.path.join('static', 'uploads')
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'xlsx', 'xls', 'csv'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
db = SQLAlchemy(app)

VALID_ENTITIES = {'requisitions', 'purchase_orders', 'invoices', 'batches', 'statements'}
INV_TOKEN_PATTERN = r'(?:INV\s*#|Invoice\s*No\.?|Invoice\s*#|Inv\s*No\.?|Inv\s*#)'


@app.context_processor
def inject_app_branding():
    return {
        'app_name': 'AkiraZ Procurement System',
        'app_tagline': 'Bubbly Pink, Blue & Purple Procurement Hub'
    }


def allowed_file(filename, allowed_set={'pdf'}):
    return bool(filename) and '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set


def parse_amount(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    cleaned = re.sub(r'[^\d.\-]', '', str(val).strip())
    try:
        return float(cleaned) if cleaned else 0.0
    except ValueError:
        return 0.0


def parse_int_or_none(val):
    try:
        return int(val) if val and str(val).strip() else None
    except ValueError:
        return None


def clean_excel_date(val):
    """Normalizes Excel Timestamps, datetime objects, MM/DD/YYYY, DD/MM/YYYY, or YYYY-MM-DD into YYYY-MM-DD."""
    if val is None or (isinstance(val, float) and pd.isna(val)) or pd.isna(val):
        return ''
    if isinstance(val, (datetime, date, pd.Timestamp)):
        return val.strftime('%Y-%m-%d')

    val_str = str(val).strip()
    if not val_str or val_str.lower() in ('nan', 'nat', 'none', 'pending', '-'):
        return ''

    date_part = val_str.split()[0]

    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%m-%d-%Y', '%d-%m-%Y', '%Y/%m/%d', '%m/%d/%y', '%d/%m/%y'):
        try:
            dt = datetime.strptime(date_part, fmt)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            continue
    return date_part


def format_date_us(val):
    if not val:
        return ''
    val_str = str(val).strip()
    try:
        dt = datetime.strptime(val_str[:10], '%Y-%m-%d')
        return dt.strftime('%m/%d/%Y')
    except ValueError:
        pass
    try:
        dt = datetime.strptime(val_str[:10], '%m/%d/%Y')
        return dt.strftime('%m/%d/%Y')
    except ValueError:
        return val_str


def save_uploaded_pdf(file_storage):
    if not file_storage or not file_storage.filename or not allowed_file(file_storage.filename, {'pdf'}):
        return None
    clean_name = secure_filename(file_storage.filename)
    if not clean_name:
        clean_name = "document.pdf"
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    unique_filename = f"{timestamp}_{clean_name}"
    file_storage.save(os.path.join(app.config['UPLOAD_FOLDER'], unique_filename))
    return unique_filename


def delete_pdf_safely(filename):
    if not filename:
        return
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
        except OSError as e:
            print(f"Warning: Could not remove file {file_path}: {e}")


# --- DATABASE MODELS ---

class Requisition(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    requisition_number = db.Column(db.String(50), unique=True, nullable=False)
    transaction_date = db.Column(db.String(20))
    quotation_number = db.Column(db.String(50))
    poid = db.Column(db.String(50))
    vendor_name = db.Column(db.String(100))
    description = db.Column(db.String(150))
    value = db.Column(db.Float, default=0.0)
    functional_area = db.Column(db.String(100))
    remarks = db.Column(db.Text)
    date_sent_to_bursary = db.Column(db.String(20))
    pdf_filename = db.Column(db.String(200))

    purchase_orders = db.relationship('PurchaseOrder', backref='requisition', lazy=True, cascade='all, delete-orphan')


class PurchaseOrder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    poid = db.Column(db.String(50), unique=True, nullable=False)
    requisition_id = db.Column(db.Integer, db.ForeignKey('requisition.id'), nullable=True)
    vendor_name = db.Column(db.String(100))
    po_amount = db.Column(db.Float, default=0.0)
    date = db.Column(db.String(20))
    service = db.Column(db.String(150))
    functional_area = db.Column(db.String(100))
    remarks = db.Column(db.Text)
    pdf_filename = db.Column(db.String(200))

    invoices = db.relationship('Invoice', backref='purchase_order', lazy=True, cascade='all, delete-orphan')


class BatchSummary(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    batch_number = db.Column(db.String(50), unique=True, nullable=False)
    functional_area = db.Column(db.String(100))
    period_collected = db.Column(db.String(50))
    date_sent_for_review = db.Column(db.String(20))
    date_reviewed = db.Column(db.String(20))
    date_sent_to_ta = db.Column(db.String(20))
    date_received_from_ta = db.Column(db.String(20))
    date_sent_to_om = db.Column(db.String(20))
    date_received_from_om = db.Column(db.String(20))
    submission_date = db.Column(db.String(20))
    remarks = db.Column(db.Text)

    invoices = db.relationship('Invoice', backref='batch', lazy=True)


class VendorStatement(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    statement_number = db.Column(db.String(50), unique=True, nullable=False)
    functional_area = db.Column(db.String(100))
    vendor = db.Column(db.String(100))
    statement_period = db.Column(db.String(50))
    date_received = db.Column(db.String(20))
    amount = db.Column(db.Float, default=0.0)
    remarks = db.Column(db.Text)
    pdf_filename = db.Column(db.String(200))

    invoices = db.relationship('Invoice', backref='statement', lazy=True)


class Invoice(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    invoice_number = db.Column(db.String(50), unique=True, nullable=False)
    invoice_date = db.Column(db.String(20))
    vendor_name = db.Column(db.String(100))
    invoice_amount = db.Column(db.Float, default=0.0)
    date_received = db.Column(db.String(20))
    invoice_details = db.Column(db.String(150))
    purchase_order_id = db.Column(db.Integer, db.ForeignKey('purchase_order.id'), nullable=True)
    vendor_statement_id = db.Column(db.Integer, db.ForeignKey('vendor_statement.id'), nullable=True)
    remarks = db.Column(db.Text)
    batch_summary_id = db.Column(db.Integer, db.ForeignKey('batch_summary.id'), nullable=True)
    date_sent_to_nekeisha = db.Column(db.String(20))
    date_returned_from_nekeisha = db.Column(db.String(20))
    date_sent_to_ta = db.Column(db.String(20))
    date_returned_from_ta = db.Column(db.String(20))
    date_sent_to_om = db.Column(db.String(20))
    date_returned_from_om = db.Column(db.String(20))
    date_submitted_to_bursary = db.Column(db.String(20))
    bursary_date_batch = db.Column(db.String(50))
    payment_received = db.Column(db.String(20))
    pdf_filename = db.Column(db.String(200))


def ensure_schema_updates():
    db_path = os.path.join('instance', 'procurement.db')
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            cursor.execute("PRAGMA table_info(requisition)")
            cols = [row[1] for row in cursor.fetchall()]
            if cols:
                if 'date' in cols and 'transaction_date' not in cols:
                    cursor.execute("ALTER TABLE requisition RENAME COLUMN date TO transaction_date")
                if 'quote_number' in cols and 'quotation_number' not in cols:
                    cursor.execute("ALTER TABLE requisition RENAME COLUMN quote_number TO quotation_number")
                if 'job' in cols and 'description' not in cols:
                    cursor.execute("ALTER TABLE requisition RENAME COLUMN job TO description")
                if 'amount' in cols and 'value' not in cols:
                    cursor.execute("ALTER TABLE requisition RENAME COLUMN amount TO value")
                if 'vendor' in cols and 'vendor_name' not in cols:
                    cursor.execute("ALTER TABLE requisition RENAME COLUMN vendor TO vendor_name")
                if 'poid' not in cols:
                    cursor.execute("ALTER TABLE requisition ADD COLUMN poid VARCHAR(50)")
                if 'remarks' not in cols:
                    cursor.execute("ALTER TABLE requisition ADD COLUMN remarks TEXT")

            cursor.execute("PRAGMA table_info(purchase_order)")
            po_cols = [row[1] for row in cursor.fetchall()]
            if po_cols:
                if 'po_number' in po_cols and 'poid' not in po_cols:
                    cursor.execute("ALTER TABLE purchase_order RENAME COLUMN po_number TO poid")
                if 'vendor' in po_cols and 'vendor_name' not in po_cols:
                    cursor.execute("ALTER TABLE purchase_order RENAME COLUMN vendor TO vendor_name")
                if 'amount' in po_cols and 'po_amount' not in po_cols:
                    cursor.execute("ALTER TABLE purchase_order RENAME COLUMN amount TO po_amount")
                if 'job' in po_cols and 'service' not in po_cols:
                    cursor.execute("ALTER TABLE purchase_order RENAME COLUMN job TO service")
                if 'remarks' not in po_cols:
                    cursor.execute("ALTER TABLE purchase_order ADD COLUMN remarks TEXT")

            cursor.execute("PRAGMA table_info(invoice)")
            inv_cols = [row[1] for row in cursor.fetchall()]
            if inv_cols:
                if 'date' in inv_cols and 'invoice_date' not in inv_cols:
                    cursor.execute("ALTER TABLE invoice RENAME COLUMN date TO invoice_date")
                if 'vendor' in inv_cols and 'vendor_name' not in inv_cols:
                    cursor.execute("ALTER TABLE invoice RENAME COLUMN vendor TO vendor_name")
                if 'amount' in inv_cols and 'invoice_amount' not in inv_cols:
                    cursor.execute("ALTER TABLE invoice RENAME COLUMN amount TO invoice_amount")
                if 'job' in inv_cols and 'invoice_details' not in inv_cols:
                    cursor.execute("ALTER TABLE invoice RENAME COLUMN job TO invoice_details")
                if 'date_sent_to_bursary' in inv_cols and 'date_submitted_to_bursary' not in inv_cols:
                    cursor.execute("ALTER TABLE invoice RENAME COLUMN date_sent_to_bursary TO date_submitted_to_bursary")
                if 'date_received_from_ta' in inv_cols and 'date_returned_from_ta' not in inv_cols:
                    cursor.execute("ALTER TABLE invoice RENAME COLUMN date_received_from_ta TO date_returned_from_ta")
                if 'date_received_from_om' in inv_cols and 'date_returned_from_om' not in inv_cols:
                    cursor.execute("ALTER TABLE invoice RENAME COLUMN date_received_from_om TO date_returned_from_om")

                for new_col in [
                    'date_received VARCHAR(20)',
                    'remarks TEXT',
                    'date_sent_to_nekeisha VARCHAR(20)',
                    'date_returned_from_nekeisha VARCHAR(20)',
                    'bursary_date_batch VARCHAR(50)',
                    'payment_received VARCHAR(20)',
                    'vendor_statement_id INTEGER REFERENCES vendor_statement(id)'
                ]:
                    col_name = new_col.split()[0]
                    if col_name not in inv_cols:
                        cursor.execute(f"ALTER TABLE invoice ADD COLUMN {new_col}")

            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Schema check warning: {e}")


def sync_requisition_po_links():
    linked_count = 0
    reqs_with_poid = Requisition.query.filter(
        Requisition.poid != None,
        Requisition.poid != ''
    ).all()
    for req in reqs_with_poid:
        po = PurchaseOrder.query.filter(
            PurchaseOrder.poid.ilike(req.poid.strip())
        ).first()
        if po and po.requisition_id != req.id:
            po.requisition_id = req.id
            linked_count += 1

    linked_pos = PurchaseOrder.query.filter(PurchaseOrder.requisition_id != None).all()
    for po in linked_pos:
        if po.requisition and not (po.requisition.poid or '').strip():
            po.requisition.poid = po.poid
            linked_count += 1

    if linked_count:
        db.session.commit()
    return linked_count


# --- PDF EXTRACTION & STATEMENT RECONCILIATION HELPERS ---

def validate_iso_date(year_str, month_str, day_str):
    try:
        dt = datetime(int(year_str), int(month_str), int(day_str))
        return dt.strftime('%Y-%m-%d')
    except ValueError:
        return None


def parse_flexible_date(text):
    d_iso = re.search(r'\b(20\d{2})-(\d{2})-(\d{2})\b', text)
    if d_iso:
        valid = validate_iso_date(*d_iso.groups())
        if valid:
            return valid

    d_slash = re.search(r'\b(\d{1,2})[\/\-](\d{1,2})[\/\-](20\d{2})\b', text)
    if d_slash:
        part1, part2, year = d_slash.groups()
        if int(part2) > 12:
            month, day = part1, part2
        else:
            day, month = part1, part2
        valid = validate_iso_date(year, month, day)
        if valid:
            return valid

    months = {
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
    }
    d_dmy = re.search(r'\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\.,]?\s+(20\d{2})\b', text, re.IGNORECASE)
    if d_dmy:
        day, mon_str, year = d_dmy.groups()
        valid = validate_iso_date(year, months[mon_str.lower()[:3]], day)
        if valid:
            return valid

    d_mdy = re.search(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\.,]?\s+(\d{1,2})[\.,]?\s+(20\d{2})\b', text, re.IGNORECASE)
    if d_mdy:
        mon_str, day, year = d_mdy.groups()
        valid = validate_iso_date(year, months[mon_str.lower()[:3]], day)
        if valid:
            return valid

    return None


def read_pdf_text_with_ocr_fallback(filepath, max_pages=5):
    """Reads text from a PDF using pypdf and falls back to OCR when text is sparse/missing."""
    full_text = ""
    extraction_method = "Digital Text"

    if not filepath or not os.path.exists(filepath):
        return "", "File Not Found"

    try:
        reader = PdfReader(filepath)
        for page in reader.pages[:max_pages]:
            full_text += (page.extract_text() or "") + "\n"
    except Exception as e:
        print(f"Error reading digital PDF ({filepath}): {e}")

    if len(full_text.strip()) < 25 and OCR_AVAILABLE:
        try:
            images = convert_from_path(filepath, first_page=1, last_page=max_pages)
            ocr_pages = [pytesseract.image_to_string(img) for img in images]
            full_text = "\n".join(ocr_pages)
            extraction_method = "Scanned OCR"
        except Exception as e:
            print(f"OCR fallback skipped/failed ({filepath}): {e}")

    return full_text, extraction_method


def extract_pdf_data(filepath, entity_type):
    full_text, extraction_method = read_pdf_text_with_ocr_fallback(filepath, max_pages=3)

    extracted = {
        'raw_extracted_text': full_text.strip(),
        'extraction_method': extraction_method
    }

    if not full_text.strip():
        return extracted

    if entity_type == 'requisitions':
        m = re.search(r'(?:Requisition|Req)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z0-9\-\/]+)', full_text, re.IGNORECASE)
        if m:
            extracted['requisition_number'] = m.group(1).strip()
        q = re.search(r'(?:Quote|Quotation)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z0-9\-\/]+)', full_text, re.IGNORECASE)
        if q:
            extracted['quotation_number'] = q.group(1).strip()

    elif entity_type == 'purchase_orders':
        m = re.search(r'(?:Purchase\s*Order|P\.?O\.?)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z0-9\-\/]+)', full_text, re.IGNORECASE)
        if m:
            extracted['poid'] = m.group(1).strip()
        r_match = re.search(r'(?:Requisition|Req)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z0-9\-\/]+)', full_text, re.IGNORECASE)
        if r_match:
            req_obj = Requisition.query.filter(Requisition.requisition_number.ilike(f"%{r_match.group(1).strip()}%")).first()
            if req_obj:
                extracted['requisition_id'] = req_obj.id
                if req_obj.functional_area:
                    extracted['functional_area'] = req_obj.functional_area
                if req_obj.vendor_name:
                    extracted['vendor_name'] = req_obj.vendor_name

    elif entity_type == 'invoices':
        m = re.search(
            rf'{INV_TOKEN_PATTERN}\s*[:\-]?\s*([A-Za-z0-9\-_\/]+)',
            full_text,
            re.IGNORECASE
        )
        if m:
            extracted['invoice_number'] = m.group(1).strip()
        po_match = re.search(r'(?:Purchase\s*Order|P\.?O\.?)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z0-9\-\/]+)', full_text, re.IGNORECASE)
        if po_match:
            po_obj = PurchaseOrder.query.filter(PurchaseOrder.poid.ilike(f"%{po_match.group(1).strip()}%")).first()
            if po_obj:
                extracted['purchase_order_id'] = po_obj.id
                if po_obj.vendor_name:
                    extracted['vendor_name'] = po_obj.vendor_name

    elif entity_type == 'statements':
        m = re.search(r'(?:Statement|Stmt)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z0-9\-\/]+)', full_text, re.IGNORECASE)
        if m:
            extracted['statement_number'] = m.group(1).strip()

    v = re.search(r'(?:Vendor|Supplier|From|Company|Payable\s*To)\s*[:\-]?\s*([A-Za-z0-9\s\.,&\-\']+)', full_text, re.IGNORECASE)
    if v:
        v_clean = v.group(1).splitlines()[0].strip()[:100]
        extracted['vendor_name'] = v_clean
        extracted['vendor'] = v_clean
    else:
        for line in full_text.splitlines():
            clean_line = line.strip()
            if len(clean_line) > 3 and not re.match(r'^(invoice|quotation|quote|statement|purchase order|tax invoice|page\s+\d+)', clean_line, re.IGNORECASE):
                extracted['vendor_name'] = clean_line[:100]
                extracted['vendor'] = clean_line[:100]
                break

    j = re.search(r'(?:Job|Project|Description|Reference|Subject|Re|Service)\s*[:\-]?\s*([A-Za-z0-9\s\.,&\-\(\)\/]+)', full_text, re.IGNORECASE)
    if j:
        desc = j.group(1).splitlines()[0].strip()[:150]
        extracted['description'] = desc
        extracted['invoice_details'] = desc
        extracted['service'] = desc

    parsed_date = parse_flexible_date(full_text)
    if parsed_date:
        extracted['transaction_date'] = parsed_date
        extracted['date'] = parsed_date
        extracted['invoice_date'] = parsed_date
        extracted['date_received'] = parsed_date

    amt_labeled = re.search(r'(?:Grand\s*Total|Total\s*Due|Balance\s*Due|Total\s*Amount|Total|Amount\s*Due)\s*[:\-]?\s*(?:BBD)?\$?\s*([\d,]+\.\d{2})', full_text, re.IGNORECASE)
    if amt_labeled:
        val = parse_amount(amt_labeled.group(1))
        extracted['value'] = val
        extracted['po_amount'] = val
        extracted['invoice_amount'] = val
        extracted['amount'] = val
    else:
        all_amounts = re.findall(r'(?:BBD)?\$\s*([\d,]+\.\d{2})', full_text, re.IGNORECASE)
        if all_amounts:
            val = max(parse_amount(a) for a in all_amounts)
            extracted['value'] = val
            extracted['po_amount'] = val
            extracted['invoice_amount'] = val
            extracted['amount'] = val

    return extracted


def extract_statement_invoices_from_pdf(filepath):
    """Scans a Vendor Statement PDF for invoices labeled after INV # / Invoice No. or in line items."""
    full_text, method = read_pdf_text_with_ocr_fallback(filepath, max_pages=5)
    if not full_text.strip():
        print(f"Warning: No readable text extracted from statement PDF {filepath} ({method}).")
        return []

    extracted_invoices = []
    seen = set()

    labeled_pattern = re.compile(
        rf'{INV_TOKEN_PATTERN}\s*[:\-]?\s*([A-Za-z0-9\-_\/]+)',
        re.IGNORECASE
    )
    for m in labeled_pattern.finditer(full_text):
        inv_num = m.group(1).strip()
        inv_key = inv_num.upper()
        if inv_key and inv_key not in seen:
            seen.add(inv_key)
            window = full_text[m.start():m.start() + 180]
            date_match = re.search(
                r'(\d{1,2}[\/\-]\d{1,2}[\/\-](?:20\d{2}|\d{2})|20\d{2}-\d{2}-\d{2})',
                window
            )
            amt_match = re.search(
                r'(?:BBD|USD|\$)?\s*([\d]{1,3}(?:,\d{3})*\.\d{2})',
                window,
                re.IGNORECASE
            )
            extracted_invoices.append({
                'invoice_number': inv_num,
                'invoice_date': clean_excel_date(date_match.group(1)) if date_match else '',
                'invoice_amount': parse_amount(amt_match.group(1)) if amt_match else 0.0
            })

    line_pattern = re.compile(
        r'\b([A-Z]{1,6}[\-]?\d{2,10}(?:[\-\/]\d{1,6})?|\d{4,10})\b.*?'
        r'(\d{1,2}[\/\-]\d{1,2}[\/\-](?:20\d{2}|\d{2})|20\d{2}-\d{2}-\d{2}).*?'
        r'(?:BBD|USD|\$)?\s*([\d]{1,3}(?:,\d{3})*\.\d{2})',
        re.IGNORECASE
    )
    for line in full_text.splitlines():
        clean_line = re.sub(r'\s+', ' ', line).strip()
        if not clean_line:
            continue
        match = line_pattern.search(clean_line)
        if match:
            inv_num, date_str, amt_str = match.groups()
            inv_key = inv_num.strip().upper()
            if inv_key not in seen:
                seen.add(inv_key)
                extracted_invoices.append({
                    'invoice_number': inv_num.strip(),
                    'invoice_date': clean_excel_date(date_str),
                    'invoice_amount': parse_amount(amt_str)
                })

    return extracted_invoices


def reconcile_statement_with_invoices_section(stmt, all_invoices):
    """Links matching Invoices section records by invoice numbers found in the statement PDF
    and returns linked Invoice objects, PDF-extracted invoices, and known invoice numbers."""
    pdf_invoices = []
    if stmt.pdf_filename:
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], stmt.pdf_filename)
        pdf_invoices = extract_statement_invoices_from_pdf(pdf_path)

    pdf_num_map = {
        (p.get('invoice_number') or '').strip().upper(): p
        for p in pdf_invoices if (p.get('invoice_number') or '').strip()
    }

    matched_invoices = list(stmt.invoices)
    seen_ids = {inv.id for inv in matched_invoices}
    linked_any = False

    if pdf_num_map:
        for inv in all_invoices:
            inv_key = (inv.invoice_number or '').strip().upper()
            if inv_key in pdf_num_map and inv.id not in seen_ids:
                if inv.vendor_statement_id is None:
                    inv.vendor_statement_id = stmt.id
                    linked_any = True
                matched_invoices.append(inv)
                seen_ids.add(inv.id)

    if linked_any:
        db.session.commit()

    known_inv_keys = {
        (inv.invoice_number or '').strip().upper()
        for inv in matched_invoices if inv.invoice_number
    }

    return matched_invoices, pdf_invoices, known_inv_keys


def resolve_po_number_for_invoice(inv=None, vendor_name='', pdf_context_text='', all_pos=None):
    """Resolves a PurchaseOrder.poid string for an Invoice record or PDF-extracted invoice."""
    all_pos = all_pos if all_pos is not None else PurchaseOrder.query.all()
    po_by_upper = {
        (po.poid or '').strip().upper(): (po.poid or '').strip()
        for po in all_pos if (po.poid or '').strip()
    }

    if inv is not None and inv.purchase_order and (inv.purchase_order.poid or '').strip():
        return inv.purchase_order.poid.strip()

    combined_texts = []
    if inv is not None:
        for field_val in (inv.invoice_details, inv.remarks):
            if field_val and str(field_val).strip():
                combined_texts.append(str(field_val).strip())
    if pdf_context_text:
        combined_texts.append(pdf_context_text)

    blob = "\n".join(combined_texts)
    if blob:
        blob_upper = blob.upper()
        for po_key, po_val in po_by_upper.items():
            if re.search(rf'(?<![A-Z0-9]){re.escape(po_key)}(?![A-Z0-9])', blob_upper):
                return po_val

        po_match = re.search(
            r'(?:Purchase\s*Order|P\.?O\.?ID|P\.?O\.?)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Za-z0-9\-_\/]+)',
            blob,
            re.IGNORECASE
        )
        if po_match:
            candidate = po_match.group(1).strip()
            if candidate.upper() in po_by_upper:
                return po_by_upper[candidate.upper()]
            return candidate

    clean_vendor = (vendor_name or '').strip().lower()
    if clean_vendor:
        vendor_pos = [
            po for po in all_pos
            if (po.vendor_name or '').strip().lower()
            and (clean_vendor in (po.vendor_name or '').strip().lower()
                 or (po.vendor_name or '').strip().lower() in clean_vendor)
        ]
        if len(vendor_pos) == 1 and (vendor_pos[0].poid or '').strip():
            return vendor_pos[0].poid.strip()
        if vendor_pos and inv is not None and (inv.invoice_date or '').strip():
            inv_dt_str = (inv.invoice_date or '').strip()[:10]
            for po in vendor_pos:
                if (po.date or '').strip()[:10] == inv_dt_str and (po.poid or '').strip():
                    return po.poid.strip()
        if len(vendor_pos) > 1 and (vendor_pos[0].poid or '').strip():
            return vendor_pos[0].poid.strip()

    return ''


# --- REPORT HELPER FUNCTIONS ---

def get_target_15th(ref_date=None):
    if ref_date:
        try:
            dt = datetime.strptime(ref_date, '%Y-%m-%d').date()
            return dt.replace(day=15)
        except ValueError:
            pass
    today = date.today()
    return today.replace(day=15)


def get_statement_eligible_invoices(statement_id, target_15th):
    invoices = Invoice.query.filter_by(vendor_statement_id=statement_id).all()
    qualifying = []
    for inv in invoices:
        sub_str = inv.date_submitted_to_bursary
        if not sub_str:
            continue
        try:
            sub_date = datetime.strptime(sub_str.strip()[:10], '%Y-%m-%d').date()
            days_before_15th = (target_15th - sub_date).days
            if days_before_15th >= 28:
                qualifying.append({
                    'invoice': inv,
                    'submission_date': sub_str,
                    'days_before_15th': days_before_15th
                })
        except ValueError:
            continue
    return qualifying


# --- ROUTES ---

@app.route('/')
def index():
    current_year_str = str(date.today().year)
    selected_year = request.args.get('year', current_year_str).strip()

    years_set = {current_year_str}
    for r in Requisition.query.all():
        if r.date_sent_to_bursary and len(str(r.date_sent_to_bursary)) >= 4:
            years_set.add(str(r.date_sent_to_bursary)[:4])
    for po in PurchaseOrder.query.all():
        if po.date and len(str(po.date)) >= 4:
            years_set.add(str(po.date)[:4])
    for inv in Invoice.query.all():
        if inv.invoice_date and len(str(inv.invoice_date)) >= 4:
            years_set.add(str(inv.invoice_date)[:4])
    available_years = sorted(list(years_set), reverse=True)

    stats_requisitions = Requisition.query.filter(
        Requisition.date_sent_to_bursary.like(f"{selected_year}%")
    ).count()

    reqs_without_po = Requisition.query.outerjoin(PurchaseOrder).filter(
        PurchaseOrder.id == None,
        or_(Requisition.poid == None, Requisition.poid == ''),
        Requisition.date_sent_to_bursary.like(f"{selected_year}%")
    ).count()

    stats_pos = PurchaseOrder.query.filter(
        PurchaseOrder.date.like(f"{selected_year}%")
    ).count()

    total_committed = db.session.query(db.func.sum(PurchaseOrder.po_amount)).filter(
        PurchaseOrder.date.like(f"{selected_year}%")
    ).scalar() or 0.0

    invoices_query = Invoice.query.filter(
        Invoice.invoice_date.like(f"{selected_year}%")
    )
    total_invoices_count = invoices_query.count()

    submitted_invoices = invoices_query.filter(
        Invoice.date_submitted_to_bursary != None,
        Invoice.date_submitted_to_bursary != ''
    ).count()

    unsubmitted_invoices = invoices_query.filter(
        or_(Invoice.date_submitted_to_bursary == None, Invoice.date_submitted_to_bursary == '')
    ).count()

    total_invoiced_sum = db.session.query(db.func.sum(Invoice.invoice_amount)).filter(
        Invoice.invoice_date.like(f"{selected_year}%")
    ).scalar() or 0.0

    submitted_amount_total = db.session.query(db.func.sum(Invoice.invoice_amount)).filter(
        Invoice.invoice_date.like(f"{selected_year}%"),
        Invoice.date_submitted_to_bursary != None,
        Invoice.date_submitted_to_bursary != ''
    ).scalar() or 0.0

    unsubmitted_amount_total = total_invoiced_sum - submitted_amount_total

    all_statements = VendorStatement.query.all()
    year_statements = []
    for stmt in all_statements:
        rec_str = (stmt.date_received or '').strip()
        period_str = (stmt.statement_period or '').strip()
        matches_year = (
            rec_str.startswith(selected_year)
            or rec_str.endswith(f"/{selected_year}")
            or (selected_year in period_str)
        )
        if matches_year:
            year_statements.append(stmt)

    statements_count = len(year_statements)
    statements_amount_total = sum((s.amount or 0.0) for s in year_statements)

    turnaround_days = []
    for inv in invoices_query.all():
        if inv.date_submitted_to_bursary and str(inv.date_submitted_to_bursary).strip() and inv.invoice_date:
            start_date_str = inv.date_received or inv.invoice_date
            try:
                d1 = datetime.strptime(str(start_date_str).strip()[:10], '%Y-%m-%d').date()
                d2 = datetime.strptime(str(inv.date_submitted_to_bursary).strip()[:10], '%Y-%m-%d').date()
                if d2 >= d1:
                    turnaround_days.append((d2 - d1).days)
            except ValueError:
                pass

    avg_turnaround = round(sum(turnaround_days) / len(turnaround_days), 1) if turnaround_days else 0

    chart_data = {
        'total_committed': total_committed,
        'total_invoiced': total_invoiced_sum,
        'submitted_amount_total': submitted_amount_total,
        'unsubmitted_amount_total': unsubmitted_amount_total,
        'statements_amount_total': statements_amount_total,
        'avg_turnaround': avg_turnaround
    }

    stats = {
        'requisitions': stats_requisitions,
        'pos': stats_pos,
        'invoices': total_invoices_count,
        'submitted_invoices': submitted_invoices,
        'unsubmitted_invoices': unsubmitted_invoices,
        'batches': BatchSummary.query.filter(BatchSummary.submission_date.like(f"{selected_year}%")).count(),
        'statements': statements_count,
        'statements_amount_total': statements_amount_total,
        'reqs_without_po': reqs_without_po,
    }

    return render_template(
        'index.html',
        stats=stats,
        chart_data=chart_data,
        available_years=available_years,
        selected_year=selected_year
    )


@app.route('/<entity_type>')
def list_entities(entity_type):
    if entity_type not in VALID_ENTITIES:
        flash("Invalid section requested.", "danger")
        return redirect(url_for('index'))

    batches = BatchSummary.query.all() if entity_type == 'invoices' else []
    statements = VendorStatement.query.all() if entity_type == 'invoices' else []

    if entity_type == 'requisitions':
        items = Requisition.query.all()
    elif entity_type == 'purchase_orders':
        items = PurchaseOrder.query.all()
    elif entity_type == 'invoices':
        items = Invoice.query.all()
    elif entity_type == 'batches':
        items = BatchSummary.query.all()
    elif entity_type == 'statements':
        items = VendorStatement.query.all()

    return render_template('list_view.html', entity_type=entity_type, items=items,
                           batches=batches, statements=statements, today_str=date.today().strftime('%Y-%m-%d'),
                           format_date_us=format_date_us)


@app.route('/import/spreadsheet/<entity_type>', methods=['POST'])
def import_spreadsheet(entity_type):
    if entity_type not in VALID_ENTITIES:
        flash("Invalid entity type for import.", "danger")
        return redirect(url_for('index'))

    file = request.files.get('spreadsheet_file')
    if not file or not file.filename:
        flash("Please select an Excel (.xlsx, .xls) or CSV file to import.", "warning")
        return redirect(url_for('list_entities', entity_type=entity_type))

    try:
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
    except Exception as e:
        flash(f"Error reading file: {e}", "danger")
        return redirect(url_for('list_entities', entity_type=entity_type))

    def norm_key(k):
        return re.sub(r'\s+', ' ', str(k)).strip().lower()

    col_lookup = {norm_key(col): col for col in df.columns}

    imported_count = 0
    updated_count = 0
    skipped_count = 0

    for _, row in df.iterrows():
        def raw_val(aliases):
            if isinstance(aliases, str):
                aliases = [aliases]
            for alias in aliases:
                actual_col = col_lookup.get(norm_key(alias))
                if actual_col is not None:
                    v = row[actual_col]
                    if not pd.isna(v):
                        return v
            return None

        def str_val(aliases, default=""):
            v = raw_val(aliases)
            if v is None:
                return default
            s = str(v).strip()
            if s.endswith('.0') and s[:-2].isdigit():
                s = s[:-2]
            return "" if s.lower() == 'nan' else s

        def date_val(aliases):
            v = raw_val(aliases)
            return clean_excel_date(v) or None

        def amt_val(aliases):
            v = raw_val(aliases)
            return parse_amount(v)

        if entity_type == 'requisitions':
            req_num = str_val(['Requisition Number', 'Req Number', 'Req #', 'requisition_number', 'Title'])
            if not req_num:
                skipped_count += 1
                continue

            existing = Requisition.query.filter_by(requisition_number=req_num).first()
            target = existing if existing else Requisition(requisition_number=req_num)

            t_date = date_val(['Transaction Date', 'Date', 'transaction_date'])
            q_num = str_val(['Quotation Number', 'Quote Number', 'Quote #', 'quotation_number', 'quote_number'])
            poid_str = str_val(['POID', 'PO ID', 'PO Number', 'PO #', 'poid'])
            v_name = str_val(['Vendor Name', 'Vendor', 'Supplier', 'vendor_name', 'vendor'])
            desc = str_val(['Description', 'Job', 'Details', 'description'])
            val_num = amt_val(['Value', 'Amount', 'Total', 'value', 'amount'])
            f_area = str_val(['Functional Area', 'functional_area'])
            rem = str_val(['Remarks', 'Notes', 'remarks'])
            bursary_dt = date_val(['Date Sent to Bursary', 'Sent to Bursary', 'date_sent_to_bursary'])

            if t_date: target.transaction_date = t_date
            if q_num: target.quotation_number = q_num
            if poid_str: target.poid = poid_str
            if v_name: target.vendor_name = v_name
            if desc: target.description = desc
            if val_num: target.value = val_num
            if f_area: target.functional_area = f_area
            if rem: target.remarks = rem
            if bursary_dt: target.date_sent_to_bursary = bursary_dt

            if existing:
                updated_count += 1
            else:
                db.session.add(target)
                imported_count += 1

        elif entity_type == 'purchase_orders':
            po_num = str_val(['POID', 'PO ID', 'PO Number', 'PO #', 'poid', 'Title'])
            if not po_num:
                skipped_count += 1
                continue

            existing = PurchaseOrder.query.filter_by(poid=po_num).first()
            target = existing if existing else PurchaseOrder(poid=po_num)

            req_str = str_val(['Requisition Number', 'Req Number', 'Req #', 'requisition_number', 'requisition_id'])
            if req_str:
                req_obj = Requisition.query.filter(Requisition.requisition_number.ilike(req_str)).first()
                if req_obj:
                    target.requisition_id = req_obj.id
            else:
                req_by_poid = Requisition.query.filter(Requisition.poid.ilike(po_num)).first()
                if req_by_poid:
                    target.requisition_id = req_by_poid.id

            v_name = str_val(['Vendor Name', 'Vendor', 'Supplier', 'vendor_name', 'vendor'])
            po_amt = amt_val(['PO Amount', 'Amount', 'Value', 'po_amount', 'amount'])
            po_dt = date_val(['Date', 'PO Date', 'Transaction Date', 'date'])
            srv = str_val(['Service', 'Description', 'Job', 'service'])
            f_area = str_val(['Functional Area', 'functional_area'])
            rem = str_val(['Remarks', 'Notes', 'remarks'])

            if v_name: target.vendor_name = v_name
            if po_amt: target.po_amount = po_amt
            if po_dt: target.date = po_dt
            if srv: target.service = srv
            if f_area: target.functional_area = f_area
            if rem: target.remarks = rem

            if existing:
                updated_count += 1
            else:
                db.session.add(target)
                imported_count += 1

        elif entity_type == 'invoices':
            inv_num = str_val(['Invoice Number', 'Invoice #', 'Inv #', 'invoice_number', 'Title'])
            if not inv_num:
                skipped_count += 1
                continue

            existing = Invoice.query.filter_by(invoice_number=inv_num).first()
            target = existing if existing else Invoice(invoice_number=inv_num)

            inv_dt = date_val(['Invoice Date', 'Date', 'invoice_date'])
            v_name = str_val(['Vendor Name', 'Vendor', 'Supplier', 'vendor_name', 'vendor'])
            inv_amt = amt_val(['Invoice Amount', 'Amount', 'Total', 'invoice_amount', 'amount'])
            dt_rec = date_val(['Date Received', 'Received Date', 'date_received'])
            details = str_val(['Invoice Details', 'Description', 'Details', 'Job', 'invoice_details'])

            poid_str = str_val(['POID', 'PO ID', 'PO Number', 'PO #', 'poid'])
            if poid_str:
                po_obj = PurchaseOrder.query.filter(PurchaseOrder.poid.ilike(poid_str)).first()
                if po_obj:
                    target.purchase_order_id = po_obj.id

            stmt_str = str_val(['Statement Number', 'Statement #', 'Vendor Statement', 'statement_number'])
            if stmt_str:
                stmt_obj = VendorStatement.query.filter(VendorStatement.statement_number.ilike(stmt_str)).first()
                if stmt_obj:
                    target.vendor_statement_id = stmt_obj.id

            rem = str_val(['Remarks', 'Notes', 'remarks'])

            batch_str = str_val(['Batch ID (Lookup)', 'Batch ID', 'Batch Number', 'Batch #', 'batch_number'])
            if batch_str:
                batch_obj = BatchSummary.query.filter(BatchSummary.batch_number.ilike(batch_str)).first()
                if not batch_obj:
                    batch_obj = BatchSummary(batch_number=batch_str)
                    db.session.add(batch_obj)
                    db.session.flush()
                target.batch_summary_id = batch_obj.id

            dt_sent_nek = date_val(['Date Sent to Nekeisha', 'Sent to Nekeisha', 'date_sent_to_nekeisha'])
            dt_ret_nek = date_val(['Date Returned from Nekeisha', 'Returned from Nekeisha', 'date_returned_from_nekeisha'])
            dt_sent_ta = date_val(['Date Sent to TA', 'Sent to TA', 'date_sent_to_ta'])
            dt_ret_ta = date_val(['Date Returned from TA', 'Date Received from TA', 'Returned from TA', 'date_returned_from_ta'])
            dt_sent_om = date_val(['Date Sent to OM', 'Sent to OM', 'date_sent_to_om'])
            dt_ret_om = date_val(['Date Returned from OM', 'Date Received from OM', 'Returned from OM', 'date_returned_from_om'])
            dt_sub_bur = date_val(['Date Submitted to Bursary', 'Date Sent to Bursary', 'Submitted to Bursary', 'date_submitted_to_bursary'])
            bur_batch = str_val(['Bursary Date Batch', 'Bursary Batch', 'bursary_date_batch'])
            pay_rec = date_val(['Payment Received', 'Date Payment Received', 'payment_received'])

            if inv_dt: target.invoice_date = inv_dt
            if v_name: target.vendor_name = v_name
            if inv_amt: target.invoice_amount = inv_amt
            if dt_rec: target.date_received = dt_rec
            if details: target.invoice_details = details
            if rem: target.remarks = rem
            if dt_sent_nek: target.date_sent_to_nekeisha = dt_sent_nek
            if dt_ret_nek: target.date_returned_from_nekeisha = dt_ret_nek
            if dt_sent_ta: target.date_sent_to_ta = dt_sent_ta
            if dt_ret_ta: target.date_returned_from_ta = dt_ret_ta
            if dt_sent_om: target.date_sent_to_om = dt_sent_om
            if dt_ret_om: target.date_returned_from_om = dt_ret_om
            if dt_sub_bur: target.date_submitted_to_bursary = dt_sub_bur
            if bur_batch: target.bursary_date_batch = bur_batch
            if pay_rec: target.payment_received = pay_rec

            if existing:
                updated_count += 1
            else:
                db.session.add(target)
                imported_count += 1

        elif entity_type == 'statements':
            stmt_num = str_val(['Statement Number', 'Statement #', 'Statement', 'statement_number', 'Title'])
            if not stmt_num:
                skipped_count += 1
                continue

            existing = VendorStatement.query.filter_by(statement_number=stmt_num).first()
            target = existing if existing else VendorStatement(statement_number=stmt_num)

            f_area = str_val(['Functional Area', 'functional_area'])
            vendor = str_val(['Vendor', 'Vendor Name', 'Supplier', 'vendor'])
            period = str_val(['Statement Period', 'Period', 'statement_period'])
            dt_rec = date_val(['Date Received', 'Received Date', 'date_received'])
            amount = amt_val(['Amount', 'Balance', 'Total', 'amount'])
            remarks = str_val(['Remarks', 'Notes', 'remarks'])

            if f_area: target.functional_area = f_area
            if vendor: target.vendor = vendor
            if period: target.statement_period = period
            if dt_rec: target.date_received = dt_rec
            if amount: target.amount = amount
            if remarks: target.remarks = remarks

            if existing:
                updated_count += 1
            else:
                db.session.add(target)
                imported_count += 1

    db.session.commit()
    sync_requisition_po_links()
    flash(f"Import complete! Added {imported_count} new record(s), updated {updated_count} existing record(s) ({skipped_count} blank rows skipped).", "success")
    return redirect(url_for('list_entities', entity_type=entity_type))


@app.route('/<entity_type>/bulk_action', methods=['POST'])
def bulk_action(entity_type):
    if entity_type not in VALID_ENTITIES:
        flash("Invalid section requested.", "danger")
        return redirect(url_for('index'))

    selected_ids = [parse_int_or_none(x) for x in request.form.getlist('selected_ids')]
    selected_ids = [x for x in selected_ids if x is not None]
    action = request.form.get('bulk_action')

    if not selected_ids:
        flash("No records were selected for bulk action.", "warning")
        return redirect(url_for('list_entities', entity_type=entity_type))

    model_map = {
        'requisitions': Requisition,
        'purchase_orders': PurchaseOrder,
        'invoices': Invoice,
        'batches': BatchSummary,
        'statements': VendorStatement
    }
    Model = model_map[entity_type]
    records = Model.query.filter(Model.id.in_(selected_ids)).all()

    if action == 'delete':
        for rec in records:
            if hasattr(rec, 'pdf_filename') and rec.pdf_filename:
                delete_pdf_safely(rec.pdf_filename)
            db.session.delete(rec)
        db.session.commit()
        flash(f"Successfully deleted {len(records)} record(s).", "success")

    elif entity_type == 'invoices':
        action_date = (request.form.get('bulk_date') or date.today().strftime('%Y-%m-%d')).strip()
        if action == 'assign_batch':
            batch_id = parse_int_or_none(request.form.get('bulk_batch_id'))
            if not batch_id:
                flash("Please select a Batch Summary before clicking Apply.", "warning")
                return redirect(url_for('list_entities', entity_type='invoices'))
            for inv in records:
                inv.batch_summary_id = batch_id
            db.session.commit()
            flash(f"Assigned {len(records)} invoice(s) to selected batch.", "success")
        elif action == 'assign_statement':
            stmt_id = parse_int_or_none(request.form.get('bulk_statement_id'))
            if not stmt_id:
                flash("Please select a Vendor Statement before clicking Apply.", "warning")
                return redirect(url_for('list_entities', entity_type='invoices'))
            for inv in records:
                inv.vendor_statement_id = stmt_id
            db.session.commit()
            flash(f"Assigned {len(records)} invoice(s) to selected vendor statement.", "success")
        elif action == 'mark_bursary':
            for inv in records:
                inv.date_submitted_to_bursary = action_date
            db.session.commit()
            flash(f"Marked {len(records)} invoice(s) as Submitted to Bursary on {action_date}.", "success")
        elif action == 'mark_ta':
            for inv in records:
                inv.date_sent_to_ta = action_date
            db.session.commit()
            flash(f"Marked {len(records)} invoice(s) as Sent to TA on {action_date}.", "success")
        elif action == 'mark_om':
            for inv in records:
                inv.date_sent_to_om = action_date
            db.session.commit()
            flash(f"Marked {len(records)} invoice(s) as Sent to OM on {action_date}.", "success")

    return redirect(url_for('list_entities', entity_type=entity_type))


@app.route('/<entity_type>/add', methods=['GET', 'POST'])
def add_entity(entity_type):
    if entity_type not in VALID_ENTITIES:
        flash("Invalid section requested.", "danger")
        return redirect(url_for('index'))

    requisitions = Requisition.query.all()
    purchase_orders = PurchaseOrder.query.all()
    batches = BatchSummary.query.all()
    statements = VendorStatement.query.all()

    if request.method == 'POST':
        action = request.form.get('action', 'save')
        file = request.files.get('pdf_file')
        filename = request.form.get('existing_pdf_filename') or None

        new_saved = save_uploaded_pdf(file)
        if new_saved:
            filename = new_saved

        if action == 'autofill':
            if not filename:
                flash('Please select a PDF file first to auto-fill.', 'warning')
                return redirect(url_for('add_entity', entity_type=entity_type))

            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            prefill = extract_pdf_data(filepath, entity_type)
            prefill['pdf_filename'] = filename

            if prefill.get('raw_extracted_text'):
                flash(f"Auto-filled using {prefill.get('extraction_method')}! Please review extracted fields below.", 'info')
            else:
                flash('No readable text found in this PDF.', 'warning')

            return render_template('add_edit.html', entity_type=entity_type, item=prefill,
                                   requisitions=requisitions, purchase_orders=purchase_orders,
                                   batches=batches, statements=statements)

        f_area = request.form.get('functional_area')
        new_item = None

        if entity_type == 'requisitions':
            req_num = (request.form.get('requisition_number') or '').strip()
            if not req_num:
                flash('Requisition Number is required.', 'danger')
                return render_template('add_edit.html', entity_type=entity_type, item=request.form,
                                       requisitions=requisitions, purchase_orders=purchase_orders,
                                       batches=batches, statements=statements)
            new_item = Requisition(
                requisition_number=req_num,
                transaction_date=request.form.get('transaction_date'),
                quotation_number=request.form.get('quotation_number'),
                poid=request.form.get('poid'),
                vendor_name=request.form.get('vendor_name'),
                description=request.form.get('description'),
                value=parse_amount(request.form.get('value')),
                functional_area=f_area,
                remarks=request.form.get('remarks'),
                date_sent_to_bursary=request.form.get('date_sent_to_bursary'),
                pdf_filename=filename
            )
        elif entity_type == 'purchase_orders':
            po_id = (request.form.get('poid') or '').strip()
            if not po_id:
                flash('POID is required.', 'danger')
                return render_template('add_edit.html', entity_type=entity_type, item=request.form,
                                       requisitions=requisitions, purchase_orders=purchase_orders,
                                       batches=batches, statements=statements)
            new_item = PurchaseOrder(
                poid=po_id,
                requisition_id=parse_int_or_none(request.form.get('requisition_id')),
                vendor_name=request.form.get('vendor_name'),
                po_amount=parse_amount(request.form.get('po_amount')),
                date=request.form.get('date'),
                service=request.form.get('service'),
                functional_area=f_area,
                remarks=request.form.get('remarks'),
                pdf_filename=filename
            )
        elif entity_type == 'invoices':
            inv_num = (request.form.get('invoice_number') or '').strip()
            if not inv_num:
                flash('Invoice Number is required.', 'danger')
                return render_template('add_edit.html', entity_type=entity_type, item=request.form,
                                       requisitions=requisitions, purchase_orders=purchase_orders,
                                       batches=batches, statements=statements)
            new_item = Invoice(
                invoice_number=inv_num,
                invoice_date=request.form.get('invoice_date'),
                vendor_name=request.form.get('vendor_name'),
                invoice_amount=parse_amount(request.form.get('invoice_amount')),
                date_received=request.form.get('date_received'),
                invoice_details=request.form.get('invoice_details'),
                purchase_order_id=parse_int_or_none(request.form.get('purchase_order_id')),
                vendor_statement_id=parse_int_or_none(request.form.get('vendor_statement_id')),
                remarks=request.form.get('remarks'),
                batch_summary_id=parse_int_or_none(request.form.get('batch_summary_id')),
                date_sent_to_nekeisha=request.form.get('date_sent_to_nekeisha'),
                date_returned_from_nekeisha=request.form.get('date_returned_from_nekeisha'),
                date_sent_to_ta=request.form.get('date_sent_to_ta'),
                date_returned_from_ta=request.form.get('date_returned_from_ta'),
                date_sent_to_om=request.form.get('date_sent_to_om'),
                date_returned_from_om=request.form.get('date_returned_from_om'),
                date_submitted_to_bursary=request.form.get('date_submitted_to_bursary'),
                bursary_date_batch=request.form.get('bursary_date_batch'),
                payment_received=request.form.get('payment_received'),
                pdf_filename=filename
            )
        elif entity_type == 'batches':
            batch_num = (request.form.get('batch_number') or '').strip()
            if not batch_num:
                flash('Batch Number is required.', 'danger')
                return render_template('add_edit.html', entity_type=entity_type, item=request.form,
                                       requisitions=requisitions, purchase_orders=purchase_orders,
                                       batches=batches, statements=statements)
            new_item = BatchSummary(
                batch_number=batch_num,
                functional_area=f_area,
                period_collected=request.form.get('period_collected'),
                date_sent_for_review=request.form.get('date_sent_for_review'),
                date_reviewed=request.form.get('date_reviewed'),
                date_sent_to_ta=request.form.get('date_sent_to_ta'),
                date_received_from_ta=request.form.get('date_received_from_ta'),
                date_sent_to_om=request.form.get('date_sent_to_om'),
                date_received_from_om=request.form.get('date_received_from_om'),
                submission_date=request.form.get('submission_date'),
                remarks=request.form.get('remarks')
            )
        elif entity_type == 'statements':
            stmt_num = (request.form.get('statement_number') or '').strip()
            if not stmt_num:
                flash('Statement Number is required.', 'danger')
                return render_template('add_edit.html', entity_type=entity_type, item=request.form,
                                       requisitions=requisitions, purchase_orders=purchase_orders,
                                       batches=batches, statements=statements)
            new_item = VendorStatement(
                statement_number=stmt_num,
                functional_area=f_area,
                vendor=request.form.get('vendor'),
                statement_period=request.form.get('statement_period'),
                date_received=request.form.get('date_received'),
                amount=parse_amount(request.form.get('amount')),
                remarks=request.form.get('remarks'),
                pdf_filename=filename
            )

        try:
            db.session.add(new_item)
            db.session.commit()
            sync_requisition_po_links()

            if entity_type == 'statements':
                report_link = url_for('statement_report', statement_id=new_item.id)
                flash(f'Vendor Statement successfully added! <a href="{report_link}" class="alert-link fw-bold">View Statement Report &rarr;</a>', 'success')
            else:
                flash('Record successfully added!', 'success')

            return redirect(url_for('list_entities', entity_type=entity_type))
        except IntegrityError:
            db.session.rollback()
            flash('Error: A record with that Document/Reference Number already exists in the system.', 'danger')
            form_data = dict(request.form)
            form_data['pdf_filename'] = filename
            return render_template('add_edit.html', entity_type=entity_type, item=form_data,
                                   requisitions=requisitions, purchase_orders=purchase_orders,
                                   batches=batches, statements=statements)

    return render_template('add_edit.html', entity_type=entity_type, item=None,
                           requisitions=requisitions, purchase_orders=purchase_orders,
                           batches=batches, statements=statements)


@app.route('/<entity_type>/edit/<int:id>', methods=['GET', 'POST'])
def edit_entity(entity_type, id):
    if entity_type not in VALID_ENTITIES:
        flash("Invalid section requested.", "danger")
        return redirect(url_for('index'))

    if entity_type == 'requisitions':
        item = Requisition.query.get_or_404(id)
    elif entity_type == 'purchase_orders':
        item = PurchaseOrder.query.get_or_404(id)
    elif entity_type == 'invoices':
        item = Invoice.query.get_or_404(id)
    elif entity_type == 'batches':
        item = BatchSummary.query.get_or_404(id)
    elif entity_type == 'statements':
        item = VendorStatement.query.get_or_404(id)

    requisitions = Requisition.query.all()
    purchase_orders = PurchaseOrder.query.all()
    batches = BatchSummary.query.all()
    statements = VendorStatement.query.all()

    if request.method == 'POST':
        file = request.files.get('pdf_file')
        new_saved = save_uploaded_pdf(file)
        if new_saved:
            if hasattr(item, 'pdf_filename') and item.pdf_filename:
                delete_pdf_safely(item.pdf_filename)
            item.pdf_filename = new_saved

        if entity_type != 'invoices':
            item.functional_area = request.form.get('functional_area')

        if entity_type == 'requisitions':
            item.requisition_number = (request.form.get('requisition_number') or '').strip()
            item.transaction_date = request.form.get('transaction_date')
            item.quotation_number = request.form.get('quotation_number')
            item.poid = request.form.get('poid')
            item.vendor_name = request.form.get('vendor_name')
            item.description = request.form.get('description')
            item.value = parse_amount(request.form.get('value'))
            item.remarks = request.form.get('remarks')
            item.date_sent_to_bursary = request.form.get('date_sent_to_bursary')
        elif entity_type == 'purchase_orders':
            item.poid = (request.form.get('poid') or '').strip()
            item.requisition_id = parse_int_or_none(request.form.get('requisition_id'))
            item.vendor_name = request.form.get('vendor_name')
            item.po_amount = parse_amount(request.form.get('po_amount'))
            item.date = request.form.get('date')
            item.service = request.form.get('service')
            item.remarks = request.form.get('remarks')
        elif entity_type == 'invoices':
            item.invoice_number = (request.form.get('invoice_number') or '').strip()
            item.invoice_date = request.form.get('invoice_date')
            item.vendor_name = request.form.get('vendor_name')
            item.invoice_amount = parse_amount(request.form.get('invoice_amount'))
            item.date_received = request.form.get('date_received')
            item.invoice_details = request.form.get('invoice_details')
            item.purchase_order_id = parse_int_or_none(request.form.get('purchase_order_id'))
            item.vendor_statement_id = parse_int_or_none(request.form.get('vendor_statement_id'))
            item.remarks = request.form.get('remarks')
            item.batch_summary_id = parse_int_or_none(request.form.get('batch_summary_id'))
            item.date_sent_to_nekeisha = request.form.get('date_sent_to_nekeisha')
            item.date_returned_from_nekeisha = request.form.get('date_returned_from_nekeisha')
            item.date_sent_to_ta = request.form.get('date_sent_to_ta')
            item.date_returned_from_ta = request.form.get('date_returned_from_ta')
            item.date_sent_to_om = request.form.get('date_sent_to_om')
            item.date_returned_from_om = request.form.get('date_returned_from_om')
            item.date_submitted_to_bursary = request.form.get('date_submitted_to_bursary')
            item.bursary_date_batch = request.form.get('bursary_date_batch')
            item.payment_received = request.form.get('payment_received')
        elif entity_type == 'batches':
            item.batch_number = (request.form.get('batch_number') or '').strip()
            item.period_collected = request.form.get('period_collected')
            item.date_sent_for_review = request.form.get('date_sent_for_review')
            item.date_reviewed = request.form.get('date_reviewed')
            item.date_sent_to_ta = request.form.get('date_sent_to_ta')
            item.date_received_from_ta = request.form.get('date_received_from_ta')
            item.date_sent_to_om = request.form.get('date_sent_to_om')
            item.date_received_from_om = request.form.get('date_received_from_om')
            item.submission_date = request.form.get('submission_date')
            item.remarks = request.form.get('remarks')
        elif entity_type == 'statements':
            item.statement_number = (request.form.get('statement_number') or '').strip()
            item.vendor = request.form.get('vendor')
            item.statement_period = request.form.get('statement_period')
            item.date_received = request.form.get('date_received')
            item.amount = parse_amount(request.form.get('amount'))
            item.remarks = request.form.get('remarks')

        try:
            db.session.commit()
            sync_requisition_po_links()

            if entity_type == 'statements':
                report_link = url_for('statement_report', statement_id=item.id)
                flash(f'Vendor Statement successfully updated! <a href="{report_link}" class="alert-link fw-bold">View Statement Report &rarr;</a>', 'success')
            else:
                flash('Record successfully updated!', 'success')

            return redirect(url_for('list_entities', entity_type=entity_type))
        except IntegrityError:
            db.session.rollback()
            flash('Error: Another record already uses that Document/Reference Number.', 'danger')

    return render_template('add_edit.html', entity_type=entity_type, item=item,
                           requisitions=requisitions, purchase_orders=purchase_orders,
                           batches=batches, statements=statements)


@app.route('/<entity_type>/delete/<int:id>', methods=['POST'])
def delete_entity(entity_type, id):
    if entity_type not in VALID_ENTITIES:
        flash("Invalid section requested.", "danger")
        return redirect(url_for('index'))

    if entity_type == 'requisitions':
        item = Requisition.query.get_or_404(id)
    elif entity_type == 'purchase_orders':
        item = PurchaseOrder.query.get_or_404(id)
    elif entity_type == 'invoices':
        item = Invoice.query.get_or_404(id)
    elif entity_type == 'batches':
        item = BatchSummary.query.get_or_404(id)
    elif entity_type == 'statements':
        item = VendorStatement.query.get_or_404(id)

    if hasattr(item, 'pdf_filename') and item.pdf_filename:
        delete_pdf_safely(item.pdf_filename)

    db.session.delete(item)
    db.session.commit()
    flash('Record successfully deleted!', 'success')
    return redirect(url_for('list_entities', entity_type=entity_type))


# --- VENDOR STATEMENT REPORT & RAW PDF TEXT ROUTES ---

@app.route('/statements/<int:statement_id>/raw_text', methods=['GET'])
def statement_raw_text(statement_id):
    statement = VendorStatement.query.get_or_404(statement_id)
    if not statement.pdf_filename:
        raw_text = "No PDF file is attached to this Vendor Statement."
        method = "None"
    else:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], statement.pdf_filename)
        raw_text, method = read_pdf_text_with_ocr_fallback(filepath, max_pages=10)
        if not raw_text.strip():
            raw_text = "[No readable text extracted. If this is a scanned PDF, ensure Tesseract and Poppler are installed and OCR is enabled.]"

    return render_template(
        'statement_raw_text.html',
        statement=statement,
        raw_text=raw_text,
        extraction_method=method
    )


@app.route('/statements/report/<int:statement_id>', methods=['GET'])
def statement_report(statement_id):
    statement = VendorStatement.query.get_or_404(statement_id)
    target_month = request.args.get('target_month')
    ref_date = f"{target_month}-15" if target_month else None
    target_15th = get_target_15th(ref_date)

    all_invoices = Invoice.query.all()
    matched_invoices, pdf_invoices, known_inv_keys = reconcile_statement_with_invoices_section(statement, all_invoices)

    eligible_entries = get_statement_eligible_invoices(statement_id, target_15th)
    eligible_ids = {e['invoice'].id for e in eligible_entries}
    ineligible_invoices = [inv for inv in matched_invoices if inv.id not in eligible_ids]

    pdf_only_invoices = [
        p for p in pdf_invoices
        if (p.get('invoice_number') or '').strip().upper() not in known_inv_keys
    ]

    total_amount = sum((e['invoice'].invoice_amount or 0.0) for e in eligible_entries)

    return render_template(
        'statement_report.html',
        statement=statement,
        eligible_entries=eligible_entries,
        ineligible_invoices=ineligible_invoices,
        pdf_invoices=pdf_invoices,
        pdf_only_invoices=pdf_only_invoices,
        total_amount=total_amount,
        target_15th=target_15th.strftime('%Y-%m-%d'),
        selected_month=target_15th.strftime('%Y-%m'),
        format_date_us=format_date_us
    )


@app.route('/statements/report/<int:statement_id>/export_csv', methods=['GET'])
def export_statement_csv(statement_id):
    statement = VendorStatement.query.get_or_404(statement_id)
    target_month = request.args.get('target_month')
    ref_date = f"{target_month}-15" if target_month else None
    target_15th = get_target_15th(ref_date)
    eligible_entries = get_statement_eligible_invoices(statement_id, target_15th)

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['AkiraZ Procurement System - Vendor Statement Report', statement.statement_number, 'Vendor:', statement.vendor])
    writer.writerow(['Target 15th Month:', target_15th.strftime('%B %Y'), 'Rule:', 'Submitted >= 28 days before 15th'])
    writer.writerow([])
    writer.writerow(['Invoice Number', 'Invoice Date', 'Vendor Name', 'Invoice Amount', 'Date Submitted to Bursary', 'Days Before 15th', 'POID', 'Remarks'])

    for entry in eligible_entries:
        inv = entry['invoice']
        writer.writerow([
            inv.invoice_number, format_date_us(inv.invoice_date), inv.vendor_name,
            f"{inv.invoice_amount or 0:.2f}", format_date_us(inv.date_submitted_to_bursary),
            entry['days_before_15th'], inv.purchase_order.poid if inv.purchase_order else 'None', inv.remarks
        ])

    filename = f"AkiraZ_Statement_{statement.statement_number}_Report_{target_15th}.csv"
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment;filename={filename}"}
    )


# --- REPORTS & AUDIT HUB ROUTES ---

@app.route('/reports')
def reports_hub():
    target_month = request.args.get('target_month')
    ref_date = f"{target_month}-15" if target_month else None
    target_15th = get_target_15th(ref_date)

    cutoff_date_str = (request.args.get('cutoff_date') or target_15th.strftime('%Y-%m-%d')).strip()
    try:
        cutoff_dt = datetime.strptime(cutoff_date_str[:10], '%Y-%m-%d').date()
    except ValueError:
        cutoff_dt = target_15th
        cutoff_date_str = cutoff_dt.strftime('%Y-%m-%d')

    outstanding_reqs = Requisition.query.outerjoin(PurchaseOrder).filter(
        PurchaseOrder.id == None,
        or_(Requisition.poid == None, Requisition.poid == '')
    ).all()
    outstanding_reqs_total = sum((r.value or 0.0) for r in outstanding_reqs)

    all_pos = PurchaseOrder.query.all()
    po_id_by_upper = {
        (po.poid or '').strip().upper(): po.id
        for po in all_pos if (po.poid or '').strip()
    }

    po_report_data = []
    po_overruns = []
    grand_po_total = 0.0
    grand_inv_total = 0.0

    for po in all_pos:
        linked_invoices = list(po.invoices)
        if not linked_invoices and po.poid:
            linked_invoices = [
                inv for inv in Invoice.query.all()
                if inv.purchase_order and inv.purchase_order.poid
                and inv.purchase_order.poid.strip().lower() == po.poid.strip().lower()
            ]

        inv_total = sum((inv.invoice_amount or 0.0) for inv in linked_invoices)
        po_amt = po.po_amount or 0.0
        balance = po_amt - inv_total
        grand_po_total += po_amt
        grand_inv_total += inv_total

        entry = {
            'po': po,
            'invoices': linked_invoices,
            'po_total': po_amt,
            'invoices_total': inv_total,
            'balance': balance
        }
        po_report_data.append(entry)
        if balance < 0:
            po_overruns.append(entry)

    all_invoices = Invoice.query.all()
    invoices_grand_total = sum((inv.invoice_amount or 0.0) for inv in all_invoices)
    unsubmitted_invoices = [
        inv for inv in all_invoices
        if not (inv.date_submitted_to_bursary and str(inv.date_submitted_to_bursary).strip())
    ]

    statements = VendorStatement.query.order_by(VendorStatement.statement_number.asc()).all()
    available_periods = sorted({
        (s.statement_period or '').strip()
        for s in statements
        if (s.statement_period or '').strip()
    }, reverse=True)

    selected_period = (request.args.get('statement_period') or '').strip()
    selected_statement_ids = [
        int(x) for x in request.args.getlist('statement_ids')
        if str(x).strip().isdigit()
    ]

    if selected_statement_ids:
        selected_statements = VendorStatement.query.filter(
            VendorStatement.id.in_(selected_statement_ids)
        ).order_by(VendorStatement.statement_number.asc()).all()
    elif selected_period:
        selected_statements = [
            s for s in statements
            if (s.statement_period or '').strip().lower() == selected_period.lower()
        ]
        selected_statement_ids = [s.id for s in selected_statements]
    else:
        selected_statements = statements

    invoice_reporting_rows = []
    invoice_reporting_total = 0.0
    before_cutoff_count = 0
    linked_po_count = 0

    for stmt in selected_statements:
        matched_invoices, pdf_invoices, known_inv_keys = reconcile_statement_with_invoices_section(stmt, all_invoices)

        stmt_pdf_text = ""
        if stmt.pdf_filename:
            stmt_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], stmt.pdf_filename)
            stmt_pdf_text, _ = read_pdf_text_with_ocr_fallback(stmt_pdf_path, max_pages=5)

        for inv in matched_invoices:
            sub_str = (inv.date_submitted_to_bursary or '').strip()[:10]
            before_cutoff = False
            if sub_str:
                try:
                    sub_dt = datetime.strptime(sub_str, '%Y-%m-%d').date()
                    before_cutoff = sub_dt <= cutoff_dt
                except ValueError:
                    before_cutoff = False

            if before_cutoff:
                before_cutoff_count += 1

            amt = float(inv.invoice_amount or 0.0)
            invoice_reporting_total += amt

            inv_vendor = inv.vendor_name or stmt.vendor or ''
            ctx_snippet = ""
            if stmt_pdf_text and inv.invoice_number:
                idx = stmt_pdf_text.upper().find(inv.invoice_number.strip().upper())
                if idx != -1:
                    ctx_snippet = stmt_pdf_text[max(0, idx - 60): idx + 180]

            resolved_poid = resolve_po_number_for_invoice(
                inv=inv,
                vendor_name=inv_vendor,
                pdf_context_text=ctx_snippet,
                all_pos=all_pos
            )
            if inv.purchase_order_id is None and resolved_poid.upper() in po_id_by_upper:
                inv.purchase_order_id = po_id_by_upper[resolved_poid.upper()]
                linked_po_count += 1

            invoice_reporting_rows.append({
                'invoice_number': inv.invoice_number,
                'vendor': inv_vendor,
                'invoice_date': inv.invoice_date,
                'date_received': inv.date_received or '',
                'poid': resolved_poid or 'N/A',
                'date_submitted': sub_str,
                'amount': amt,
                'before_cutoff': before_cutoff,
                'source': 'Invoices Section'
            })

        for pdf_inv in pdf_invoices:
            pdf_inv_num = (pdf_inv.get('invoice_number') or '').strip()
            pdf_key = pdf_inv_num.upper()
            if pdf_key and pdf_key not in known_inv_keys:
                known_inv_keys.add(pdf_key)
                amt = float(pdf_inv.get('invoice_amount') or 0.0)
                invoice_reporting_total += amt

                ctx_snippet = ""
                if stmt_pdf_text:
                    idx = stmt_pdf_text.upper().find(pdf_key)
                    if idx != -1:
                        ctx_snippet = stmt_pdf_text[max(0, idx - 60): idx + 180]

                resolved_poid = resolve_po_number_for_invoice(
                    inv=None,
                    vendor_name=stmt.vendor or '',
                    pdf_context_text=ctx_snippet,
                    all_pos=all_pos
                )

                invoice_reporting_rows.append({
                    'invoice_number': pdf_inv_num,
                    'vendor': stmt.vendor or '',
                    'invoice_date': pdf_inv.get('invoice_date') or '',
                    'date_received': '',
                    'poid': resolved_poid or 'N/A',
                    'date_submitted': '',
                    'amount': amt,
                    'before_cutoff': False,
                    'source': 'Statement PDF'
                })

    if linked_po_count:
        db.session.commit()

    return render_template(
        'reports.html',
        outstanding_reqs=outstanding_reqs,
        outstanding_reqs_total=outstanding_reqs_total,
        po_report_data=po_report_data,
        po_overruns=po_overruns,
        grand_po_total=grand_po_total,
        grand_inv_total=grand_inv_total,
        all_invoices=all_invoices,
        invoices_grand_total=invoices_grand_total,
        unsubmitted_invoices=unsubmitted_invoices,
        statements=statements,
        available_periods=available_periods,
        selected_period=selected_period,
        selected_statement_ids=selected_statement_ids,
        invoice_reporting_rows=invoice_reporting_rows,
        invoice_reporting_total=invoice_reporting_total,
        before_cutoff_count=before_cutoff_count,
        cutoff_date=cutoff_date_str,
        target_15th=target_15th.strftime('%Y-%m-%d'),
        selected_month=target_15th.strftime('%Y-%m'),
        format_date_us=format_date_us
    )


@app.route('/reports/export/<report_type>')
def export_report(report_type):
    output = io.StringIO()
    writer = csv.writer(output)

    if report_type == 'requisitions':
        writer.writerow(['Requisition Number', 'Transaction Date', 'Quotation Number', 'POID', 'Vendor Name', 'Description', 'Value', 'Functional Area', 'Remarks', 'Date Sent to Bursary'])
        reqs = Requisition.query.outerjoin(PurchaseOrder).filter(
            PurchaseOrder.id == None,
            or_(Requisition.poid == None, Requisition.poid == '')
        ).all()
        for r in reqs:
            writer.writerow([
                r.requisition_number, format_date_us(r.transaction_date), r.quotation_number, r.poid,
                r.vendor_name, r.description, f"{r.value or 0:.2f}", r.functional_area, r.remarks,
                format_date_us(r.date_sent_to_bursary)
            ])
        filename = f"AkiraZ_outstanding_requisitions_{date.today()}.csv"

    elif report_type == 'purchase_orders':
        writer.writerow([
            'POID', 'Requisition Number', 'Vendor Name', 'PO Amount', 'Date', 'Service', 'Functional Area', 'Remarks'
        ])
        for po in PurchaseOrder.query.all():
            writer.writerow([
                po.poid,
                po.requisition.requisition_number if po.requisition else 'None',
                po.vendor_name, f"{po.po_amount or 0:.2f}", format_date_us(po.date),
                po.service, po.functional_area, po.remarks
            ])
        filename = f"AkiraZ_purchase_order_report_{date.today()}.csv"

    elif report_type == 'invoices':
        writer.writerow([
            'Invoice Number', 'Invoice Date', 'Vendor Name', 'Invoice Amount', 'Date Received',
            'Invoice Details', 'POID', 'Remarks', 'Batch ID (Lookup)', 'Date Sent to Nekeisha',
            'Date Returned from Nekeisha', 'Date Sent to TA', 'Date Returned from TA',
            'Date Sent to OM', 'Date Returned from OM', 'Date Submitted to Bursary',
            'Bursary Date Batch', 'Payment Received'
        ])
        for inv in Invoice.query.all():
            writer.writerow([
                inv.invoice_number, format_date_us(inv.invoice_date), inv.vendor_name,
                f"{inv.invoice_amount or 0:.2f}", format_date_us(inv.date_received),
                inv.invoice_details, inv.purchase_order.poid if inv.purchase_order else 'None',
                inv.remarks, inv.batch.batch_number if inv.batch else 'None',
                format_date_us(inv.date_sent_to_nekeisha), format_date_us(inv.date_returned_from_nekeisha),
                format_date_us(inv.date_sent_to_ta), format_date_us(inv.date_returned_from_ta),
                format_date_us(inv.date_sent_to_om), format_date_us(inv.date_returned_from_om),
                format_date_us(inv.date_submitted_to_bursary), inv.bursary_date_batch,
                format_date_us(inv.payment_received)
            ])
        filename = f"AkiraZ_invoices_report_{date.today()}.csv"
    else:
        flash("Invalid report type", "danger")
        return redirect(url_for('reports_hub'))

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment;filename={filename}"}
    )


@app.route('/reports/export_excel', endpoint='export_excel_report')
def export_excel_report():
    target_month = request.args.get('target_month')
    ref_date = f"{target_month}-15" if target_month else None
    target_15th = get_target_15th(ref_date)

    wb = openpyxl.Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")
    bold_font = Font(bold=True)

    def style_sheet(ws, currency_cols=None, sum_cols=None):
        currency_cols = currency_cols or []
        sum_cols = sum_cols or []
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for col_idx in currency_cols:
                cell = row[col_idx - 1]
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '$#,##0.00'
        if sum_cols and ws.max_row >= 2:
            last_data_row = ws.max_row
            total_row = last_data_row + 1
            ws.cell(row=total_row, column=1, value="TOTAL").font = bold_font
            for col_idx in sum_cols:
                col_letter = get_column_letter(col_idx)
                c = ws.cell(row=total_row, column=col_idx, value=f"=SUM({col_letter}2:{col_letter}{last_data_row})")
                c.font = bold_font
                c.number_format = '$#,##0.00'
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 45)

    ws1 = wb.active
    ws1.title = "Outstanding Requisitions"
    ws1.append(['Requisition Number', 'Transaction Date', 'Quotation Number', 'POID', 'Vendor Name', 'Description', 'Value', 'Functional Area', 'Remarks', 'Date Sent to Bursary'])
    reqs = Requisition.query.outerjoin(PurchaseOrder).filter(
        PurchaseOrder.id == None,
        or_(Requisition.poid == None, Requisition.poid == '')
    ).all()
    for r in reqs:
        ws1.append([
            r.requisition_number, format_date_us(r.transaction_date), r.quotation_number, r.poid,
            r.vendor_name, r.description, float(r.value or 0), r.functional_area, r.remarks,
            format_date_us(r.date_sent_to_bursary)
        ])
    style_sheet(ws1, currency_cols=[7], sum_cols=[7])

    ws2 = wb.create_sheet(title="PO vs Invoiced")
    ws2.append(['POID', 'Requisition Number', 'Vendor Name', 'Service', 'Functional Area', 'PO Amount', 'Invoiced Total', 'Balance Remaining'])
    for po in PurchaseOrder.query.all():
        linked_invoices = list(po.invoices)
        if not linked_invoices and po.poid:
            linked_invoices = [inv for inv in Invoice.query.all() if inv.purchase_order and inv.purchase_order.poid and inv.purchase_order.poid.strip().lower() == po.poid.strip().lower()]
        inv_tot = sum((inv.invoice_amount or 0.0) for inv in linked_invoices)
        po_amt = float(po.po_amount or 0)
        ws2.append([
            po.poid,
            po.requisition.requisition_number if po.requisition else 'None',
            po.vendor_name, po.service, po.functional_area,
            po_amt, float(inv_tot), float(po_amt - inv_tot)
        ])
    style_sheet(ws2, currency_cols=[6, 7, 8], sum_cols=[6, 7, 8])

    ws3 = wb.create_sheet(title="Invoices Audit Trail")
    ws3.append([
        'Invoice Number', 'Invoice Date', 'Vendor Name', 'Invoice Amount', 'POID',
        'Statement #', 'Batch #', 'Sent to Nekeisha', 'Returned from Nekeisha',
        'Sent to TA', 'Returned from TA', 'Sent to OM', 'Returned from OM',
        'Submitted to Bursary', 'Payment Received'
    ])
    for inv in Invoice.query.all():
        ws3.append([
            inv.invoice_number, format_date_us(inv.invoice_date), inv.vendor_name,
            float(inv.invoice_amount or 0), inv.purchase_order.poid if inv.purchase_order else 'None',
            inv.statement.statement_number if inv.statement else 'None',
            inv.batch.batch_number if inv.batch else 'None',
            format_date_us(inv.date_sent_to_nekeisha), format_date_us(inv.date_returned_from_nekeisha),
            format_date_us(inv.date_sent_to_ta), format_date_us(inv.date_returned_from_ta),
            format_date_us(inv.date_sent_to_om), format_date_us(inv.date_returned_from_om),
            format_date_us(inv.date_submitted_to_bursary), format_date_us(inv.payment_received)
        ])
    style_sheet(ws3, currency_cols=[4], sum_cols=[4])

    ws4 = wb.create_sheet(title="28-Day Statement Audit")
    ws4.append(['Statement Number', 'Vendor', 'Statement Period', 'Statement Amount', f'Qualifying Invoices (>=28d before {target_15th})', 'Qualifying Amount'])
    for stmt in VendorStatement.query.all():
        qual = get_statement_eligible_invoices(stmt.id, target_15th)
        qual_amt = sum((q['invoice'].invoice_amount or 0.0) for q in qual)
        ws4.append([
            stmt.statement_number, stmt.vendor, stmt.statement_period,
            float(stmt.amount or 0), len(qual), float(qual_amt)
        ])
    style_sheet(ws4, currency_cols=[4, 6], sum_cols=[4, 6])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"AkiraZ_Procurement_Bursary_Report_{target_15th}.xlsx"
    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment;filename={filename}"}
    )


@app.route('/backup_db')
def backup_db():
    db_path = os.path.join(app.instance_path, 'procurement.db')
    if not os.path.exists(db_path):
        db_path = 'procurement.db'
    if not os.path.exists(db_path):
        flash("Database file not found for backup.", "danger")
        return redirect(url_for('index'))

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(
        db_path,
        as_attachment=True,
        download_name=f"AkiraZ_procurement_backup_{timestamp}.db",
        mimetype="application/octet-stream"
    )


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


if __name__ == '__main__':
    ensure_schema_updates()
    with app.app_context():
        db.create_all()
        sync_requisition_po_links()
    app.run(debug=True)