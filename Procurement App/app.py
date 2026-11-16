import os
import csv
import io
import re
import sqlite3
from datetime import datetime, date
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, send_file, Response
from flask_sqlalchemy import SQLAlchemy
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
app.config['SECRET_KEY'] = 'procurement-secret-key-change-me'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///procurement.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = os.path.join('static', 'uploads')
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'xlsx', 'xls', 'csv'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
db = SQLAlchemy(app)

VALID_ENTITIES = {'requisitions', 'purchase_orders', 'invoices', 'batches', 'statements'}

def allowed_file(filename, allowed_set={'pdf'}):
    return bool(filename) and '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

def parse_amount(val):
    if val is None:
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
            if 'date' in cols and 'transaction_date' not in cols:
                cursor.execute("ALTER TABLE requisition RENAME COLUMN date TO transaction_date")
            if 'job' in cols and 'description' not in cols:
                cursor.execute("ALTER TABLE requisition RENAME COLUMN job TO description")
            if 'amount' in cols and 'value' not in cols:
                cursor.execute("ALTER TABLE requisition RENAME COLUMN amount TO value")
            if 'vendor' in cols and 'vendor_name' not in cols:
                cursor.execute("ALTER TABLE requisition RENAME COLUMN vendor TO vendor_name")

            cursor.execute("PRAGMA table_info(purchase_order)")
            po_cols = [row[1] for row in cursor.fetchall()]
            if 'po_number' in po_cols and 'poid' not in po_cols:
                cursor.execute("ALTER TABLE purchase_order RENAME COLUMN po_number TO poid")
            if 'vendor' in po_cols and 'vendor_name' not in po_cols:
                cursor.execute("ALTER TABLE purchase_order RENAME COLUMN vendor TO vendor_name")
            if 'amount' in po_cols and 'po_amount' not in po_cols:
                cursor.execute("ALTER TABLE purchase_order RENAME COLUMN amount TO po_amount")
            if 'job' in po_cols and 'service' not in po_cols:
                cursor.execute("ALTER TABLE purchase_order RENAME COLUMN job TO service")

            cursor.execute("PRAGMA table_info(invoice)")
            inv_cols = [row[1] for row in cursor.fetchall()]
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

            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Schema check warning: {e}")

# --- UPGRADED PDF EXTRACTION HELPER ---

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

def extract_pdf_data(filepath, entity_type):
    full_text = ""
    extraction_method = "Digital Text"

    if not os.path.exists(filepath):
        return {'raw_extracted_text': '', 'extraction_method': 'File Not Found'}

    try:
        reader = PdfReader(filepath)
        for page in reader.pages:
            full_text += (page.extract_text() or "") + "\n"
    except Exception as e:
        print(f"Error reading digital PDF: {e}")

    if len(full_text.strip()) < 25 and OCR_AVAILABLE:
        try:
            images = convert_from_path(filepath, first_page=1, last_page=3)
            ocr_pages = [pytesseract.image_to_string(img) for img in images]
            full_text = "\n".join(ocr_pages)
            extraction_method = "Scanned OCR"
        except Exception as e:
            print(f"OCR fallback skipped/failed: {e}")

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
        m = re.search(r'(?:Invoice|Inv)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z0-9\-\/]+)', full_text, re.IGNORECASE)
        if m:
            extracted['invoice_number'] = m.group(1).strip()
        po_match = re.search(r'(?:Purchase\s*Order|P\.?O\.?)\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z0-9\-\/]+)', full_text, re.IGNORECASE)
        if po_match:
            po_obj = PurchaseOrder.query.filter(PurchaseOrder.poid.ilike(f"%{po_match.group(1).strip()}%")).first()
            if po_obj:
                extracted['purchase_order_id'] = po_obj.id
                if po_obj.vendor_name:
                    extracted['vendor_name'] = po_obj.vendor_name

    v = re.search(r'(?:Vendor|Supplier|From|Company|Payable\s*To)\s*[:\-]?\s*([A-Za-z0-9\s\.,&\-\']+)', full_text, re.IGNORECASE)
    if v:
        extracted['vendor_name'] = v.group(1).splitlines()[0].strip()[:100]
    else:
        for line in full_text.splitlines():
            clean_line = line.strip()
            if len(clean_line) > 3 and not re.match(r'^(invoice|quotation|quote|statement|purchase order|tax invoice|page\s+\d+)', clean_line, re.IGNORECASE):
                extracted['vendor_name'] = clean_line[:100]
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
    else:
        all_amounts = re.findall(r'(?:BBD)?\$\s*([\d,]+\.\d{2})', full_text, re.IGNORECASE)
        if all_amounts:
            val = max(parse_amount(a) for a in all_amounts)
            extracted['value'] = val
            extracted['po_amount'] = val
            extracted['invoice_amount'] = val

    return extracted

# --- HELPER FUNCTIONS FOR REPORTS ---

def get_target_15th(ref_date=None):
    if ref_date:
        try:
            dt = datetime.strptime(ref_date, '%Y-%m-%d').date()
            return dt.replace(day=15)
        except ValueError:
            pass
    today = date.today()
    return today.replace(day=15)

def get_overdue_invoices(target_15th):
    all_invoices = Invoice.query.all()
    qualifying = []
    for inv in all_invoices:
        sub_str = inv.date_submitted_to_bursary or inv.invoice_date
        if not sub_str:
            continue
        try:
            sub_date = datetime.strptime(sub_str, '%Y-%m-%d').date()
            days_before_15th = (target_15th - sub_date).days
            if days_before_15th > 28:
                qualifying.append({
                    'invoice': inv,
                    'submission_date_used': sub_str,
                    'date_type': 'Bursary Submission' if inv.date_submitted_to_bursary else 'Invoice Date',
                    'days_before_15th': days_before_15th
                })
        except ValueError:
            continue
    return qualifying

# --- ROUTES ---

@app.route('/')
def index():
    unsubmitted_invoices = Invoice.query.filter(
        (Invoice.date_submitted_to_bursary == None) | (Invoice.date_submitted_to_bursary == '')
    ).count()

    reqs_without_po = Requisition.query.outerjoin(PurchaseOrder).filter(
        PurchaseOrder.id == None
    ).count()

    target_15th = get_target_15th()
    overdue_invoices_count = len(get_overdue_invoices(target_15th))

    areas = [
        "Custodial/Grounds",
        "Stores",
        "Administration",
        "Services- Electrical, Plumbing, Water",
        "Services- AC, Vehicles, Elevators",
        "Civil"
    ]
    po_by_area = {a: 0.0 for a in areas}
    inv_by_area = {a: 0.0 for a in areas}

    for po in PurchaseOrder.query.all():
        if po.functional_area in po_by_area:
            po_by_area[po.functional_area] += (po.po_amount or 0.0)

    for inv in Invoice.query.all():
        if inv.purchase_order and inv.purchase_order.functional_area in inv_by_area:
            inv_by_area[inv.purchase_order.functional_area] += (inv.invoice_amount or 0.0)

    turnaround_days = []
    for inv in Invoice.query.all():
        if inv.invoice_date and inv.date_submitted_to_bursary:
            try:
                d1 = datetime.strptime(inv.invoice_date, '%Y-%m-%d').date()
                d2 = datetime.strptime(inv.date_submitted_to_bursary, '%Y-%m-%d').date()
                if d2 >= d1:
                    turnaround_days.append((d2 - d1).days)
            except ValueError:
                pass

    avg_turnaround = round(sum(turnaround_days) / len(turnaround_days), 1) if turnaround_days else 0

    chart_data = {
        'labels': areas,
        'po_totals': [round(po_by_area[a], 2) for a in areas],
        'inv_totals': [round(inv_by_area[a], 2) for a in areas],
        'total_committed': sum(po_by_area.values()),
        'total_invoiced': sum(inv_by_area.values()),
        'avg_turnaround': avg_turnaround
    }

    stats = {
        'requisitions': Requisition.query.count(),
        'pos': PurchaseOrder.query.count(),
        'invoices': Invoice.query.count(),
        'batches': BatchSummary.query.count(),
        'statements': VendorStatement.query.count(),
        'unsubmitted_invoices': unsubmitted_invoices,
        'reqs_without_po': reqs_without_po,
        'overdue_invoices': overdue_invoices_count
    }
    return render_template('index.html', stats=stats, chart_data=chart_data)

@app.route('/<entity_type>')
def list_entities(entity_type):
    if entity_type not in VALID_ENTITIES:
        flash("Invalid section requested.", "danger")
        return redirect(url_for('index'))

    batches = BatchSummary.query.all() if entity_type == 'invoices' else []

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
                           batches=batches, today_str=date.today().strftime('%Y-%m-%d'),
                           format_date_us=format_date_us)

@app.route('/import/spreadsheet/<entity_type>', methods=['POST'])
def import_spreadsheet(entity_type):
    if entity_type not in VALID_ENTITIES:
        flash("Invalid entity type for import.", "danger")
        return redirect(url_for('index'))

    file = request.files.get('spreadsheet_file')
    if not file or not file.filename:
        flash("Please select an Excel (.xlsx) or CSV file to import.", "warning")
        return redirect(url_for('list_entities', entity_type=entity_type))

    try:
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
    except Exception as e:
        flash(f"Error reading file: {e}", "danger")
        return redirect(url_for('list_entities', entity_type=entity_type))

    imported_count = 0
    skipped_count = 0

    for _, row in df.iterrows():
        def val(col_names, default=""):
            if isinstance(col_names, str):
                col_names = [col_names]
            for c in col_names:
                if c in row and not pd.isna(row[c]):
                    return str(row[c]).strip()
            return default

        if entity_type == 'requisitions':
            req_num = val(['Requisition Number', 'Req #', 'requisition_number', 'Title'])
            if not req_num:
                skipped_count += 1
                continue
            if not Requisition.query.filter_by(requisition_number=req_num).first():
                db.session.add(Requisition(
                    requisition_number=req_num,
                    transaction_date=val(['Transaction Date', 'Date', 'transaction_date']),
                    quotation_number=val(['Quotation Number', 'Quote Number', 'quotation_number']),
                    poid=val(['POID', 'PO Number', 'poid']),
                    vendor_name=val(['Vendor Name', 'Vendor', 'vendor_name']),
                    description=val(['Description', 'Job', 'description']),
                    value=parse_amount(row.get('Value', row.get('Amount', 0))),
                    functional_area=val(['Functional Area', 'functional_area']),
                    remarks=val(['Remarks', 'remarks']),
                    date_sent_to_bursary=val(['Date Sent to Bursary', 'date_sent_to_bursary']) or None
                ))
                imported_count += 1
            else:
                skipped_count += 1

        elif entity_type == 'purchase_orders':
            po_num = val(['POID', 'PO Number', 'poid', 'Title'])
            if not po_num:
                skipped_count += 1
                continue
            if not PurchaseOrder.query.filter_by(poid=po_num).first():
                db.session.add(PurchaseOrder(
                    poid=po_num,
                    vendor_name=val(['Vendor Name', 'Vendor', 'vendor_name']),
                    po_amount=parse_amount(row.get('PO Amount', row.get('Amount', 0))),
                    date=val(['Date', 'date']),
                    service=val(['Service', 'Job', 'service']),
                    functional_area=val(['Functional Area', 'functional_area']),
                    remarks=val(['Remarks', 'remarks'])
                ))
                imported_count += 1
            else:
                skipped_count += 1

        elif entity_type == 'invoices':
            inv_num = val(['Invoice Number', 'Invoice #', 'invoice_number', 'Title'])
            if not inv_num:
                skipped_count += 1
                continue
            if not Invoice.query.filter_by(invoice_number=inv_num).first():
                db.session.add(Invoice(
                    invoice_number=inv_num,
                    invoice_date=val(['Invoice Date', 'Date', 'invoice_date']),
                    vendor_name=val(['Vendor Name', 'Vendor', 'vendor_name']),
                    invoice_amount=parse_amount(row.get('Invoice Amount', row.get('Amount', 0))),
                    date_received=val(['Date Received', 'date_received']),
                    invoice_details=val(['Invoice Details', 'Description', 'invoice_details']),
                    remarks=val(['Remarks', 'remarks']),
                    bursary_date_batch=val(['Bursary Date Batch', 'bursary_date_batch']),
                    payment_received=val(['Payment Received', 'payment_received']) or None
                ))
                imported_count += 1
            else:
                skipped_count += 1

    db.session.commit()
    flash(f"Spreadsheet imported successfully! Added {imported_count} new record(s) ({skipped_count} skipped/duplicates).", "success")
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
                flash('No readable text found in this PDF (it may be a scanned image without OCR enabled).', 'warning')

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
            flash('Record successfully added!', 'success')

            if entity_type == 'invoices' and new_item.purchase_order:
                po = new_item.purchase_order
                total_invoiced = sum((inv.invoice_amount or 0.0) for inv in po.invoices)
                if total_invoiced > (po.po_amount or 0.0):
                    overage = total_invoiced - (po.po_amount or 0.0)
                    flash(
                        f"Audit Warning: Total invoiced (${total_invoiced:,.2f}) exceeds PO {po.poid} approved amount (${po.po_amount or 0:,.2f}) by ${overage:,.2f}.",
                        "warning"
                    )

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
            flash('Record successfully updated!', 'success')

            if entity_type == 'invoices' and item.purchase_order:
                po = item.purchase_order
                total_invoiced = sum((inv.invoice_amount or 0.0) for inv in po.invoices)
                if total_invoiced > (po.po_amount or 0.0):
                    overage = total_invoiced - (po.po_amount or 0.0)
                    flash(
                        f"Audit Warning: Total invoiced (${total_invoiced:,.2f}) exceeds PO {po.poid} approved amount (${po.po_amount or 0:,.2f}) by ${overage:,.2f}.",
                        "warning"
                    )

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

# --- REPORTS ROUTES ---

@app.route('/reports')
def reports_hub():
    target_month = request.args.get('target_month')
    ref_date = f"{target_month}-15" if target_month else None
    target_15th = get_target_15th(ref_date)

    outstanding_reqs = Requisition.query.outerjoin(PurchaseOrder).filter(
        PurchaseOrder.id == None
    ).all()
    outstanding_reqs_total = sum((r.value or 0.0) for r in outstanding_reqs)

    all_pos = PurchaseOrder.query.all()
    po_report_data = []
    grand_po_total = 0.0
    grand_inv_total = 0.0
    for po in all_pos:
        inv_total = sum((inv.invoice_amount or 0.0) for inv in po.invoices)
        po_amt = po.po_amount or 0.0
        grand_po_total += po_amt
        grand_inv_total += inv_total
        po_report_data.append({
            'po': po,
            'invoices': po.invoices,
            'po_total': po_amt,
            'invoices_total': inv_total,
            'balance': po_amt - inv_total
        })

    overdue_invoices = get_overdue_invoices(target_15th)
    overdue_invoices_total = sum((item['invoice'].invoice_amount or 0.0) for item in overdue_invoices)

    return render_template(
        'reports.html',
        outstanding_reqs=outstanding_reqs,
        outstanding_reqs_total=outstanding_reqs_total,
        po_report_data=po_report_data,
        grand_po_total=grand_po_total,
        grand_inv_total=grand_inv_total,
        overdue_invoices=overdue_invoices,
        overdue_invoices_total=overdue_invoices_total,
        target_15th=target_15th.strftime('%Y-%m-%d'),
        selected_month=target_15th.strftime('%Y-%m')
    )

@app.route('/reports/export/<report_type>')
def export_report(report_type):
    output = io.StringIO()
    writer = csv.writer(output)

    if report_type == 'requisitions':
        writer.writerow(['Requisition Number', 'Transaction Date', 'Quotation Number', 'POID', 'Vendor Name', 'Description', 'Value', 'Functional Area', 'Remarks', 'Date Sent to Bursary'])
        reqs = Requisition.query.outerjoin(PurchaseOrder).filter(PurchaseOrder.id == None).all()
        for r in reqs:
            writer.writerow([
                r.requisition_number, format_date_us(r.transaction_date), r.quotation_number, r.poid,
                r.vendor_name, r.description, f"{r.value or 0:.2f}", r.functional_area, r.remarks,
                format_date_us(r.date_sent_to_bursary)
            ])
        filename = f"outstanding_requisitions_{date.today()}.csv"

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
        filename = f"purchase_order_report_{date.today()}.csv"

    elif report_type == 'invoices':
        target_month = request.args.get('target_month')
        ref_date = f"{target_month}-15" if target_month else None
        target_15th = get_target_15th(ref_date)
        overdue = get_overdue_invoices(target_15th)

        writer.writerow([
            'Invoice Number', 'Invoice Date', 'Vendor Name', 'Invoice Amount', 'Date Received',
            'Invoice Details', 'POID', 'Remarks', 'Batch ID (Lookup)', 'Date Sent to Nekeisha',
            'Date Returned from Nekeisha', 'Date Sent to TA', 'Date Returned from TA',
            'Date Sent to OM', 'Date Returned from OM', 'Date Submitted to Bursary',
            'Bursary Date Batch', 'Payment Received'
        ])
        for entry in overdue:
            inv = entry['invoice']
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
        filename = f"overdue_invoices_report_{target_15th}.csv"
    else:
        flash("Invalid report type", "danger")
        return redirect(url_for('reports_hub'))

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment;filename={filename}"}
    )

@app.route('/reports/export_excel')
def export_excel_report():
    target_month = request.args.get('target_month')
    ref_date = f"{target_month}-15" if target_month else None
    target_15th = get_target_15th(ref_date)

    wb = openpyxl.Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
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

    # --- SHEET 1: Outstanding Requisitions ---
    ws1 = wb.active
    ws1.title = "Outstanding Requisitions"
    ws1.append(['Requisition Number', 'Transaction Date', 'Quotation Number', 'POID', 'Vendor Name', 'Description', 'Value', 'Functional Area', 'Remarks', 'Date Sent to Bursary'])
    reqs = Requisition.query.outerjoin(PurchaseOrder).filter(PurchaseOrder.id == None).all()
    for r in reqs:
        ws1.append([
            r.requisition_number, format_date_us(r.transaction_date), r.quotation_number, r.poid,
            r.vendor_name, r.description, float(r.value or 0), r.functional_area, r.remarks,
            format_date_us(r.date_sent_to_bursary)
        ])
    style_sheet(ws1, currency_cols=[7], sum_cols=[7])

    # --- SHEET 2: PO vs Invoiced ---
    ws2 = wb.create_sheet(title="PO vs Invoiced")
    ws2.append([
        'POID', 'Requisition Number', 'Vendor Name', 'PO Amount', 'Date', 'Service', 'Functional Area', 'Remarks'
    ])
    for po in PurchaseOrder.query.all():
        ws2.append([
            po.poid,
            po.requisition.requisition_number if po.requisition else 'None',
            po.vendor_name, float(po.po_amount or 0), format_date_us(po.date),
            po.service, po.functional_area, po.remarks
        ])
    style_sheet(ws2, currency_cols=[4], sum_cols=[4])

    # --- SHEET 3: Overdue Invoices ---
    ws3 = wb.create_sheet(title="Overdue Invoices")
    ws3.append([
        'Invoice Number', 'Invoice Date', 'Vendor Name', 'Invoice Amount', 'Date Received',
        'Invoice Details', 'POID', 'Remarks', 'Batch ID (Lookup)', 'Date Sent to Nekeisha',
        'Date Returned from Nekeisha', 'Date Sent to TA', 'Date Returned from TA',
        'Date Sent to OM', 'Date Returned from OM', 'Date Submitted to Bursary',
        'Bursary Date Batch', 'Payment Received'
    ])
    for entry in get_overdue_invoices(target_15th):
        inv = entry['invoice']
        ws3.append([
            inv.invoice_number, format_date_us(inv.invoice_date), inv.vendor_name,
            float(inv.invoice_amount or 0), format_date_us(inv.date_received),
            inv.invoice_details, inv.purchase_order.poid if inv.purchase_order else 'None',
            inv.remarks, inv.batch.batch_number if inv.batch else 'None',
            format_date_us(inv.date_sent_to_nekeisha), format_date_us(inv.date_returned_from_nekeisha),
            format_date_us(inv.date_sent_to_ta), format_date_us(inv.date_returned_from_ta),
            format_date_us(inv.date_sent_to_om), format_date_us(inv.date_returned_from_om),
            format_date_us(inv.date_submitted_to_bursary), inv.bursary_date_batch,
            format_date_us(inv.payment_received)
        ])
    style_sheet(ws3, currency_cols=[4], sum_cols=[4])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Procurement_Bursary_Report_{target_15th}.xlsx"
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
        download_name=f"procurement_backup_{timestamp}.db",
        mimetype="application/octet-stream"
    )

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    ensure_schema_updates()
    with app.app_context():
        db.create_all()
    app.run(debug=True)