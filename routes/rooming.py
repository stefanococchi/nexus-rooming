import os
import io
import re
from datetime import datetime, date

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from flask import (Blueprint, render_template, redirect, url_for,
                   flash, request, send_file, jsonify)

from models import db
from models.models import RoomingList

rooming_bp = Blueprint('rooming', __name__)

# ── Mappa colonne XLS ────────────────────────────────────────────────────────

COL_MAP = {
    0:  'registration_state',
    1:  'latest_changes',
    2:  'hotel',
    3:  'upgrade',
    4:  'participant_display',
    5:  'billing',
    6:  'company_name',
    7:  'company_country',
    8:  'nexus_bd',
    9:  'is_parent_manager',
    10: 'registered_colleagues',
    11: 'internal_reference',
    12: 'internal_parent_reference',
    13: 'ean8_barcode',
    14: 'participant_number',
    15: 'external_reference',
    16: 'delegation_key',
    17: 'status_vp_bd',
    18: 'status_organisator',
    19: 'status_board_nai',
    20: 'status_climate_day',
    21: 'status_prospective_council',
    22: 'status_spouse',
    23: 'comment',
    24: 'title',
    25: 'last_name',
    26: 'first_name',
    27: 'job_position',
    28: 'email',
    29: 'phone',
    30: 'prospective_title',
    31: 'prospective_response',
    32: 'night_no_need',
    33: 'night_sat_28mar',
    34: 'night_sun_29mar',
    35: 'night_mon_30mar',
    36: 'night_tue_31mar',
    37: 'night_wed_1apr',
    38: 'night_thu_2apr',
    39: 'night_fri_3apr',
    40: 'night_sat_4apr',
    41: 'night_other',
    42: 'diet_restrictions',
    43: 'arrival_mode',
    44: 'need_smooth_checkin',
    45: 'need_visa',
    46: 'visa_birth_date',
    47: 'visa_birth_place',
    48: 'visa_passport',
    49: 'visa_delivery_date',
    50: 'visa_expiration_date',
    51: 'visa_company_address',
    52: 'company_category',
    53: 'company_subcategory',
}

DATE_FIELDS = {'visa_birth_date', 'visa_delivery_date', 'visa_expiration_date',
               'change_date', 'file_timestamp', 'check_in', 'check_out'}
INT_FIELDS  = {'registered_colleagues', 'participant_number'}

# Colonne per l'Excel di output hotel (sequenza standard)
EXPORT_COLS = [
    ('title',                'Title'),
    ('last_name',            'Last Name'),
    ('first_name',           'First Name'),
    ('comment',              'Comment'),
    ('email',                'Email'),
    ('phone',                'Phone'),
    ('billing',              'Billing'),
    ('company_category',     'Category'),
    ('upgrade',              'Upgrade'),
    ('arrival_mode',         'Arrival Mode'),
    ('check_in',             'Check In'),
    ('check_out',            'Check Out'),
    ('need_visa',            'Visa?'),
    ('visa_birth_date',      'Birth Date'),
    ('visa_birth_place',     'Birth Place'),
    ('visa_passport',        'Passport N.'),
    ('visa_delivery_date',   'Doc Issued'),
    ('visa_expiration_date', 'Doc Expiry'),
    ('latest_changes',       'Latest Changes'),
    ('change_type',          'Change Type'),
    ('change_date',          'Change Date'),
]

VIRTUAL_FIELDS = {'check_in', 'check_out'}
DIFF_FIELDS = [f for f, _ in EXPORT_COLS if f not in VIRTUAL_FIELDS]

COL_WIDTHS = {
    'title': 6, 'last_name': 18, 'first_name': 16,
    'comment': 35, 'email': 28, 'phone': 16,
    'billing': 14, 'upgrade': 12,
    'check_in': 12, 'check_out': 12,
    'need_visa': 7, 'visa_birth_date': 12,
    'visa_birth_place': 16, 'visa_passport': 14,
    'visa_delivery_date': 12, 'visa_expiration_date': 12,
    'latest_changes': 28, 'change_type': 16, 'change_date': 12,
    # altri usati nell'export dinamico
    'company_name': 28, 'company_country': 12, 'job_position': 30,
    'company_category': 14, 'company_subcategory': 14,
    'registration_state': 10, 'hotel': 20,
    'night_no_need': 8, 'night_sat_28mar': 8, 'night_sun_29mar': 8,
    'night_mon_30mar': 8, 'night_tue_31mar': 8, 'night_wed_1apr': 8,
    'night_thu_2apr': 8, 'night_fri_3apr': 8, 'night_sat_4apr': 8,
    'arrival_mode': 10, 'need_smooth_checkin': 9, 'diet_restrictions': 20,
    'visa_company_address': 30, 'nexus_bd': 20, 'internal_reference': 14,
}


# ── Helpers ──────────────────────────────────────────────────────────────────

def clean_value(field, val):
    if pd.isna(val) if not isinstance(val, (list, dict)) else False:
        return None
    if field in DATE_FIELDS:
        if isinstance(val, (datetime, date)):
            return val if isinstance(val, date) else val.date()
        try:
            return pd.to_datetime(val).date()
        except Exception:
            return None
    if field in INT_FIELDS:
        try:
            return int(val)
        except Exception:
            return None
    s = str(val).strip()
    return s if s and s.lower() not in ('nan', 'none') else None


def make_batch_id(filename):
    ts    = datetime.now().strftime('%Y%m%d_%H%M%S')
    fname = re.sub(r'[^\w\-_. ]', '_', filename)
    return f"{ts}_{fname}"


def get_batches():
    rows = db.session.execute(db.text("""
        SELECT import_batch,
               COUNT(*)        AS total,
               MIN(imported_at) AS imported_at
        FROM rooming_list
        GROUP BY import_batch
        ORDER BY imported_at DESC
    """)).fetchall()
    return rows


def fetch_batch_dict(batch_id):
    rows = RoomingList.query.filter_by(import_batch=batch_id).all()
    result = {}
    for r in rows:
        key = r.internal_reference or f'__noref_{r.last_name}_{r.first_name}'
        result[key] = r
    return result


def compute_diff(current_dict, previous_dict):
    diff = {}
    for ref, row in current_dict.items():
        if ref not in previous_dict:
            diff[ref] = ['NEW']
        else:
            prev = previous_dict[ref]
            changed = [f for f in DIFF_FIELDS
                       if str(getattr(row, f) or '').strip() !=
                          str(getattr(prev, f) or '').strip()]
            if changed:
                diff[ref] = changed
    return diff


from datetime import date as _date_type

NIGHT_DATES = [
    ('night_sat_28mar', _date_type(2026, 3, 28)),
    ('night_sun_29mar', _date_type(2026, 3, 29)),
    ('night_mon_30mar', _date_type(2026, 3, 30)),
    ('night_tue_31mar', _date_type(2026, 3, 31)),
    ('night_wed_1apr',  _date_type(2026, 4,  1)),
    ('night_thu_2apr',  _date_type(2026, 4,  2)),
    ('night_fri_3apr',  _date_type(2026, 4,  3)),
    ('night_sat_4apr',  _date_type(2026, 4,  4)),
]

def get_checkin(row):
    for field, d in NIGHT_DATES:
        if getattr(row, field):
            return d
    return None

def get_checkout(row):
    from datetime import timedelta
    last = None
    for field, d in NIGHT_DATES:
        if getattr(row, field):
            last = d
    return last + timedelta(days=1) if last else None


def fmt_date(val):
    if val is None:
        return ''
    if hasattr(val, 'strftime'):
        return val.strftime('%d/%m/%Y')
    return str(val)


def cell_val(field, row):
    if field == 'check_in':
        return fmt_date(get_checkin(row))
    if field == 'check_out':
        return fmt_date(get_checkout(row))
    val = getattr(row, field, None)
    if val is None:
        return ''
    if field in DATE_FIELDS:
        return fmt_date(val)
    s = str(val).strip()
    return '' if s.lower() in ('none', 'nan') else s


def build_hotel_excel(hotel_name, rows, diff, batch_id, prev_batch_id):
    from models.models import HotelContract
    from datetime import date as _date_type

    wb = openpyxl.Workbook()
    # Rinomina il foglio default che verrà usato come Pivot
    ws_pivot = wb.active
    ws_pivot.title = 'Pivot'

    hdr_font  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', start_color='1F3864')
    tot_font  = Font(name='Calibri', bold=True, size=10)
    tot_fill  = PatternFill('solid', start_color='D9E1F2')
    norm_font_p = Font(name='Calibri', size=10)
    center_p  = Alignment(horizontal='center', vertical='center')
    left_p    = Alignment(horizontal='left',   vertical='center')
    thin_p    = Border(left=Side(style='thin'), right=Side(style='thin'),
                       top=Side(style='thin'),  bottom=Side(style='thin'))

    PIVOT_NIGHTS = [
        ('night_sat_28mar', '28-mar', _date_type(2026, 3, 28)),
        ('night_sun_29mar', '29-mar', _date_type(2026, 3, 29)),
        ('night_mon_30mar', '30-mar', _date_type(2026, 3, 30)),
        ('night_tue_31mar', '31-mar', _date_type(2026, 3, 31)),
        ('night_wed_1apr',  '1-apr',  _date_type(2026, 4,  1)),
        ('night_thu_2apr',  '2-apr',  _date_type(2026, 4,  2)),
        ('night_fri_3apr',  '3-apr',  _date_type(2026, 4,  3)),
        ('night_sat_4apr',  '4-apr',  _date_type(2026, 4,  4)),
    ]

    # Titolo pivot
    ws_pivot.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(PIVOT_NIGHTS)+2)
    tc = ws_pivot.cell(row=1, column=1, value=f'CONFIRMED vs BY CONTRACT — {hotel_name.upper()}')
    tc.font = Font(name='Calibri', bold=True, size=13, color='1F3864')
    tc.alignment = center_p
    ws_pivot.row_dimensions[1].height = 24

    # Header date
    ws_pivot.cell(row=2, column=1, value='').border = thin_p
    ws_pivot.cell(row=2, column=2, value='Voce').border = thin_p
    ws_pivot.cell(row=2, column=2).font = hdr_font
    ws_pivot.cell(row=2, column=2).fill = hdr_fill
    ws_pivot.cell(row=2, column=2).alignment = center_p
    for ci, (_, label, _) in enumerate(PIVOT_NIGHTS, 3):
        c = ws_pivot.cell(row=2, column=ci, value=label)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center_p; c.border = thin_p
    ws_pivot.row_dimensions[2].height = 18

    # Conta confirmed per notte (no CXL)
    active_rows = [r for r in rows if not r.is_cxl]
    confirmed   = [sum(1 for r in active_rows if getattr(r, nf)) for nf, _, _ in PIVOT_NIGHTS]

    # Carica contratti
    contracts   = {str(c.date): c.rooms for c in
                   HotelContract.query.filter_by(hotel=hotel_name).all()}
    contract    = [contracts.get(str(d), 0) for _, _, d in PIVOT_NIGHTS]
    delta       = [confirmed[i] - contract[i] for i in range(len(PIVOT_NIGHTS))]

    rows_data = [
        ('CONFIRMED', confirmed, PatternFill('solid', start_color='DBEAFE')),
        ('BY CONTRACT', contract, PatternFill('solid', start_color='EDE9FE')),
        ('DELTA', delta, None),
    ]

    for ri, (label, values, bg) in enumerate(rows_data, 3):
        ws_pivot.cell(row=ri, column=1, value='').border = thin_p
        lc = ws_pivot.cell(row=ri, column=2, value=label)
        lc.font = tot_font; lc.border = thin_p; lc.alignment = left_p
        if bg:
            lc.fill = bg
        for ci, val in enumerate(values, 3):
            c = ws_pivot.cell(row=ri, column=ci, value=val if val else None)
            c.border = thin_p; c.alignment = center_p
            c.font = tot_font if label != 'DELTA' else Font(name='Calibri', size=10,
                bold=True,
                color='CC0000' if (val < 0) else ('00AA00' if val > 0 else '888888'))
            if bg:
                c.fill = bg
        ws_pivot.row_dimensions[ri].height = 16

    ws_pivot.column_dimensions['A'].width = 2
    ws_pivot.column_dimensions['B'].width = 14
    for ci in range(3, len(PIVOT_NIGHTS) + 3):
        ws_pivot.column_dimensions[get_column_letter(ci)].width = 10

    # ── Foglio 2: Rooming List ────────────────────────────────────────────────
    ws = wb.create_sheet('Rooming List')

    hdr_font  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', start_color='1F3864')
    norm_font = Font(name='Calibri', size=10, color='000000')
    green_fill  = PatternFill('solid', start_color='C6EFCE')
    yellow_fill = PatternFill('solid', start_color='FFEB9C')
    red_fill    = PatternFill('solid', start_color='FFC7CE')
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')
    thin   = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    # Riga 1 — Titolo
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1,   end_column=len(EXPORT_COLS))
    tc = ws.cell(row=1, column=1,
                 value=f'ROOMING LIST — {hotel_name.upper()}')
    tc.font      = Font(name='Calibri', bold=True, size=14, color='1F3864')
    tc.alignment = center
    ws.row_dimensions[1].height = 26

    # Riga 2 — Sottotitolo
    ws.merge_cells(start_row=2, start_column=1,
                   end_row=2,   end_column=len(EXPORT_COLS))
    sub = f"Import: {batch_id[:19].replace('_', ' ')}"
    if prev_batch_id:
        sub += (f"  |  Confronto con: {prev_batch_id[:19].replace('_', ' ')}"
                f"  |  🟡 Modificato   🟢 Nuovo   🔴 CXL")
    ws.cell(row=2, column=1, value=sub).font = Font(name='Calibri', size=9, italic=True)
    ws.row_dimensions[2].height = 14

    # Riga 3 — Intestazioni
    for col, (field, label) in enumerate(EXPORT_COLS, 1):
        c = ws.cell(row=3, column=col, value=label)
        c.font      = hdr_font
        c.fill      = hdr_fill
        c.alignment = center
        c.border    = thin
    ws.row_dimensions[3].height = 20

    def sort_key(r):
        if r.is_cxl:
            group = 2
        elif r.change_date or (r.latest_changes and str(r.latest_changes).strip()):
            group = 1
        else:
            group = 0
        return (group, str(r.last_name or '').upper(), str(r.first_name or '').upper())

    rows_sorted = sorted(rows, key=sort_key)

    for rn, row in enumerate(rows_sorted, 4):
        ref = row.internal_reference or f'__noref_{row.last_name}_{row.first_name}'
        is_cxl     = row.is_cxl
        is_changed = row.change_date or (row.latest_changes and str(row.latest_changes).strip())

        if is_cxl:
            row_fill = red_fill
        elif is_changed:
            row_fill = yellow_fill
        else:
            row_fill = green_fill

        for col, (field, _) in enumerate(EXPORT_COLS, 1):
            c        = ws.cell(row=rn, column=col, value=cell_val(field, row))
            c.border = thin
            c.alignment = left
            c.font = norm_font
            c.fill = row_fill
        ws.row_dimensions[rn].height = 15

    for col, (field, _) in enumerate(EXPORT_COLS, 1):
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS.get(field, 14)

    ws.freeze_panes = 'A4'

    # Riepilogo
    total    = len(rows_sorted)
    n_cxl    = sum(1 for r in rows_sorted if r.is_cxl)
    n_active = total - n_cxl
    n_new    = sum(1 for r in rows_sorted
                   if diff.get(r.internal_reference or
                               f'__noref_{r.last_name}_{r.first_name}') == ['NEW'])
    n_chg    = sum(1 for r in rows_sorted
                   if (r.internal_reference or
                       f'__noref_{r.last_name}_{r.first_name}') in diff
                   and diff[r.internal_reference or
                            f'__noref_{r.last_name}_{r.first_name}'] != ['NEW'])

    sr = total + 5
    ws.cell(row=sr, column=1,
            value=(f'Totale: {total}  |  Attivi: {n_active}  |  CXL: {n_cxl}'
                   f'  |  Nuovi: {n_new}  |  Modificati: {n_chg}')
            ).font = Font(name='Calibri', bold=True, size=10)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, total, n_active, n_cxl


# ── Routes ───────────────────────────────────────────────────────────────────

@rooming_bp.route('/')

def index():
    batches = get_batches()
    return render_template('index.html', batches=batches)


@rooming_bp.route('/upload', methods=['POST'])

def upload():
    import re as _re
    from datetime import date as _date

    f = request.files.get('file')
    if not f or not f.filename.endswith('.xlsx'):
        flash('Seleziona un file .xlsx valido.', 'error')
        return redirect(url_for('rooming.index'))

    # Leggi file_timestamp dai metadati del file Excel
    file_ts = None
    try:
        import openpyxl as _opx
        from io import BytesIO as _BytesIO
        raw_bytes = f.read()
        wb_meta = _opx.load_workbook(_BytesIO(raw_bytes), read_only=True)
        props = wb_meta.properties
        file_ts = props.modified or props.created
        wb_meta.close()
        f.seek(0)
    except Exception:
        f.seek(0)

    try:
        df = pd.read_excel(f, sheet_name='RL', header=0)
    except Exception as e:
        flash(f'Errore lettura file: {e}', 'error')
        return redirect(url_for('rooming.index'))

    def parse_latest_changes(text):
        """Estrae (change_type, change_date) da latest_changes."""
        if not text or str(text).strip().lower() in ('nan', ''):
            return None, None
        t = str(text).strip()
        # Pattern: TIPO - DD.MM.YYYY o TIPO - DD/MM/YYYY
        m = _re.match(r'^([A-Za-z\s]+?)\s*[-–]\s*(\d{2}[./]\d{2}[./]\d{4})', t)
        if not m:
            return 'OTHER', None
        raw_type = m.group(1).strip().upper()
        raw_date = m.group(2).replace('/', '.')
        # Normalizza tipo
        if 'CXL' in raw_type:
            ctype = 'CXL'
        elif 'NEW HOTEL' in raw_type:
            ctype = 'NEW HOTEL'
        elif 'NEW BILLING' in raw_type:
            ctype = 'NEW BILLING'
        elif 'NIGHTS' in raw_type or 'NIGHT' in raw_type:
            ctype = 'NIGHTS CHANGE'
        elif 'NAME' in raw_type:
            ctype = 'NAME CHANGE'
        elif 'NEW' in raw_type:
            ctype = 'NEW'
        elif 'OTHER' in raw_type:
            ctype = 'OTHER'
        else:
            ctype = raw_type[:50]
        # Parsa data
        try:
            parts = raw_date.split('.')
            cdate = _date(int(parts[2]), int(parts[1]), int(parts[0]))
        except Exception:
            cdate = None
        return ctype, cdate

    batch_id = make_batch_id(f.filename)
    inserted = 0
    skipped  = 0

    for _, row in df.iterrows():
        first = clean_value('first_name', row.iloc[26] if len(row) > 26 else None)
        last  = clean_value('last_name',  row.iloc[25] if len(row) > 25 else None)
        if not first and not last:
            skipped += 1
            continue

        obj = RoomingList(import_batch=batch_id, file_timestamp=file_ts)
        for col_idx, field in COL_MAP.items():
            raw = row.iloc[col_idx] if col_idx < len(row) else None
            setattr(obj, field, clean_value(field, raw))

        # Deriva change_type e change_date da latest_changes
        ctype, cdate = parse_latest_changes(obj.latest_changes)
        obj.change_type = ctype
        obj.change_date = cdate

        db.session.add(obj)
        inserted += 1

    db.session.commit()
    flash(f'Import completato: {inserted} partecipanti caricati (batch: {batch_id[:19]}).', 'success')
    return redirect(url_for('rooming.index'))


@rooming_bp.route('/batch/<path:batch_id>')

def batch_detail(batch_id):
    rows = RoomingList.query.filter_by(import_batch=batch_id)\
                            .order_by(RoomingList.hotel, RoomingList.last_name)\
                            .all()
    if not rows:
        flash('Batch non trovato.', 'error')
        return redirect(url_for('rooming.index'))

    hotels = {}
    for r in rows:
        h = r.hotel or 'SENZA HOTEL'
        hotels.setdefault(h, []).append(r)

    # Raggruppa modifiche per hotel: {hotel: [(change_type, change_date, count), ...]}
    hotel_changes = {}
    for h, hrows in hotels.items():
        counts = {}
        for r in hrows:
            if r.change_type:
                date_str = r.change_date.strftime('%d/%m') if r.change_date else '?'
                key = (r.change_type, date_str)
                counts[key] = counts.get(key, 0) + 1
        # Ordina per data
        hotel_changes[h] = sorted(
            [(ct, cd, n) for (ct, cd), n in counts.items()],
            key=lambda x: x[1]
        )

    # Batch precedente per diff
    all_batches = get_batches()
    batch_ids   = [b[0] for b in all_batches]
    try:
        idx          = batch_ids.index(batch_id)
        prev_batch   = batch_ids[idx + 1] if idx + 1 < len(batch_ids) else None
    except ValueError:
        prev_batch   = None

    diff = {}
    if prev_batch:
        current_dict  = fetch_batch_dict(batch_id)
        previous_dict = fetch_batch_dict(prev_batch)
        diff          = compute_diff(current_dict, previous_dict)

    return render_template('batch_detail.html',
                           batch_id=batch_id,
                           hotels=hotels,
                           hotel_changes=hotel_changes,
                           diff=diff,
                           prev_batch=prev_batch)


@rooming_bp.route('/export/<path:batch_id>/<hotel_name>')

def export_hotel(batch_id, hotel_name):
    rows = RoomingList.query.filter_by(import_batch=batch_id, hotel=hotel_name).all()
    if not rows:
        flash('Nessun partecipante trovato.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    all_batches = get_batches()
    batch_ids   = [b[0] for b in all_batches]
    try:
        idx        = batch_ids.index(batch_id)
        prev_batch = batch_ids[idx + 1] if idx + 1 < len(batch_ids) else None
    except ValueError:
        prev_batch = None

    diff = {}
    if prev_batch:
        diff = compute_diff(fetch_batch_dict(batch_id), fetch_batch_dict(prev_batch))

    buf, total, active, cxl = build_hotel_excel(hotel_name, rows, diff, batch_id, prev_batch)
    safe = hotel_name.replace('/', '-').replace(' ', '_')
    ts   = datetime.now().strftime('%Y%m%d')
    return send_file(buf,
                     as_attachment=True,
                     download_name=f'RoomingList_{safe}_{ts}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@rooming_bp.route('/export-all/<path:batch_id>')

def export_all(batch_id):
    """Scarica un ZIP con un Excel per ogni hotel."""
    import zipfile

    all_batches = get_batches()
    batch_ids   = [b[0] for b in all_batches]
    try:
        idx        = batch_ids.index(batch_id)
        prev_batch = batch_ids[idx + 1] if idx + 1 < len(batch_ids) else None
    except ValueError:
        prev_batch = None

    diff = {}
    if prev_batch:
        diff = compute_diff(fetch_batch_dict(batch_id), fetch_batch_dict(prev_batch))

    rows  = RoomingList.query.filter_by(import_batch=batch_id).all()
    hotels = {}
    for r in rows:
        h = r.hotel or 'SENZA_HOTEL'
        hotels.setdefault(h, []).append(r)

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for hotel, hrows in sorted(hotels.items()):
            buf, _, _, _ = build_hotel_excel(hotel, hrows, diff, batch_id, prev_batch)
            safe = hotel.replace('/', '-').replace(' ', '_')
            ts   = datetime.now().strftime('%Y%m%d')
            zf.writestr(f'RoomingList_{safe}_{ts}.xlsx', buf.read())

    zip_buf.seek(0)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(zip_buf,
                     as_attachment=True,
                     download_name=f'RoomingList_ALL_{ts}.zip',
                     mimetype='application/zip')


@rooming_bp.route('/report-category/<path:batch_id>', methods=['POST'])
def report_category(batch_id):
    """Excel con partecipanti filtrati per company_category, raggruppati per hotel."""
    from sqlalchemy import or_
    categories = request.form.getlist('categories')
    if not categories:
        flash('Seleziona almeno una categoria.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    # Gestisce blank separatamente
    conditions = []
    for cat in categories:
        if cat == '__blank__':
            conditions.append(RoomingList.company_category == None)
        else:
            conditions.append(RoomingList.company_category == cat)

    rows = RoomingList.query.filter_by(import_batch=batch_id)\
                            .filter(or_(*conditions))\
                            .order_by(RoomingList.hotel, RoomingList.last_name,
                                      RoomingList.first_name).all()

    # Ordina per gruppo: invariati, modificati, CXL — dentro ogni hotel
    def _cat_sort(r):
        group = 1 if r.is_cxl else 0
        return (r.hotel or '', group,
                str(r.last_name or '').upper(),
                str(r.first_name or '').upper())

    rows = sorted(rows, key=_cat_sort)

    if not rows:
        flash('Nessun partecipante trovato per le categorie selezionate.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    cat_labels = [c if c != '__blank__' else '(blank)' for c in categories]
    title_str  = ' · '.join(cat_labels)

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = 'Report'

    hdr_font  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', start_color='1F3864')
    norm_font = Font(name='Calibri', size=10)
    green_fill  = PatternFill('solid', start_color='C6EFCE')
    yellow_fill = PatternFill('solid', start_color='FFEB9C')
    red_fill    = PatternFill('solid', start_color='FFC7CE')
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')
    thin   = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    n_cols = len(EXPORT_COLS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value=f'REPORT — {title_str.upper()}')
    tc.font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    tc.alignment = center
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ts_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    ws.cell(row=2, column=1,
            value=f'Generato il {ts_str}  |  {len(rows)} partecipanti'
            ).font = Font(name='Calibri', size=9, italic=True)
    ws.row_dimensions[2].height = 14

    current_hotel = None
    row_num = 3

    for r in rows:
        hotel = r.hotel or 'SENZA HOTEL'
        if hotel != current_hotel:
            current_hotel = hotel
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=n_cols)
            hc = ws.cell(row=row_num, column=1, value=hotel.upper())
            hc.font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
            hc.fill = PatternFill('solid', start_color='2F5496')
            hc.alignment = left
            ws.row_dimensions[row_num].height = 20
            row_num += 1
            for col, (field, label) in enumerate(EXPORT_COLS, 1):
                c = ws.cell(row=row_num, column=col, value=label)
                c.font = hdr_font; c.fill = hdr_fill
                c.alignment = center; c.border = thin
            ws.row_dimensions[row_num].height = 18
            row_num += 1

        is_cxl    = r.is_cxl
        is_changed = r.change_date or (r.latest_changes and str(r.latest_changes).strip())
        row_fill = red_fill if is_cxl else (yellow_fill if is_changed else green_fill)

        for col, (field, _) in enumerate(EXPORT_COLS, 1):
            c = ws.cell(row=row_num, column=col, value=cell_val(field, r))
            c.font = norm_font; c.fill = row_fill
            c.alignment = left; c.border = thin
        ws.row_dimensions[row_num].height = 15
        row_num += 1

    for col, (field, _) in enumerate(EXPORT_COLS, 1):
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS.get(field, 14)
    ws.freeze_panes = 'A3'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe = title_str.replace(' ', '_').replace('/', '-')[:30]
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(buf, as_attachment=True,
                     download_name=f'Report_{safe}_{ts}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@rooming_bp.route('/api/pivot/<path:batch_id>')
def api_pivot(batch_id):
    """Ritorna JSON con pivot notti x hotel — esclude solo CXL."""
    rows = RoomingList.query.filter_by(import_batch=batch_id)\
                            .filter(RoomingList.registration_state != 'CXL')\
                            .all()
    if not rows:
        return jsonify({'error': 'Batch non trovato'}), 404

    NIGHTS = [
        ('night_no_need',    'No need'),
        ('night_sat_28mar',  'Sat 28/03'),
        ('night_sun_29mar',  'Sun 29/03'),
        ('night_mon_30mar',  'Mon 30/03'),
        ('night_tue_31mar',  'Tue 31/03 (Nexus)'),
        ('night_wed_1apr',   'Wed 01/04 (Nexus)'),
        ('night_thu_2apr',   'Thu 02/04'),
        ('night_fri_3apr',   'Fri 03/04'),
        ('night_sat_4apr',   'Sat 04/04'),
        ('night_other',      'Other…'),
    ]

    hotels = sorted(set(r.hotel or 'None' for r in rows))
    pivot  = {nf: {h: 0 for h in hotels} for nf, _ in NIGHTS}

    for r in rows:
        h = r.hotel or 'None'
        for nf, _ in NIGHTS:
            if getattr(r, nf):
                pivot[nf][h] += 1

    return jsonify({
        'hotels': hotels,
        'nights': [{'field': nf, 'label': nl} for nf, nl in NIGHTS],
        'pivot':  {nf: pivot[nf] for nf, _ in NIGHTS},
    })



@rooming_bp.route('/export-pivot/<path:batch_id>')
def export_pivot(batch_id):
    rows = RoomingList.query.filter_by(import_batch=batch_id)\
                            .filter(RoomingList.registration_state != 'CXL')\
                            .all()
    if not rows:
        flash('Batch non trovato.', 'error')
        return redirect(url_for('rooming.index'))

    # Definizione notti in ordine
    NIGHTS = [
        ('night_no_need',    'No need'),
        ('night_sat_28mar',  'Night of Saturday, March 28th'),
        ('night_sun_29mar',  'Night of Sunday, March 29th'),
        ('night_mon_30mar',  'Night of Monday, March 30th'),
        ('night_tue_31mar',  'Night of Tuesday, March 31st (covered by Nexus)'),
        ('night_wed_1apr',   'Night of Wednesday, April 1st (covered by Nexus)'),
        ('night_thu_2apr',   'Night of Thursday, April 2nd'),
        ('night_fri_3apr',   'Night of Friday, April 3rd'),
        ('night_sat_4apr',   'Night of Saturday, April 4th'),
        ('night_other',      'Other…'),
    ]

    # Hotel in ordine alfabetico
    hotels = sorted(set(r.hotel or 'None' for r in rows))

    # Costruisce matrice pivot[notte][hotel] = count
    pivot = {nf: {h: 0 for h in hotels} for nf, _ in NIGHTS}
    for r in rows:
        h = r.hotel or 'None'
        for nf, _ in NIGHTS:
            if getattr(r, nf):
                pivot[nf][h] += 1

    # ── Excel ─────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Pivot Notti x Hotel'

    hdr_font   = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill   = PatternFill('solid', start_color='1F3864')
    tot_font   = Font(name='Calibri', bold=True, size=10)
    tot_fill   = PatternFill('solid', start_color='D9E1F2')
    norm_font  = Font(name='Calibri', size=10)
    zero_font  = Font(name='Calibri', size=10, color='CCCCCC')
    center     = Alignment(horizontal='center', vertical='center')
    left       = Alignment(horizontal='left',   vertical='center')
    thin       = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'),  bottom=Side(style='thin'))

    # Riga 1 — titolo
    n_cols = len(hotels) + 2  # notte + hotels + grand total
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value='PIVOT — NOTTI PER HOTEL')
    tc.font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    tc.alignment = center
    ws.row_dimensions[1].height = 26

    # Riga 2 — sottotitolo
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ts_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    ws.cell(row=2, column=1,
            value=f'Generato il {ts_str}  |  batch: {batch_id[:19]}'
            ).font = Font(name='Calibri', size=9, italic=True)
    ws.row_dimensions[2].height = 14

    # Riga 3 — intestazioni colonne
    ws.cell(row=3, column=1, value='Notte').font = hdr_font
    ws.cell(row=3, column=1).fill = hdr_fill
    ws.cell(row=3, column=1).alignment = left
    ws.cell(row=3, column=1).border = thin

    for ci, h in enumerate(hotels, 2):
        c = ws.cell(row=3, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center; c.border = thin

    # Colonna Grand Total
    gt_col = len(hotels) + 2
    c = ws.cell(row=3, column=gt_col, value='Grand Total')
    c.font = hdr_font; c.fill = hdr_fill
    c.alignment = center; c.border = thin
    ws.row_dimensions[3].height = 20

    # Righe dati
    for rn, (nf, label) in enumerate(NIGHTS, 4):
        ws.cell(row=rn, column=1, value=label).font = norm_font
        ws.cell(row=rn, column=1).border = thin
        ws.cell(row=rn, column=1).alignment = left

        row_total = 0
        for ci, h in enumerate(hotels, 2):
            val = pivot[nf][h]
            row_total += val
            c = ws.cell(row=rn, column=ci, value=val if val else None)
            c.border = thin; c.alignment = center
            c.font = norm_font if val else zero_font

        # Grand Total riga
        c = ws.cell(row=rn, column=gt_col, value=row_total if row_total else None)
        c.font = tot_font; c.fill = tot_fill
        c.border = thin; c.alignment = center
        ws.row_dimensions[rn].height = 16

    # Riga Grand Total colonne
    tot_row = len(NIGHTS) + 4
    ws.cell(row=tot_row, column=1, value='Grand Total').font = tot_font
    ws.cell(row=tot_row, column=1).fill = tot_fill
    ws.cell(row=tot_row, column=1).border = thin

    grand_total = 0
    for ci, h in enumerate(hotels, 2):
        col_total = sum(pivot[nf][h] for nf, _ in NIGHTS)
        grand_total += col_total
        c = ws.cell(row=tot_row, column=ci, value=col_total if col_total else None)
        c.font = tot_font; c.fill = tot_fill
        c.border = thin; c.alignment = center

    c = ws.cell(row=tot_row, column=gt_col, value=grand_total)
    c.font = Font(name='Calibri', bold=True, size=11, color='1F3864')
    c.fill = PatternFill('solid', start_color='B8CCE4')
    c.border = thin; c.alignment = center
    ws.row_dimensions[tot_row].height = 18

    # Larghezze colonne
    ws.column_dimensions['A'].width = 44
    for ci in range(2, gt_col + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 18

    ws.freeze_panes = 'B4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(buf, as_attachment=True,
                     download_name=f'Pivot_Notti_{ts}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')




@rooming_bp.route('/batch/delete/<path:batch_id>', methods=['POST'])
def delete_batch(batch_id):
    RoomingList.query.filter_by(import_batch=batch_id).delete()
    db.session.commit()
    flash('Batch eliminato.', 'success')
    return redirect(url_for('rooming.index'))


# ── Export dinamico ──────────────────────────────────────────────────────────

# Tutte le colonne disponibili con etichetta
ALL_COLS = [
    ('title',                'Title'),
    ('last_name',            'Last Name'),
    ('first_name',           'First Name'),
    ('comment',              'Comment'),
    ('email',                'Email'),
    ('phone',                'Phone'),
    ('billing',              'Billing'),
    ('upgrade',              'Upgrade'),
    ('check_in',             'Check In'),
    ('check_out',            'Check Out'),
    ('company_name',         'Company'),
    ('company_country',      'Country'),
    ('job_position',         'Job Position'),
    ('email',                'Email'),
    ('phone',                'Phone'),
    ('company_category',     'Category'),
    ('company_subcategory',  'Sub-Category'),
    ('registration_state',   'Reg. State'),
    ('hotel',                'Hotel'),
    ('upgrade',              'Upgrade'),
    ('billing',              'Billing'),
    ('nexus_bd',             'NEXUS BD'),
    ('night_no_need',        'No need'),
    ('night_sat_28mar',      'Sat 28/03'),
    ('night_sun_29mar',      'Sun 29/03'),
    ('night_mon_30mar',      'Mon 30/03'),
    ('night_tue_31mar',      'Tue 31/03'),
    ('night_wed_1apr',       'Wed 01/04'),
    ('night_thu_2apr',       'Thu 02/04'),
    ('night_fri_3apr',       'Fri 03/04'),
    ('night_sat_4apr',       'Sat 04/04'),
    ('arrival_mode',         'Arrival'),
    ('need_smooth_checkin',  'Smooth CI'),
    ('diet_restrictions',    'Diet'),
    ('need_visa',            'Visa?'),
    ('visa_birth_date',      'Birth Date'),
    ('visa_birth_place',     'Birth Place'),
    ('visa_passport',        'Passport N.'),
    ('visa_delivery_date',   'Doc Issued'),
    ('visa_expiration_date', 'Doc Expiry'),
    ('visa_company_address', 'Company Address'),
    ('comment',              'Comment'),
    ('internal_reference',   'Internal Ref'),
    ('participant_number',   'Participant N.'),
    ('ean8_barcode',         'EAN8'),
    ('status_vp_bd',         'VP/BD'),
    ('status_organisator',   'Organisator'),
    ('status_board_nai',     'Board NAI'),
    ('status_climate_day',   'Climate Day'),
    ('status_spouse',        'Spouse Flag'),
    ('spouse_name',          'Spouse Name'),   # colonna virtuale — valorizzata solo se elaborate_spouse
    ('is_parent_manager',    'Parent Mgr'),
    ('registered_colleagues','Colleagues'),
    ('latest_changes',       'Latest Changes'),
    ('change_type',          'Change Type'),
    ('change_date',          'Change Date'),
    ('file_timestamp',       'File Timestamp'),
]

@rooming_bp.route('/export-custom/<path:batch_id>')
def export_custom(batch_id):
    from models.models import ExportConfig
    rows = RoomingList.query.filter_by(import_batch=batch_id).all()
    if not rows:
        flash('Batch non trovato.', 'error')
        return redirect(url_for('rooming.index'))

    hotels       = sorted(set(r.hotel or 'SENZA HOTEL' for r in rows))
    stati        = sorted(set(r.registration_state or '' for r in rows if r.registration_state))
    saved_configs = [c.name for c in ExportConfig.query.order_by(ExportConfig.name).all()]

    return render_template('export_custom.html',
                           batch_id=batch_id,
                           all_cols=ALL_COLS,
                           hotels=hotels,
                           stati=stati,
                           saved_configs=saved_configs)


@rooming_bp.route('/api/preview/<path:batch_id>', methods=['POST'])
def api_preview(batch_id):
    """Ritorna JSON con anteprima righe filtrate e colonne scelte."""
    data        = request.get_json()
    cols        = data.get('cols', [])          # lista di field names in ordine
    f_hotels    = data.get('hotels', [])        # [] = tutti
    f_stati     = data.get('stati', [])         # [] = tutti
    f_notti     = data.get('notti', [])         # [] = tutte
    include_cxl = data.get('include_cxl', True)
    page        = int(data.get('page', 1))
    per_page    = 50

    if not cols:
        return jsonify({'rows': [], 'total': 0, 'cols': []})

    q = RoomingList.query.filter_by(import_batch=batch_id)
    if f_hotels:
        q = q.filter(RoomingList.hotel.in_(f_hotels))
    if f_stati:
        q = q.filter(RoomingList.registration_state.in_(f_stati))
    if not include_cxl:
        q = q.filter(RoomingList.registration_state != 'CXL')

    all_rows = q.order_by(RoomingList.last_name, RoomingList.first_name).all()

    # Filtro notti (AND: deve avere almeno una delle notti selezionate)
    if f_notti:
        night_map = {
            'sat_28': 'night_sat_28mar', 'sun_29': 'night_sun_29mar',
            'mon_30': 'night_mon_30mar', 'tue_31': 'night_tue_31mar',
            'wed_1':  'night_wed_1apr',  'thu_2':  'night_thu_2apr',
            'fri_3':  'night_fri_3apr',  'sat_4':  'night_sat_4apr',
        }
        filtered = []
        for r in all_rows:
            for n in f_notti:
                field = night_map.get(n)
                if field and getattr(r, field):
                    filtered.append(r)
                    break
        all_rows = filtered

    total  = len(all_rows)
    offset = (page - 1) * per_page
    paged  = all_rows[offset:offset + per_page]

    # Costruisce label mappa
    col_labels = {f: l for f, l in ALL_COLS}

    result_rows = []
    for r in paged:
        row_data = {}
        for f in cols:
            row_data[f] = cell_val(f, r)
        result_rows.append(row_data)

    return jsonify({
        'rows':   result_rows,
        'total':  total,
        'page':   page,
        'pages':  (total + per_page - 1) // per_page,
        'labels': {f: col_labels.get(f, f) for f in cols},
    })


@rooming_bp.route('/api/export-custom/<path:batch_id>', methods=['POST'])
def api_export_custom(batch_id):
    """Genera Excel con configurazione dinamica."""
    data              = request.get_json()
    cols              = data.get('cols', [])
    f_hotels          = data.get('hotels', [])
    f_stati           = data.get('stati', [])
    f_notti           = data.get('notti', [])
    include_cxl       = data.get('include_cxl', True)
    title_name        = data.get('title', 'Export')
    elaborate_spouse  = data.get('elaborate_spouse', False)

    if not cols:
        return jsonify({'error': 'Nessuna colonna selezionata'}), 400

    q = RoomingList.query.filter_by(import_batch=batch_id)
    if f_hotels:
        q = q.filter(RoomingList.hotel.in_(f_hotels))
    if f_stati:
        q = q.filter(RoomingList.registration_state.in_(f_stati))
    if not include_cxl:
        q = q.filter(RoomingList.registration_state != 'CXL')

    all_rows = q.order_by(RoomingList.last_name, RoomingList.first_name).all()

    if f_notti:
        night_map = {
            'sat_28': 'night_sat_28mar', 'sun_29': 'night_sun_29mar',
            'mon_30': 'night_mon_30mar', 'tue_31': 'night_tue_31mar',
            'wed_1':  'night_wed_1apr',  'thu_2':  'night_thu_2apr',
            'fri_3':  'night_fri_3apr',  'sat_4':  'night_sat_4apr',
        }
        filtered = []
        for r in all_rows:
            for n in f_notti:
                field = night_map.get(n)
                if field and getattr(r, field):
                    filtered.append(r)
                    break
        all_rows = filtered

    # Ordinamento: prima invariati, poi modificati, in fondo CXL
    def _sort_key(r):
        if r.is_cxl:
            group = 2
        elif r.change_date or (r.latest_changes and str(r.latest_changes).strip()):
            group = 1
        else:
            group = 0
        return (group, str(r.last_name or '').upper(), str(r.first_name or '').upper())

    all_rows = sorted(all_rows, key=_sort_key)

    # ── Elaborazione spouse runtime ───────────────────────────────────────
    spouse_notes  = {}   # id riga spouse -> "Spouse of <nome>"
    spouse_ids    = set()  # id righe che sono spouse
    parent_spouse = {}   # id riga titolare -> nome spouse

    if elaborate_spouse:
        all_batch = RoomingList.query.filter_by(import_batch=batch_id).all()
        ref_to_row = {
            r.internal_reference: r
            for r in all_batch if r.internal_reference
        }
        import re as _re
        for r in all_batch:
            if (r.status_spouse or '').strip().lower() == 'yes':
                spouse_ids.add(r.id)
                spouse_full = f"{r.first_name or ''} {r.last_name or ''}".strip()
                parent_name = None
                parent_row  = None

                # 1. Cerca via internal_parent_reference
                if r.internal_parent_reference:
                    parent_row = ref_to_row.get(r.internal_parent_reference)
                    if parent_row:
                        parent_name = f"{parent_row.first_name or ''} {parent_row.last_name or ''}".strip()

                # 2. Fallback: cerca nel comment pattern noti
                if not parent_name and r.comment:
                    m = _re.search(
                        r'(?:wife|spouse)\s+of\s+(.+)|'
                        r'double\s+room\s+with\s+(.+)|'
                        r'shared\s+room\s+with\s+(.+)|'
                        r'twin\s+room\s+with\s+(.+)',
                        r.comment, _re.IGNORECASE)
                    if m:
                        parent_name = next(g for g in m.groups() if g).strip().rstrip('.')

                # Nota sulla riga spouse
                if parent_name:
                    spouse_notes[r.id] = f"Spouse of {parent_name}"

                # Nota sulla riga titolare
                if parent_row and spouse_full:
                    parent_spouse[parent_row.id] = spouse_full

    col_labels = {f: l for f, l in ALL_COLS}

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = 'Export'

    hdr_font    = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill    = PatternFill('solid', start_color='1F3864')
    norm_font   = Font(name='Calibri', size=10)
    cxl_font    = Font(name='Calibri', size=10, color='CC0000', strike=True)
    spouse_font = Font(name='Calibri', size=10, color='888888', strike=True)
    left        = Alignment(horizontal='left', vertical='center')
    center      = Alignment(horizontal='center', vertical='center')
    thin        = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'),  bottom=Side(style='thin'))

    # Titolo
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    tc = ws.cell(row=1, column=1, value=title_name.upper())
    tc.font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    tc.alignment = center
    ws.row_dimensions[1].height = 26

    # Sottotitolo
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(cols))
    ts_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    ws.cell(row=2, column=1,
            value=f'Generato il {ts_str}  |  {len(all_rows)} partecipanti'
            ).font = Font(name='Calibri', size=9, italic=True)
    ws.row_dimensions[2].height = 14

    # Header
    for ci, f in enumerate(cols, 1):
        c = ws.cell(row=3, column=ci, value=col_labels.get(f, f))
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center; c.border = thin
    ws.row_dimensions[3].height = 20

    for rn, row in enumerate(all_rows, 4):
        is_cxl    = row.is_cxl
        is_spouse = row.id in spouse_ids

        for ci, f in enumerate(cols, 1):
            # Colonne virtuali elaborate a runtime
            if f == 'spouse_name':
                val = parent_spouse.get(row.id, '')
            elif is_spouse and f == 'latest_changes' and row.id in spouse_notes:
                val = spouse_notes[row.id]
            else:
                val = cell_val(f, row)

            c = ws.cell(row=rn, column=ci, value=val)
            c.border = thin; c.alignment = left

            if is_cxl:
                c.font = cxl_font
            elif is_spouse:
                # Latest Changes leggibile (non barrato), resto grigio barrato
                if f == 'latest_changes' and row.id in spouse_notes:
                    c.font = Font(name='Calibri', size=10, bold=True, color='1F3864')
                else:
                    c.font = spouse_font
            else:
                c.font = norm_font
        ws.row_dimensions[rn].height = 15

    for ci, f in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(f, 16)

    ws.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(buf, as_attachment=True,
                     download_name=f'Export_{ts}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')



# ── Preview report API (unified) ─────────────────────────────────────────────

PREVIEW_COLS = [
    # Sequenza standard (uguale a EXPORT_COLS)
    ('title',                     'Title'),
    ('last_name',                 'Last Name'),
    ('first_name',                'First Name'),
    ('comment',                   'Comment'),
    ('email',                     'Email'),
    ('phone',                     'Phone'),
    ('billing',                   'Billing'),
    ('upgrade',                   'Upgrade'),
    ('arrival_mode',              'Arrival Mode'),
    ('check_in',                  'Check In'),
    ('check_out',                 'Check Out'),
    ('need_visa',                 'Visa?'),
    ('visa_birth_date',           'Birth Date'),
    ('visa_birth_place',          'Birth Place'),
    ('visa_passport',             'Passport N.'),
    ('visa_delivery_date',        'Doc Issued'),
    ('visa_expiration_date',      'Doc Expiry'),
    ('latest_changes',            'Latest Changes'),
    ('change_type',               'Change Type'),
    ('change_date',               'Change Date'),
    # Campi aggiuntivi in coda
    ('registration_state',        'Stato'),
    ('hotel',                     'Hotel'),
    ('company_name',              'Azienda'),
    ('company_country',           'Paese'),
    ('company_category',          'Categoria'),
    ('company_subcategory',       'Sub-Categoria'),
    ('job_position',              'Job Position'),
    ('nexus_bd',                  'NEXUS BD'),
    ('need_smooth_checkin',       'Smooth CI'),
    ('diet_restrictions',         'Dieta'),
    ('internal_reference',        'Internal Ref'),
    ('participant_number',        'N. Partecipante'),
    ('visa_company_address',      'Indirizzo azienda'),
    ('status_vp_bd',              'VP/BD'),
    ('status_organisator',        'Organisator'),
    ('status_board_nai',          'Board NAI'),
    ('status_climate_day',        'Climate Day'),
    ('status_prospective_council','Prospective Council'),
    ('status_spouse',             'Spouse'),
    ('night_sat_28mar',           'Notte 28/03'),
    ('night_sun_29mar',           'Notte 29/03'),
    ('night_mon_30mar',           'Notte 30/03'),
    ('night_tue_31mar',           'Notte 31/03'),
    ('night_wed_1apr',            'Notte 01/04'),
    ('night_thu_2apr',            'Notte 02/04'),
    ('night_fri_3apr',            'Notte 03/04'),
    ('night_sat_4apr',            'Notte 04/04'),
    ('delegation_key',            'Delegation Key'),
    ('internal_parent_reference', 'Parent Ref'),
    ('prospective_response',      'Prospective Response'),
]

@rooming_bp.route('/api/preview-report/<path:batch_id>', methods=['POST'])
def api_preview_report(batch_id):
    from sqlalchemy import or_, and_
    data     = request.get_json()
    rtype    = data.get('type')
    values   = data.get('values', [])
    conditions_raw = data.get('conditions', [])
    logic    = data.get('logic', 'AND')

    q = RoomingList.query.filter_by(import_batch=batch_id)

    if rtype == 'category':
        conds = [RoomingList.company_category == None if v == '__blank__'
                 else RoomingList.company_category == v for v in values]
        q = q.filter(or_(*conds))
    elif rtype == 'billing':
        conds = [RoomingList.billing == None if v == '__blank__'
                 else RoomingList.billing == v for v in values]
        q = q.filter(or_(*conds))
    elif rtype == 'arrival':
        conds = [RoomingList.arrival_mode == None if v == '__blank__'
                 else RoomingList.arrival_mode == v for v in values]
        q = q.filter(or_(*conds))
    elif rtype == 'query':
        conds = []
        for c in conditions_raw:
            field = c.get('field', '')
            op    = c.get('op', 'eq')
            val   = c.get('val', '')
            col   = getattr(RoomingList, field, None)
            if col is None:
                continue
            if op == 'eq':
                conds.append(col == val)
            elif op == 'neq':
                conds.append(col != val)
            elif op == 'contains':
                conds.append(col.ilike(f'%{val}%'))
            elif op == 'not_empty':
                conds.append(col != None)
            elif op == 'empty':
                conds.append(col == None)
        if conds:
            q = q.filter(and_(*conds) if logic == 'AND' else or_(*conds))

    rows = q.order_by(RoomingList.hotel, RoomingList.last_name,
                      RoomingList.first_name).all()

    def _sort(r):
        return (r.hotel or '', 1 if r.is_cxl else 0,
                str(r.last_name or '').upper())
    rows = sorted(rows, key=_sort)

    result = []
    for r in rows:
        row_data = {}
        for field, label in PREVIEW_COLS:
            row_data[label] = cell_val(field, r)
        result.append(row_data)

    return jsonify({'rows': result, 'total': len(result)})


@rooming_bp.route('/report-query/<path:batch_id>', methods=['POST'])
def report_query(batch_id):
    """Excel da query libera."""
    from sqlalchemy import or_, and_
    import json as _json
    conditions_raw = _json.loads(request.form.get('conditions', '[]'))
    logic          = request.form.get('logic', 'AND')

    q = RoomingList.query.filter_by(import_batch=batch_id)
    conds = []
    for c in conditions_raw:
        field = c.get('field', '')
        op    = c.get('op', 'eq')
        val   = c.get('val', '')
        col   = getattr(RoomingList, field, None)
        if col is None:
            continue
        if op == 'eq':       conds.append(col == val)
        elif op == 'neq':    conds.append(col != val)
        elif op == 'contains': conds.append(col.ilike(f'%{val}%'))
        elif op == 'not_empty': conds.append(col != None)
        elif op == 'empty':  conds.append(col == None)
    if conds:
        q = q.filter(and_(*conds) if logic == 'AND' else or_(*conds))

    rows = q.order_by(RoomingList.hotel, RoomingList.last_name,
                      RoomingList.first_name).all()

    def _sort(r):
        return (r.hotel or '', 1 if r.is_cxl else 0,
                str(r.last_name or '').upper(), str(r.first_name or '').upper())
    rows = sorted(rows, key=_sort)

    if not rows:
        flash('Nessun risultato.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = 'Query Result'

    hdr_font  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', start_color='1F3864')
    norm_font = Font(name='Calibri', size=10)
    green_fill  = PatternFill('solid', start_color='C6EFCE')
    yellow_fill = PatternFill('solid', start_color='FFEB9C')
    red_fill    = PatternFill('solid', start_color='FFC7CE')
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')
    thin   = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    n_cols = len(PREVIEW_COLS)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value='QUERY RESULT')
    tc.font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    tc.alignment = center
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ts_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    ws.cell(row=2, column=1,
            value=f'Generato il {ts_str}  |  {len(rows)} partecipanti'
            ).font = Font(name='Calibri', size=9, italic=True)
    ws.row_dimensions[2].height = 14

    current_hotel = None
    row_num = 3
    for r in rows:
        hotel = r.hotel or 'SENZA HOTEL'
        if hotel != current_hotel:
            current_hotel = hotel
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=n_cols)
            hc = ws.cell(row=row_num, column=1, value=hotel.upper())
            hc.font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
            hc.fill = PatternFill('solid', start_color='2F5496')
            hc.alignment = left
            ws.row_dimensions[row_num].height = 20
            row_num += 1
            for col, (field, label) in enumerate(PREVIEW_COLS, 1):
                c = ws.cell(row=row_num, column=col, value=label)
                c.font = hdr_font; c.fill = hdr_fill
                c.alignment = center; c.border = thin
            ws.row_dimensions[row_num].height = 18
            row_num += 1

        is_changed = r.change_date or (r.latest_changes and str(r.latest_changes).strip())
        row_fill = red_fill if r.is_cxl else (yellow_fill if is_changed else green_fill)
        for col, (field, _) in enumerate(PREVIEW_COLS, 1):
            c = ws.cell(row=row_num, column=col, value=cell_val(field, r))
            c.font = norm_font; c.fill = row_fill
            c.alignment = left; c.border = thin
        ws.row_dimensions[row_num].height = 15
        row_num += 1

    for col, (field, _) in enumerate(PREVIEW_COLS, 1):
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS.get(field, 14)
    ws.freeze_panes = 'A3'

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(buf, as_attachment=True,
                     download_name=f'Query_{ts}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@rooming_bp.route('/api/save-config', methods=['POST'])
def api_save_config():
    import json
    from models.models import ExportConfig
    data = request.get_json()
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': 'Nome configurazione mancante'}), 400
    cfg = ExportConfig.query.filter_by(name=name).first()
    if not cfg:
        cfg = ExportConfig(name=name)
        db.session.add(cfg)
    cfg.cols        = json.dumps(data.get('cols', []))
    cfg.hotels      = json.dumps(data.get('hotels', []))
    cfg.stati       = json.dumps(data.get('stati', []))
    cfg.notti       = json.dumps(data.get('notti', []))
    cfg.include_cxl = data.get('include_cxl', True)
    db.session.commit()
    all_names = [c.name for c in ExportConfig.query.order_by(ExportConfig.name).all()]
    return jsonify({'ok': True, 'configs': all_names})


@rooming_bp.route('/api/load-config/<path:name>')
def api_load_config(name):
    import json
    from models.models import ExportConfig
    cfg = ExportConfig.query.filter_by(name=name).first()
    if not cfg:
        return jsonify({'error': 'Configurazione non trovata'}), 404
    return jsonify({
        'cols':        json.loads(cfg.cols or '[]'),
        'hotels':      json.loads(cfg.hotels or '[]'),
        'stati':       json.loads(cfg.stati or '[]'),
        'notti':       json.loads(cfg.notti or '[]'),
        'include_cxl': cfg.include_cxl,
    })

# ── Hotel contracts ───────────────────────────────────────────────────────────

# Mapping nomi file → nomi DB
HOTEL_NAME_MAP = {
    'movenpick':    'Movenpick',
    'beau rivage':  'Beau-Rivage Palace',
    'beau-rivage':  'Beau-Rivage Palace',
    'de la paix':   'Hotel de la Paix',
    'lausanne':     'Lausanne Palace',
    'novotel':      'Novotel',
    'royal savoy':  'Royal Savoy',
    'starling':     'Starling',
}

DATES = [
    '2026-03-28', '2026-03-29', '2026-03-30', '2026-03-31',
    '2026-04-01', '2026-04-02', '2026-04-03', '2026-04-04',
]

DATE_LABELS = ['28-mar','29-mar','30-mar','31-mar','1-apr','2-apr','3-apr','4-apr']


def normalize_hotel(raw):
    """Normalizza nome hotel grezzo verso nome DB."""
    s = str(raw).strip().lower()
    for key, val in HOTEL_NAME_MAP.items():
        if key in s:
            return val
    return None


@rooming_bp.route('/upload-contracts', methods=['POST'])
def upload_contracts():
    from models.models import HotelContract
    import pandas as pd
    from datetime import date as _date

    f = request.files.get('file')
    if not f or not f.filename.endswith('.xlsx'):
        flash('Seleziona un file .xlsx valido.', 'error')
        return redirect(url_for('rooming.index'))

    df = pd.read_excel(f, sheet_name=0, header=None)

    inserted = 0
    errors   = []
    i = 0
    while i < len(df):
        raw_name = df.iloc[i, 1] if len(df.columns) > 1 else None
        if pd.isna(raw_name):
            i += 1
            continue

        hotel = normalize_hotel(raw_name)
        if not hotel:
            i += 1
            continue

        # Riga date (colonne 2-9) e riga valori (riga successiva)
        if i + 1 >= len(df):
            break

        val_row = df.iloc[i + 1]

        for col_offset, date_str in enumerate(DATES):
            col_idx = col_offset + 2
            try:
                val = val_row.iloc[col_idx] if col_idx < len(val_row) else 0
                rooms = int(val) if pd.notna(val) else 0
            except Exception:
                rooms = 0

            d = _date.fromisoformat(date_str)
            existing = HotelContract.query.filter_by(hotel=hotel, date=d).first()
            if existing:
                existing.rooms = rooms
            else:
                db.session.add(HotelContract(hotel=hotel, date=d, rooms=rooms))
            inserted += 1

        i += 2  # salta la coppia hotel+valori

    # Starling: assicura zero per tutte le date
    for date_str in DATES:
        d = _date.fromisoformat(date_str)
        if not HotelContract.query.filter_by(hotel='Starling', date=d).first():
            db.session.add(HotelContract(hotel='Starling', date=d, rooms=0))

    db.session.commit()
    flash(f'Contratti caricati: {inserted} record aggiornati.', 'success')
    return redirect(url_for('rooming.index'))


@rooming_bp.route('/api/categories/<path:batch_id>')
def api_categories(batch_id):
    """Ritorna le categorie azienda presenti nel batch, ordinate."""
    cats = db.session.execute(db.text("""
        SELECT DISTINCT company_category
        FROM rooming_list
        WHERE import_batch = :bid
        ORDER BY company_category NULLS LAST
    """), {'bid': batch_id}).fetchall()
    categories = [r[0] for r in cats]
    return jsonify({'categories': categories})



@rooming_bp.route('/report-transfer/<path:batch_id>', methods=['POST'])
def report_transfer(batch_id):
    """Excel transfer: righe=partecipanti, colonne=notti, no CXL.
    Le spouse (status_spouse=Yes) ereditano le notti del titolare via internal_parent_reference.
    """
    from sqlalchemy import or_
    import re as _re

    hotels = request.form.getlist('hotels')

    q = RoomingList.query.filter_by(import_batch=batch_id)\
                         .filter(RoomingList.registration_state != 'CXL')
    if hotels:
        conditions = [RoomingList.hotel == h for h in hotels]
        q = q.filter(or_(*conditions))

    rows = q.order_by(RoomingList.last_name, RoomingList.first_name).all()

    if not rows:
        flash('Nessun partecipante trovato.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    NIGHTS = [
        ('night_sat_28mar', '28/03'),
        ('night_sun_29mar', '29/03'),
        ('night_mon_30mar', '30/03'),
        ('night_tue_31mar', '31/03'),
        ('night_wed_1apr',  '01/04'),
        ('night_thu_2apr',  '02/04'),
        ('night_fri_3apr',  '03/04'),
        ('night_sat_4apr',  '04/04'),
    ]

    # Mappa internal_reference -> riga per tutto il batch (per lookup spouse)
    all_batch = RoomingList.query.filter_by(import_batch=batch_id)\
                                 .filter(RoomingList.registration_state != 'CXL').all()
    ref_map = {str(r.internal_reference).strip(): r
               for r in all_batch if r.internal_reference}


    def get_nights(r):
        """Restituisce dict {field: bool} con le notti effettive del partecipante.
        Se il partecipante non ha notti, cerca il compagno di stanza e usa le sue.
        La logica è simmetrica: non importa chi è il titolare.
        Gestisce anche night_no_need=Yes senza notti proprie."""

        own_nights = {f: bool(getattr(r, f)) for f, _ in NIGHTS}
        if any(own_nights.values()):
            return own_nights

        # Non ha notti — cerca compagno di stanza solo se:
        # - è spouse (status_spouse=Yes), oppure
        # - ha night_no_need=Yes (condivide camera ma non ha notti proprie)
        is_spouse   = (r.status_spouse or '').strip().lower() == 'yes'
        is_no_need  = (r.night_no_need or '').strip().lower() == 'yes'
        if not is_spouse and not is_no_need:
            return own_nights

        # 1. Via internal_parent_reference → cerca titolare
        if r.internal_parent_reference:
            parent = ref_map.get(str(r.internal_parent_reference).strip())
            if parent and any(bool(getattr(parent, f)) for f, _ in NIGHTS):
                return {f: bool(getattr(parent, f)) for f, _ in NIGHTS}

            # 1b. Stesso parent_ref — cerca un altro partecipante con stesso parent che ha notti
            parent_ref_str = str(r.internal_parent_reference).strip()
            for candidate in all_batch:
                if candidate.id == r.id:
                    continue
                if str(candidate.internal_parent_reference or '').strip() == parent_ref_str or \
                   str(candidate.internal_reference or '').strip() == parent_ref_str:
                    cand_nights = {f: bool(getattr(candidate, f)) for f, _ in NIGHTS}
                    if any(cand_nights.values()):
                        return cand_nights

        # 2. Cerca nel campo upgrade del batch chi ha "DOUBLE ROOM WITH / TWIN ROOM WITH <nome o cognome>"
        first = (r.first_name or '').strip().lower()
        last  = (r.last_name or '').strip().lower()
        for candidate in all_batch:
            if candidate.id == r.id:
                continue
            upgrade_text = (candidate.upgrade or '').lower()
            if any(kw in upgrade_text for kw in ('double room with', 'twin room with', 'shared room with')):
                if last in upgrade_text or (first and first in upgrade_text):
                    candidate_nights = {f: bool(getattr(candidate, f)) for f, _ in NIGHTS}
                    if any(candidate_nights.values()):
                        return candidate_nights
                    # Anche il candidato non ha notti — prova il suo parent
                    if candidate.internal_parent_reference:
                        cp = ref_map.get(str(candidate.internal_parent_reference).strip())
                        if cp:
                            cp_nights = {f: bool(getattr(cp, f)) for f, _ in NIGHTS}
                            if any(cp_nights.values()):
                                return cp_nights

        # 3. Cerca nel campo upgrade della spouse stessa
        upgrade_text = (r.upgrade or '').lower()
        if any(kw in upgrade_text for kw in ('double room with', 'twin room with', 'shared room with',
                                              'wife of', 'spouse of')):
            m = _re.search(r'(?:double|twin|shared)\s+room\s+with\s+([\w\s]+)|'
                           r'(?:wife|spouse)\s+of\s+([\w\s]+)',
                           upgrade_text, _re.IGNORECASE)
            if m:
                name_str = (m.group(1) or m.group(2) or '').strip().lower()
                name_parts = name_str.split()
                for candidate in all_batch:
                    if candidate.id == r.id:
                        continue
                    cln = (candidate.last_name or '').lower()
                    cfn = (candidate.first_name or '').lower()
                    if any(p == cln or p == cfn for p in name_parts):
                        candidate_nights = {f: bool(getattr(candidate, f)) for f, _ in NIGHTS}
                        if any(candidate_nights.values()):
                            return candidate_nights

        # 4. Fallback: stesso cognome + stesso hotel + non spouse + ha notti
        spouse_last  = (r.last_name or '').strip().lower()
        spouse_hotel = (r.hotel or '').strip().lower()
        for candidate in all_batch:
            if candidate.id == r.id:
                continue
            cand_last = (candidate.last_name or '').strip().lower()
            cand_hotel = (candidate.hotel or '').strip().lower()
            cand_nights = {f: bool(getattr(candidate, f)) for f, _ in NIGHTS}
            if cand_last == spouse_last and \
               cand_hotel == spouse_hotel and \
               (candidate.status_spouse or '').strip().lower() != 'yes':
                if any(cand_nights.values()):
                    return cand_nights

        # Nessun compagno trovato con notti — vuoto
        return {f: False for f, _ in NIGHTS}

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = 'Transfer'

    hdr_font  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', start_color='1F3864')
    norm_font = Font(name='Calibri', size=10)
    tot_font  = Font(name='Calibri', bold=True, size=10)
    tot_fill  = PatternFill('solid', start_color='D9E1F2')
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')
    thin   = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    n_cols = 3 + len(NIGHTS)

    # Titolo
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value='TRANSFER LIST')
    tc.font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    tc.alignment = center
    ws.row_dimensions[1].height = 26

    # Sottotitolo
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    hotel_str = ', '.join(hotels) if hotels else 'tutti gli hotel'
    ts_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    ws.cell(row=2, column=1,
            value=f'Hotel: {hotel_str}  |  {len(rows)} partecipanti  |  {ts_str}'
            ).font = Font(name='Calibri', size=9, italic=True)
    ws.row_dimensions[2].height = 14

    # Intestazioni
    headers = ['Cognome', 'Nome', 'Hotel'] + [label for _, label in NIGHTS]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = center; c.border = thin
    ws.row_dimensions[3].height = 20

    # Righe partecipanti + contatori totale
    totals = [0] * len(NIGHTS)
    for rn, r in enumerate(rows, 4):
        ws.cell(row=rn, column=1, value=r.last_name or '').font = norm_font
        ws.cell(row=rn, column=1).border = thin
        ws.cell(row=rn, column=1).alignment = left
        ws.cell(row=rn, column=2, value=r.first_name or '').font = norm_font
        ws.cell(row=rn, column=2).border = thin
        ws.cell(row=rn, column=2).alignment = left
        ws.cell(row=rn, column=3, value=r.hotel or '').font = norm_font
        ws.cell(row=rn, column=3).border = thin
        ws.cell(row=rn, column=3).alignment = left

        nights = get_nights(r)
        for ni, (field, _) in enumerate(NIGHTS):
            val = 'X' if nights[field] else ''
            c = ws.cell(row=rn, column=4+ni, value=val)
            c.font = norm_font; c.border = thin; c.alignment = center
            if val:
                totals[ni] += 1
        ws.row_dimensions[rn].height = 15

    # Riga totale
    tot_row = len(rows) + 4
    ws.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=3)
    tc = ws.cell(row=tot_row, column=1, value='TOTALE')
    tc.font = tot_font; tc.fill = tot_fill; tc.alignment = left; tc.border = thin
    for ni, t in enumerate(totals):
        c = ws.cell(row=tot_row, column=4+ni, value=t)
        c.font = tot_font; c.fill = tot_fill
        c.alignment = center; c.border = thin
    ws.row_dimensions[tot_row].height = 18

    # Larghezze colonne
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 20
    for ni in range(len(NIGHTS)):
        ws.column_dimensions[get_column_letter(4+ni)].width = 8

    ws.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    hotel_tag = hotels[0].replace(' ', '_')[:20] if len(hotels)==1 else 'multi'
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(buf, as_attachment=True,
                     download_name=f'Transfer_{hotel_tag}_{ts}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@rooming_bp.route('/api/billings/<path:batch_id>')
def api_billings(batch_id):
    """Ritorna i valori billing presenti nel batch, ordinati."""
    rows = db.session.execute(db.text("""
        SELECT DISTINCT billing
        FROM rooming_list
        WHERE import_batch = :bid
        ORDER BY billing NULLS LAST
    """), {'bid': batch_id}).fetchall()
    return jsonify({'billings': [r[0] for r in rows]})


@rooming_bp.route('/report-billing/<path:batch_id>', methods=['POST'])
def report_billing(batch_id):
    """Excel con partecipanti filtrati per billing, raggruppati per hotel."""
    from sqlalchemy import or_
    billings = request.form.getlist('billings')
    if not billings:
        flash('Seleziona almeno un valore billing.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    conditions = []
    for b in billings:
        if b == '__blank__':
            conditions.append(RoomingList.billing == None)
        else:
            conditions.append(RoomingList.billing == b)

    rows = RoomingList.query.filter_by(import_batch=batch_id)\
                            .filter(or_(*conditions))\
                            .order_by(RoomingList.hotel, RoomingList.last_name,
                                      RoomingList.first_name).all()

    if not rows:
        flash('Nessun partecipante trovato per i valori billing selezionati.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    def _sort(r):
        return (r.hotel or '', 1 if r.is_cxl else 0,
                str(r.last_name or '').upper(), str(r.first_name or '').upper())
    rows = sorted(rows, key=_sort)

    bill_labels = [b if b != '__blank__' else '(blank)' for b in billings]
    title_str   = ' · '.join(bill_labels)

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = 'Report Billing'

    hdr_font  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', start_color='1F3864')
    norm_font = Font(name='Calibri', size=10)
    green_fill  = PatternFill('solid', start_color='C6EFCE')
    yellow_fill = PatternFill('solid', start_color='FFEB9C')
    red_fill    = PatternFill('solid', start_color='FFC7CE')
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')
    thin   = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    n_cols = len(PREVIEW_COLS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value=f'REPORT BILLING — {title_str.upper()}')
    tc.font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    tc.alignment = center
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ts_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    ws.cell(row=2, column=1,
            value=f'Generato il {ts_str}  |  {len(rows)} partecipanti'
            ).font = Font(name='Calibri', size=9, italic=True)
    ws.row_dimensions[2].height = 14

    current_hotel = None
    row_num = 3

    for r in rows:
        hotel = r.hotel or 'SENZA HOTEL'
        if hotel != current_hotel:
            current_hotel = hotel
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=n_cols)
            hc = ws.cell(row=row_num, column=1, value=hotel.upper())
            hc.font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
            hc.fill = PatternFill('solid', start_color='2F5496')
            hc.alignment = left
            ws.row_dimensions[row_num].height = 20
            row_num += 1
            for col, (field, label) in enumerate(PREVIEW_COLS, 1):
                c = ws.cell(row=row_num, column=col, value=label)
                c.font = hdr_font; c.fill = hdr_fill
                c.alignment = center; c.border = thin
            ws.row_dimensions[row_num].height = 18
            row_num += 1

        is_changed = r.change_date or (r.latest_changes and str(r.latest_changes).strip())
        row_fill = red_fill if r.is_cxl else (yellow_fill if is_changed else green_fill)

        for col, (field, _) in enumerate(PREVIEW_COLS, 1):
            c = ws.cell(row=row_num, column=col, value=cell_val(field, r))
            c.font = norm_font; c.fill = row_fill
            c.alignment = left; c.border = thin
        ws.row_dimensions[row_num].height = 15
        row_num += 1

    for col, (field, _) in enumerate(PREVIEW_COLS, 1):
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS.get(field, 14)
    ws.freeze_panes = 'A3'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe = title_str.replace(' ', '_').replace('/', '-')[:30]
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(buf, as_attachment=True,
                     download_name=f'Report_Billing_{safe}_{ts}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@rooming_bp.route('/api/arrivals/<path:batch_id>')
def api_arrivals(batch_id):
    rows = db.session.execute(db.text("""
        SELECT DISTINCT arrival_mode
        FROM rooming_list
        WHERE import_batch = :bid
        ORDER BY arrival_mode NULLS LAST
    """), {'bid': batch_id}).fetchall()
    return jsonify({'arrivals': [r[0] for r in rows]})


@rooming_bp.route('/report-arrival/<path:batch_id>', methods=['POST'])
def report_arrival(batch_id):
    from sqlalchemy import or_
    arrivals = request.form.getlist('arrivals')
    if not arrivals:
        flash('Seleziona almeno un valore.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    conditions = [RoomingList.arrival_mode == None if a == '__blank__'
                  else RoomingList.arrival_mode == a for a in arrivals]

    rows = RoomingList.query.filter_by(import_batch=batch_id)\
                            .filter(or_(*conditions))\
                            .order_by(RoomingList.hotel, RoomingList.last_name,
                                      RoomingList.first_name).all()

    if not rows:
        flash('Nessun partecipante trovato.', 'error')
        return redirect(url_for('rooming.batch_detail', batch_id=batch_id))

    def _sort(r):
        return (r.hotel or '', 1 if r.is_cxl else 0,
                str(r.last_name or '').upper(), str(r.first_name or '').upper())
    rows = sorted(rows, key=_sort)

    labels    = [a if a != '__blank__' else '(blank)' for a in arrivals]
    title_str = ' · '.join(labels)

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = 'Report Arrival'

    hdr_font  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', start_color='1F3864')
    norm_font = Font(name='Calibri', size=10)
    green_fill  = PatternFill('solid', start_color='C6EFCE')
    yellow_fill = PatternFill('solid', start_color='FFEB9C')
    red_fill    = PatternFill('solid', start_color='FFC7CE')
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')
    thin   = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    n_cols = len(PREVIEW_COLS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value=f'REPORT ARRIVAL — {title_str.upper()}')
    tc.font = Font(name='Calibri', bold=True, size=14, color='1F3864')
    tc.alignment = center
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ts_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    ws.cell(row=2, column=1,
            value=f'Generato il {ts_str}  |  {len(rows)} partecipanti'
            ).font = Font(name='Calibri', size=9, italic=True)
    ws.row_dimensions[2].height = 14

    current_hotel = None
    row_num = 3

    for r in rows:
        hotel = r.hotel or 'SENZA HOTEL'
        if hotel != current_hotel:
            current_hotel = hotel
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=n_cols)
            hc = ws.cell(row=row_num, column=1, value=hotel.upper())
            hc.font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
            hc.fill = PatternFill('solid', start_color='2F5496')
            hc.alignment = left
            ws.row_dimensions[row_num].height = 20
            row_num += 1
            for col, (field, label) in enumerate(PREVIEW_COLS, 1):
                c = ws.cell(row=row_num, column=col, value=label)
                c.font = hdr_font; c.fill = hdr_fill
                c.alignment = center; c.border = thin
            ws.row_dimensions[row_num].height = 18
            row_num += 1

        is_changed = r.change_date or (r.latest_changes and str(r.latest_changes).strip())
        row_fill = red_fill if r.is_cxl else (yellow_fill if is_changed else green_fill)

        for col, (field, _) in enumerate(PREVIEW_COLS, 1):
            c = ws.cell(row=row_num, column=col, value=cell_val(field, r))
            c.font = norm_font; c.fill = row_fill
            c.alignment = left; c.border = thin
        ws.row_dimensions[row_num].height = 15
        row_num += 1

    for col, (field, _) in enumerate(PREVIEW_COLS, 1):
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS.get(field, 14)
    ws.freeze_panes = 'A3'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe = title_str.replace(' ', '_').replace('/', '-')[:30]
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(buf, as_attachment=True,
                     download_name=f'Report_Arrival_{safe}_{ts}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@rooming_bp.route('/api/occupancy/<path:batch_id>')

def api_occupancy(batch_id):
    """Ritorna JSON con CONFIRMED vs BY CONTRACT per hotel x data."""
    from models.models import HotelContract
    from datetime import date as _date

    rows = RoomingList.query.filter_by(import_batch=batch_id)\
                            .filter(RoomingList.registration_state != 'CXL')\
                            .all()

    # Notti field → date
    NIGHT_TO_DATE = {
        'night_sat_28mar': '2026-03-28',
        'night_sun_29mar': '2026-03-29',
        'night_mon_30mar': '2026-03-30',
        'night_tue_31mar': '2026-03-31',
        'night_wed_1apr':  '2026-04-01',
        'night_thu_2apr':  '2026-04-02',
        'night_fri_3apr':  '2026-04-03',
        'night_sat_4apr':  '2026-04-04',
    }

    hotels = sorted(set(r.hotel for r in rows if r.hotel))

    # Costruisce confirmed[hotel][date] = count
    confirmed = {h: {d: 0 for d in DATES} for h in hotels}
    for r in rows:
        if not r.hotel:
            continue
        for field, date_str in NIGHT_TO_DATE.items():
            if getattr(r, field):
                confirmed[r.hotel][date_str] += 1

    # Carica contratti
    contracts = HotelContract.query.all()
    contract_map = {(c.hotel, str(c.date)): c.rooms for c in contracts}

    result = {}
    for h in hotels:
        result[h] = {
            'confirmed': [confirmed[h].get(d, 0) for d in DATES],
            'contract':  [contract_map.get((h, d), 0) for d in DATES],
        }

    return jsonify({
        'hotels':       hotels,
        'dates':        DATE_LABELS,
        'data':         result,
        'has_contracts': len(contracts) > 0,
    })
