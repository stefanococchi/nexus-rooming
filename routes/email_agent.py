"""
Email Agent — Interpreta email con richieste di modifica rooming
e applica night change / hotel change / add person automaticamente.
"""

import os
import json
from datetime import datetime

from flask import (Blueprint, render_template, request, jsonify,
                   session as flask_session)
from sqlalchemy import func

from models import db
from models.models import RoomingList, ManualOverride, ModificationLog
from routes.rooming import get_batches, EDITABLE_FIELDS, NIGHT_DATES

email_agent_bp = Blueprint('email_agent', __name__)

# ── Mapping notti ────────────────────────────────────────────────────────────

NIGHT_FIELD_LABELS = {
    'night_sat_28mar': '28/03',
    'night_sun_29mar': '29/03',
    'night_mon_30mar': '30/03',
    'night_tue_31mar': '31/03',
    'night_wed_1apr':  '01/04',
    'night_thu_2apr':  '02/04',
    'night_fri_3apr':  '03/04',
    'night_sat_4apr':  '04/04',
}

NIGHT_FIELDS = list(NIGHT_FIELD_LABELS.keys())


# ── Helpers ──────────────────────────────────────────────────────────────────

def _get_active_nights(row):
    """Restituisce lista dei campi notte attivi per un partecipante."""
    return [f for f in NIGHT_FIELDS if getattr(row, f, None) in ('Yes', 'yes', 'YES')]


def _lookup_participant(last_name, first_name, batch_id):
    """Cerca partecipante per cognome (+ nome opzionale).
    Ritorna (matches, status) dove status in: exact, unique, ambiguous, not_found.
    """
    base = RoomingList.query.filter(
        RoomingList.import_batch == batch_id,
        func.upper(RoomingList.last_name) == last_name.upper().strip()
    )

    if first_name:
        exact = base.filter(
            func.upper(RoomingList.first_name) == first_name.upper().strip()
        ).all()
        if len(exact) == 1:
            return exact, 'exact'
        if len(exact) > 1:
            return exact, 'ambiguous'

    results = base.all()
    if len(results) == 1:
        return results, 'unique'
    if len(results) == 0:
        return [], 'not_found'
    return results, 'ambiguous'


def _apply_override(ref, field, new_val, username, batch_id):
    """Applica un override manuale (replica logica di /api/override)."""
    row = RoomingList.query.filter_by(
        import_batch=batch_id, internal_reference=ref).first()
    original_val = str(getattr(row, field) or '') if row else ''

    ov = ManualOverride.query.filter_by(
        internal_reference=ref, field=field).first()
    if ov:
        ov.original_value = original_val
        ov.override_value = new_val
        ov.modified_at    = datetime.now()
        ov.modified_by    = username
    else:
        ov = ManualOverride(
            internal_reference=ref, field=field,
            original_value=original_val, override_value=new_val,
            modified_at=datetime.now(), modified_by=username,
        )
        db.session.add(ov)

    if row:
        setattr(row, field, new_val)
        today_str = datetime.now().strftime('%d/%m/%Y')
        row.latest_changes = f'MANUAL EDIT - {today_str} ({username})'
        row.change_type    = 'MANUAL EDIT'
        row.change_date    = datetime.now().date()


def _log_modification(ref, name, category, action, hotel, details,
                      night_impacts, username):
    """Registra entry nel ModificationLog."""
    log = ModificationLog(
        internal_reference=ref,
        participant_name=name,
        category=category,
        action=action,
        hotel=hotel,
        details=details,
        night_impacts=json.dumps(night_impacts) if night_impacts else None,
        modified_at=datetime.now(),
        modified_by=username,
    )
    db.session.add(log)


def _add_manual_participant(fields, username, batch_id):
    """Crea nuovo partecipante MANUAL-NNN. Ritorna (new_row, new_ref)."""
    existing = db.session.execute(db.text(
        "SELECT internal_reference FROM rooming_list "
        "WHERE internal_reference LIKE 'MANUAL-%' "
        "ORDER BY internal_reference DESC LIMIT 1"
    )).fetchone()

    if existing:
        try:
            last_num = int(existing[0].replace('MANUAL-', ''))
            new_ref  = f'MANUAL-{last_num + 1:03d}'
        except Exception:
            new_ref = 'MANUAL-001'
    else:
        new_ref = 'MANUAL-001'

    today     = datetime.now().date()
    today_str = datetime.now().strftime('%d/%m/%Y')

    row = RoomingList(
        import_batch       = batch_id,
        internal_reference = new_ref,
        change_type        = 'MANUAL ADD',
        change_date        = today,
        latest_changes     = f'MANUAL ADD - {today_str} ({username})',
        registration_state = 'OK',
    )

    allowed = {f for f, _ in EDITABLE_FIELDS}
    for field, value in fields.items():
        if field in allowed and hasattr(row, field):
            setattr(row, field, value if value != '' else None)

    db.session.add(row)
    return row, new_ref


# ── Claude API prompt ────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are parsing email requests about hotel rooming changes for the NEXUS event (28 March - 5 April 2026).

Extract ALL requested changes as a JSON object.

Available action types:
- CANCEL_NIGHTS: Cancel/remove specific nights or ALL nights for an existing participant
- ADD_NIGHTS: Add specific nights for an existing participant
- HOTEL_CHANGE: Change hotel assignment for an existing participant
- ADD_PERSON: Add a brand new person not currently in the system

Night field mapping (use these EXACT field names):
- night_sat_28mar = night of March 28 (Saturday)
- night_sun_29mar = night of March 29 (Sunday)
- night_mon_30mar = night of March 30 (Monday)
- night_tue_31mar = night of March 31 (Tuesday)
- night_wed_1apr  = night of April 1 (Wednesday)
- night_thu_2apr  = night of April 2 (Thursday)
- night_fri_3apr  = night of April 3 (Friday)
- night_sat_4apr  = night of April 4 (Saturday)

Known hotels in the system:
{hotel_list}

Return ONLY valid JSON (no markdown, no code blocks) with this schema:
{{
  "actions": [
    {{
      "type": "CANCEL_NIGHTS" | "ADD_NIGHTS" | "HOTEL_CHANGE" | "ADD_PERSON",
      "last_name": "SURNAME IN UPPERCASE",
      "first_name": "FirstName or null",
      "hotel": "exact hotel name from the known list, or null",
      "nights": ["night_tue_31mar", "night_wed_1apr"],
      "all_nights": true/false,
      "new_hotel": "for HOTEL_CHANGE: the new hotel name, otherwise null",
      "reason": "brief explanation"
    }}
  ],
  "summary": "One-line summary of the email request"
}}

Rules:
- Last names MUST be UPPERCASE
- When email says "cancel nights" without specifying which ones, set all_nights=true and nights=[]
- When email says someone is "replaced by" a new person, that means: CANCEL_NIGHTS for the original person + ADD_PERSON for the new one
- For ADD_PERSON, include the specific nights and hotel
- "the 31st" in context of March/April means March 31 = night_tue_31mar
- "1st April" = night_wed_1apr
- Match hotel names to the known list as closely as possible
- If the email mentions multiple people, create separate actions for each
- Be precise about which nights are referenced
"""


# ── Routes ───────────────────────────────────────────────────────────────────

@email_agent_bp.route('/email-agent')
def email_agent_page():
    return render_template('email_agent.html')


@email_agent_bp.route('/api/email-agent/parse', methods=['POST'])
def api_parse_email():
    """Riceve testo email, chiama Claude API per estrarre azioni strutturate."""
    import anthropic

    data = request.get_json()
    email_text = (data.get('email_text') or '').strip()
    if not email_text:
        return jsonify({'ok': False, 'error': 'Testo email vuoto'})

    batches = get_batches()
    if not batches:
        return jsonify({'ok': False, 'error': 'Nessun batch disponibile'})
    batch_id = batches[0][0]

    # Raccogli hotel distinti
    hotels = db.session.execute(db.text(
        "SELECT DISTINCT hotel FROM rooming_list "
        "WHERE import_batch = :bid AND hotel IS NOT NULL ORDER BY hotel"
    ), {'bid': batch_id}).fetchall()
    hotel_list = [h[0] for h in hotels if h[0]]

    # Chiama Claude API
    api_key = os.environ.get('ANTHROPIC_API_KEY') or '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'

    client = anthropic.Anthropic(api_key=api_key)
    system = SYSTEM_PROMPT.format(hotel_list='\n'.join(f'- {h}' for h in hotel_list))

    try:
        response = client.messages.create(
            model='claude-haiku-4-5-20251001',
            max_tokens=2048,
            system=system,
            messages=[{'role': 'user', 'content': email_text}],
        )
        raw = response.content[0].text.strip()
        # Rimuovi eventuale markdown code block
        if raw.startswith('```'):
            raw = raw.split('\n', 1)[1] if '\n' in raw else raw[3:]
            if raw.endswith('```'):
                raw = raw[:-3]
            raw = raw.strip()
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return jsonify({'ok': False,
                        'error': f'Risposta Claude non valida: {raw[:500]}'})
    except Exception as e:
        return jsonify({'ok': False, 'error': f'Errore Claude API: {str(e)}'})

    # Arricchisci ogni azione con lookup partecipante
    enriched = []
    for i, action in enumerate(parsed.get('actions', [])):
        act = {
            'id': i,
            'type': action.get('type', ''),
            'last_name': action.get('last_name', ''),
            'first_name': action.get('first_name'),
            'hotel': action.get('hotel'),
            'nights': action.get('nights', []),
            'all_nights': action.get('all_nights', False),
            'new_hotel': action.get('new_hotel'),
            'reason': action.get('reason', ''),
            'status': 'resolved',
            'match': None,
            'candidates': [],
        }

        if act['type'] == 'ADD_PERSON':
            act['status'] = 'new_person'
        else:
            matches, status = _lookup_participant(
                act['last_name'], act['first_name'], batch_id)

            if status in ('exact', 'unique'):
                row = matches[0]
                active_nights = _get_active_nights(row)
                act['match'] = {
                    'internal_reference': row.internal_reference,
                    'display_name': f'{row.first_name or ""} {row.last_name or ""}'.strip(),
                    'current_hotel': row.hotel or '',
                    'current_nights': active_nights,
                }
                # Per CANCEL_NIGHTS con all_nights, risolvi le notti effettive
                if act['type'] == 'CANCEL_NIGHTS' and act['all_nights']:
                    act['nights'] = active_nights
            elif status == 'ambiguous':
                act['status'] = 'ambiguous'
                act['candidates'] = [{
                    'internal_reference': m.internal_reference,
                    'display_name': f'{m.first_name or ""} {m.last_name or ""}'.strip(),
                    'current_hotel': m.hotel or '',
                    'current_nights': _get_active_nights(m),
                } for m in matches]
            else:
                act['status'] = 'not_found'

        enriched.append(act)

    return jsonify({
        'ok': True,
        'actions': enriched,
        'summary': parsed.get('summary', ''),
    })


@email_agent_bp.route('/api/email-agent/execute', methods=['POST'])
def api_execute_actions():
    """Esegue le azioni confermate dall'utente."""
    data = request.get_json()
    actions = data.get('actions', [])
    username = flask_session.get('username', 'unknown')

    batches = get_batches()
    if not batches:
        return jsonify({'ok': False, 'error': 'Nessun batch disponibile'})
    batch_id = batches[0][0]

    results = []
    try:
        for act in actions:
            act_type = act.get('type', '')
            ref      = act.get('internal_reference', '')
            name     = act.get('display_name', '')
            hotel    = act.get('hotel') or ''
            nights   = act.get('nights', [])

            if act_type == 'CANCEL_NIGHTS':
                if not ref:
                    results.append({'id': act['id'], 'ok': False,
                                    'error': 'Nessun partecipante selezionato'})
                    continue

                # Cancella ogni notte
                night_impacts = {}
                for nf in nights:
                    _apply_override(ref, nf, '', username, batch_id)
                    night_impacts[nf] = -1

                # Recupera hotel corrente per il log
                row = RoomingList.query.filter_by(
                    import_batch=batch_id, internal_reference=ref).first()
                current_hotel = row.hotel if row else hotel

                _log_modification(
                    ref=ref, name=name,
                    category='NIGHT_CHANGE', action='DEL',
                    hotel=current_hotel,
                    details=f'Email Agent: cancellate notti {", ".join(NIGHT_FIELD_LABELS.get(n, n) for n in nights)}',
                    night_impacts=night_impacts,
                    username=username,
                )
                results.append({'id': act['id'], 'ok': True})

            elif act_type == 'ADD_NIGHTS':
                if not ref:
                    results.append({'id': act['id'], 'ok': False,
                                    'error': 'Nessun partecipante selezionato'})
                    continue

                night_impacts = {}
                for nf in nights:
                    _apply_override(ref, nf, 'Yes', username, batch_id)
                    night_impacts[nf] = 1

                row = RoomingList.query.filter_by(
                    import_batch=batch_id, internal_reference=ref).first()
                current_hotel = row.hotel if row else hotel

                _log_modification(
                    ref=ref, name=name,
                    category='NIGHT_CHANGE', action='ADD',
                    hotel=current_hotel,
                    details=f'Email Agent: aggiunte notti {", ".join(NIGHT_FIELD_LABELS.get(n, n) for n in nights)}',
                    night_impacts=night_impacts,
                    username=username,
                )
                results.append({'id': act['id'], 'ok': True})

            elif act_type == 'HOTEL_CHANGE':
                if not ref:
                    results.append({'id': act['id'], 'ok': False,
                                    'error': 'Nessun partecipante selezionato'})
                    continue

                new_hotel = act.get('new_hotel') or hotel
                row = RoomingList.query.filter_by(
                    import_batch=batch_id, internal_reference=ref).first()
                old_hotel = row.hotel if row else ''
                active_nights = _get_active_nights(row) if row else []

                _apply_override(ref, 'hotel', new_hotel, username, batch_id)

                # Log DEL dal vecchio hotel
                if old_hotel:
                    _log_modification(
                        ref=ref, name=name,
                        category='HOTEL_CHANGE', action='DEL',
                        hotel=old_hotel,
                        details=f'Email Agent: spostato a {new_hotel}',
                        night_impacts={n: -1 for n in active_nights},
                        username=username,
                    )
                # Log ADD al nuovo hotel
                _log_modification(
                    ref=ref, name=name,
                    category='HOTEL_CHANGE', action='ADD',
                    hotel=new_hotel,
                    details=f'Email Agent: da {old_hotel}',
                    night_impacts={n: 1 for n in active_nights},
                    username=username,
                )
                results.append({'id': act['id'], 'ok': True})

            elif act_type == 'ADD_PERSON':
                fields = {
                    'last_name':  act.get('last_name', ''),
                    'first_name': act.get('first_name', ''),
                    'hotel':      hotel,
                }
                for nf in nights:
                    fields[nf] = 'Yes'

                new_row, new_ref = _add_manual_participant(
                    fields, username, batch_id)

                night_impacts = {nf: 1 for nf in nights}
                display = f'{act.get("first_name", "")} {act.get("last_name", "")}'.strip()
                _log_modification(
                    ref=new_ref, name=display,
                    category='ADD_PERSON', action='NEW',
                    hotel=hotel,
                    details=f'Email Agent: aggiunto con notti {", ".join(NIGHT_FIELD_LABELS.get(n, n) for n in nights)}',
                    night_impacts=night_impacts,
                    username=username,
                )
                results.append({'id': act['id'], 'ok': True, 'ref': new_ref})

            else:
                results.append({'id': act['id'], 'ok': False,
                                'error': f'Tipo azione sconosciuto: {act_type}'})

        db.session.commit()
    except Exception as e:
        db.session.rollback()
        return jsonify({'ok': False, 'error': f'Errore: {str(e)}'})

    return jsonify({'ok': True, 'results': results})
