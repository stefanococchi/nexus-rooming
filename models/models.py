from datetime import datetime
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from models import db


# ─── USER (condiviso con saba-form — stessa tabella) ─────────────────────────

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id            = db.Column(db.Integer, primary_key=True)
    email         = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=True)
    is_superuser  = db.Column(db.Boolean, default=False)
    can_create_events   = db.Column(db.Boolean, default=False)
    can_see_all_events  = db.Column(db.Boolean, default=False)
    created_at    = db.Column(db.DateTime, default=datetime.utcnow)
    is_active     = db.Column(db.Boolean, default=True)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def has_password(self):
        return self.password_hash is not None


# ─── ROOMING LIST ─────────────────────────────────────────────────────────────

class RoomingList(db.Model):
    __tablename__ = 'rooming_list'

    id                          = db.Column(db.Integer, primary_key=True)

    # Metadati import
    imported_at                 = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    import_batch                = db.Column(db.String(300), nullable=False, index=True)

    # Col 0
    registration_state          = db.Column(db.String(50))
    # Col 1
    latest_changes              = db.Column(db.Text)
    # Col 2
    hotel                       = db.Column(db.String(200), index=True)
    # Col 3
    upgrade                     = db.Column(db.String(100))
    # Col 4
    participant_display         = db.Column(db.String(300))
    # Col 5
    billing                     = db.Column(db.String(100))
    # Col 6
    company_name                = db.Column(db.String(300))
    # Col 7
    company_country             = db.Column(db.String(100))
    # Col 8
    nexus_bd                    = db.Column(db.String(200))
    # Col 9
    is_parent_manager           = db.Column(db.String(10))
    # Col 10
    registered_colleagues       = db.Column(db.Integer)
    # Col 11 — chiave univoca Nexus
    internal_reference          = db.Column(db.String(50), index=True)
    # Col 12
    internal_parent_reference   = db.Column(db.String(50))
    # Col 13
    ean8_barcode                = db.Column(db.String(20))
    # Col 14
    participant_number          = db.Column(db.Integer)
    # Col 15
    external_reference          = db.Column(db.String(50))
    # Col 16
    delegation_key              = db.Column(db.String(100))
    # Col 17-22: ruoli
    status_vp_bd                = db.Column(db.String(10))
    status_organisator          = db.Column(db.String(10))
    status_board_nai            = db.Column(db.String(10))
    status_climate_day          = db.Column(db.String(10))
    status_prospective_council  = db.Column(db.String(10))
    status_spouse               = db.Column(db.String(10))
    # Col 23
    comment                     = db.Column(db.Text)
    # Col 24-26
    title                       = db.Column(db.String(20))
    last_name                   = db.Column(db.String(100))
    first_name                  = db.Column(db.String(100))
    # Col 27-29
    job_position                = db.Column(db.String(300))
    email                       = db.Column(db.String(200))
    phone                       = db.Column(db.String(50))
    # Col 30-31
    prospective_title           = db.Column(db.String(100))
    prospective_response        = db.Column(db.String(100))
    # Col 32-41: notti
    night_no_need               = db.Column(db.String(10))
    night_sat_28mar             = db.Column(db.String(10))
    night_sun_29mar             = db.Column(db.String(10))
    night_mon_30mar             = db.Column(db.String(10))
    night_tue_31mar             = db.Column(db.String(10))
    night_wed_1apr              = db.Column(db.String(10))
    night_thu_2apr              = db.Column(db.String(10))
    night_fri_3apr              = db.Column(db.String(10))
    night_sat_4apr              = db.Column(db.String(10))
    night_other                 = db.Column(db.String(10))
    # Col 42-44
    diet_restrictions           = db.Column(db.String(300))
    arrival_mode                = db.Column(db.String(50))
    need_smooth_checkin         = db.Column(db.String(10))
    # Col 45-51
    need_visa                   = db.Column(db.String(10))
    visa_birth_date             = db.Column(db.Date)
    visa_birth_place            = db.Column(db.String(200))
    visa_passport               = db.Column(db.String(50))
    visa_delivery_date          = db.Column(db.Date)
    visa_expiration_date        = db.Column(db.Date)
    visa_company_address        = db.Column(db.Text)
    # Col 52-53
    company_category            = db.Column(db.String(100))
    company_subcategory         = db.Column(db.String(100))

    @property
    def full_name(self):
        parts = [self.title or '', self.first_name or '', self.last_name or '']
        return ' '.join(p for p in parts if p).strip()

    @property
    def is_cxl(self):
        return (self.registration_state or '').strip().upper() == 'CXL'


# ─── EXPORT CONFIG ────────────────────────────────────────────────────────────

class ExportConfig(db.Model):
    __tablename__ = 'export_configs'

    id          = db.Column(db.Integer, primary_key=True)
    name        = db.Column(db.String(200), nullable=False, unique=True)
    cols        = db.Column(db.Text, nullable=False)   # JSON array
    hotels      = db.Column(db.Text, nullable=True)    # JSON array
    stati       = db.Column(db.Text, nullable=True)    # JSON array
    notti       = db.Column(db.Text, nullable=True)    # JSON array
    include_cxl = db.Column(db.Boolean, default=True)
    created_at  = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at  = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


# ─── HOTEL CONTRACT ───────────────────────────────────────────────────────────

class HotelContract(db.Model):
    __tablename__ = 'hotel_contracts'

    id       = db.Column(db.Integer, primary_key=True)
    hotel    = db.Column(db.String(200), nullable=False, index=True)
    date     = db.Column(db.Date, nullable=False)
    rooms    = db.Column(db.Integer, nullable=False, default=0)

    __table_args__ = (
        db.UniqueConstraint('hotel', 'date', name='uq_hotel_contract_date'),
    )
