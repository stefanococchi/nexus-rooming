import os
from flask import Flask, session, redirect, url_for, request, render_template, flash
from models import db

if os.environ.get('RAILWAY_ENVIRONMENT') is None:
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass


def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'nexus-rooming-secret-2024')

    # ── DATABASE ──────────────────────────────────────────────────────────────
    db_url = os.environ.get('DATABASE_URL')
    if not db_url:
        if os.environ.get('RAILWAY_ENVIRONMENT'):
            db_url = os.environ.get('DATABASE_PRIVATE_URL', '')
        else:
            db_url = 'postgresql://postgres:123456@localhost:5432/nexus_rooming'

    app.config['SQLALCHEMY_DATABASE_URI'] = db_url
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_pre_ping': True,
        'pool_recycle': 300,
    }
    app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

    # ── VERSIONE ──────────────────────────────────────────────────────────────
    _version_file = os.path.join(app.root_path, 'VERSION')
    try:
        with open(_version_file) as vf:
            _version = vf.read().strip()
    except Exception:
        _version = '—'

    @app.context_processor
    def inject_version():
        return {'app_version': _version}

    # ── ESTENSIONI ────────────────────────────────────────────────────────────
    db.init_app(app)

    # ── AUTH SEMPLICE ─────────────────────────────────────────────────────────
    APP_PASSWORD = os.environ.get('APP_PASSWORD', 'nexus2026')

    @app.before_request
    def check_auth():
        # Percorsi sempre accessibili
        if request.endpoint in ('login', 'logout', 'static'):
            return
        if not session.get('authenticated'):
            return redirect(url_for('login', next=request.url))

    @app.route('/login', methods=['GET', 'POST'])
    def login():
        if request.method == 'POST':
            pwd = request.form.get('password', '')
            if pwd == APP_PASSWORD:
                session['authenticated'] = True
                session.permanent = False
                next_url = request.args.get('next') or url_for('rooming.index')
                return redirect(next_url)
            flash('Password errata.', 'error')
        return render_template('login.html')

    @app.route('/logout')
    def logout():
        session.clear()
        return redirect(url_for('login'))

    # ── BLUEPRINT ─────────────────────────────────────────────────────────────
    from routes.rooming import rooming_bp
    app.register_blueprint(rooming_bp)

    with app.app_context():
        db.create_all()

    return app


if __name__ == '__main__':
    app = create_app()
    app.run(debug=True, port=5001)
