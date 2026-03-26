import os
from flask import Flask, session, redirect, url_for, request, render_template, flash
from models import db

if os.environ.get('RAILWAY_ENVIRONMENT') is None:
    try:
        from dotenv import load_dotenv
        # Cerca sia .env che env
        _env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'env')
        if os.path.exists(_env_path):
            load_dotenv(_env_path)
        else:
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
            db_url = 'postgresql://postgres:OgzeTFkNlxACnybDyXaCpFvlXXekkAln@postgres.railway.internal:5432/railway'
        else:
            db_url = 'postgresql://postgres:123456@localhost:5432/nexus_rooming'

    # Railway a volte fornisce postgres:// invece di postgresql://
    if db_url.startswith('postgres://'):
        db_url = db_url.replace('postgres://', 'postgresql://', 1)

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

    # ── AUTH NOMINALE ────────────────────────────────────────────────────────
    @app.before_request
    def check_auth():
        if request.endpoint in ('login', 'logout', 'static'):
            return
        if not session.get('username'):
            return redirect(url_for('login', next=request.url))

    @app.route('/login', methods=['GET', 'POST'])
    def login():
        from models.models import User
        if request.method == 'POST':
            username = request.form.get('username', '').strip()
            password = request.form.get('password', '')
            user = User.query.filter(
                (User.username == username) | (User.email == username)
            ).first()
            if user and user.check_password(password) and user.is_active:
                session['username'] = user.username or user.email
                session.permanent = False
                next_url = request.args.get('next') or url_for('rooming.index')
                return redirect(next_url)
            flash('Credenziali errate.', 'error')
        return render_template('login.html')

    @app.route('/logout')
    def logout():
        session.clear()
        return redirect(url_for('login'))

    @app.route('/admin/users', methods=['GET', 'POST'])
    def admin_users():
        from models.models import User
        if request.method == 'POST':
            action   = request.form.get('action')
            username = request.form.get('username', '').strip()
            password = request.form.get('password', '').strip()
            user_id  = request.form.get('user_id')
            if action == 'add':
                if User.query.filter_by(username=username).first():
                    flash('Username già esistente.', 'error')
                else:
                    u = User(username=username, email=username)
                    u.set_password(password)
                    db.session.add(u)
                    db.session.commit()
                    flash(f'Utente {username} creato.', 'success')
            elif action == 'delete' and user_id:
                u = User.query.get(int(user_id))
                if u and (u.username or u.email) != session.get('username'):
                    db.session.delete(u)
                    db.session.commit()
                    flash('Utente eliminato.', 'success')
                else:
                    flash('Non puoi eliminare te stesso.', 'error')
            elif action == 'reset_password' and user_id:
                u = User.query.get(int(user_id))
                if u:
                    u.set_password(password)
                    db.session.commit()
                    flash(f'Password aggiornata per {u.username or u.email}.', 'success')
        users = User.query.order_by(User.username).all()
        return render_template('admin_users.html', users=users)

    # ── BLUEPRINT ─────────────────────────────────────────────────────────────
    from routes.rooming import rooming_bp
    app.register_blueprint(rooming_bp)

    with app.app_context():
        db.create_all()

    return app


if __name__ == '__main__':
    app = create_app()
    app.run(debug=True, port=5001)
