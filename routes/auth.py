from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user
from models import db
from models.models import User

auth_bp = Blueprint('auth', __name__)


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('rooming.index'))

    if request.method == 'POST':
        email    = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        user     = User.query.filter_by(email=email, is_active=True).first()

        if not user:
            flash('Email non autorizzata.', 'error')
            return render_template('login.html')

        if not user.has_password():
            if len(password) < 6:
                flash('La password deve essere di almeno 6 caratteri.', 'error')
                return render_template('login.html', first_login=True, email=email)
            password2 = request.form.get('password2', '')
            if password != password2:
                flash('Le password non coincidono.', 'error')
                return render_template('login.html', first_login=True, email=email)
            user.set_password(password)
            db.session.commit()
            login_user(user)
            flash('Benvenuto! Password impostata con successo.', 'success')
            return redirect(url_for('rooming.index'))

        if not user.check_password(password):
            flash('Password errata.', 'error')
            return render_template('login.html')

        login_user(user)
        next_page = request.args.get('next')
        return redirect(next_page or url_for('rooming.index'))

    email_param = request.args.get('email', '')
    first_login = False
    if email_param:
        u = User.query.filter_by(email=email_param.lower(), is_active=True).first()
        first_login = u and not u.has_password()

    return render_template('login.html', first_login=first_login, email=email_param)


@auth_bp.route('/check-email', methods=['POST'])
def check_email():
    email = request.form.get('email', '').strip().lower()
    user  = User.query.filter_by(email=email, is_active=True).first()
    if not user:
        return {'status': 'not_found'}, 200
    if not user.has_password():
        return {'status': 'first_login'}, 200
    return {'status': 'ok'}, 200


@auth_bp.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('auth.login'))
