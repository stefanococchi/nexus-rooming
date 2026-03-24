"""
Script una-tantum per creare il primo utente admin.
Eseguire con:  python create_admin.py
"""
import os
import sys

# Carica .env se presente
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from app import create_app
from models import db
from models.models import User

app = create_app()

with app.app_context():
    print("=== Creazione utente admin ===")
    username = input("Username: ").strip()
    if not username:
        print("Username non può essere vuoto.")
        sys.exit(1)

    if User.query.filter_by(username=username).first():
        print(f"Utente '{username}' già esistente.")
        sys.exit(1)

    import getpass
    password = getpass.getpass("Password: ")
    if len(password) < 4:
        print("Password troppo corta (minimo 4 caratteri).")
        sys.exit(1)

    confirm = getpass.getpass("Conferma password: ")
    if password != confirm:
        print("Le password non coincidono.")
        sys.exit(1)

    u = User(username=username, email=username, is_active=True)
    u.set_password(password)
    db.session.add(u)
    db.session.commit()

    print(f"\nUtente '{username}' creato con successo.")
    print("Puoi ora accedere al portale con queste credenziali.")
