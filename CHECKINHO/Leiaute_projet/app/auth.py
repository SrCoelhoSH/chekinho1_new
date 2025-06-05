from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import LoginManager, login_user, login_required, logout_user, UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3
from db import get_db_connection

bp_auth = Blueprint('auth', __name__)
login_manager = LoginManager()
login_manager.login_view = 'auth.login'

class User(UserMixin):
    def __init__(self, id_, username, password_hash):
        self.id = id_
        self.username = username
        self.password_hash = password_hash

    @staticmethod
    def get(user_id):
        conn = get_db_connection()
        user = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,)).fetchone()
        conn.close()
        if user:
            return User(user['id'], user['username'], user['password_hash'])
        return None

@login_manager.user_loader
def load_user(user_id):
    return User.get(int(user_id))

@bp_auth.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        conn = get_db_connection()
        try:
            conn.execute(
                'INSERT INTO users (username, password_hash) VALUES (?, ?)',
                (username, generate_password_hash(password))
            )
            conn.commit()
            flash('Usuário criado com sucesso.')
            return redirect(url_for('auth.login'))
        except sqlite3.IntegrityError:
            flash('Nome de usuário já existe.')
        finally:
            conn.close()
    return render_template('register.html')

@bp_auth.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        conn = get_db_connection()
        user = conn.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
        conn.close()
        if user and check_password_hash(user['password_hash'], password):
            user_obj = User(user['id'], user['username'], user['password_hash'])
            login_user(user_obj)
            return redirect(url_for('main.upload_files'))
        flash('Credenciais inválidas.')
    return render_template('login.html')

@bp_auth.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('auth.login'))

