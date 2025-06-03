# app.py

import socket
from flask import Flask
from routes import bp as main_bp
import os

def create_app():
    """
    Cria a instância da aplicação Flask, registra o blueprint e retorna o app.
    """
    THIS_DIR = os.path.dirname(os.path.abspath(__file__))
    # Sobe um nível, para sair de 'app' e chegar no diretório raiz
    PARENT_DIR = os.path.dirname(THIS_DIR)
    templates_dir = os.path.join(PARENT_DIR, "templates")

    app = Flask(__name__, template_folder=templates_dir)
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # Limite de 100MB (exemplo)
    app.register_blueprint(main_bp)  # Registra o blueprint definido em routes.py
    return app

if __name__ == '__main__':
    # Obtém o IP local da máquina
    host_ip = socket.gethostbyname(socket.gethostname())
    # Cria e roda a aplicação Flask
    create_app().run(host=host_ip, debug=True, port=80)
