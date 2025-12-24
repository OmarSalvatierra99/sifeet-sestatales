"""
SIFET Estatales - Sistema de Información Financiera de Entes Estatales
Órgano de Fiscalización Superior del Estado de Tlaxcala
"""
import logging
from pathlib import Path
from logging.handlers import RotatingFileHandler
from flask import Flask
from app.config import Config
from app.routes import bp as main_bp


def create_app(config_class=Config):
    """Flask application factory"""
    app = Flask(__name__)
    app.config.from_object(config_class)
    app.secret_key = config_class.SECRET_KEY

    # Logging configuration
    log_dir = Path('log')
    log_dir.mkdir(exist_ok=True)
    handler = RotatingFileHandler('log/sifet-estatales.log', maxBytes=10*1024*1024, backupCount=10)
    handler.setFormatter(logging.Formatter('[%(asctime)s] %(levelname)s: %(message)s'))
    app.logger.addHandler(handler)
    app.logger.setLevel(logging.INFO)
    app.logger.info("SIFET Estatales application started")

    # Register blueprints
    app.register_blueprint(main_bp)

    return app
