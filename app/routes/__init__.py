"""
Routes for SIFET Estatales
"""
from flask import jsonify, Blueprint

bp = Blueprint('main', __name__)


@bp.route("/")
def index():
    """Página principal"""
    return jsonify({
        "app": "SIFET Estatales",
        "status": "active",
        "message": "Sistema de Información Financiera de Entes Estatales"
    })


@bp.route("/health")
def health():
    """Health check endpoint"""
    return jsonify({"status": "healthy", "service": "SIFET Estatales"}), 200
