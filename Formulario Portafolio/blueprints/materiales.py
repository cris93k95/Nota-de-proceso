"""
Blueprint: Materiales de Clases
Serves the class materials hub and static HTML files.
"""
import os
import json
from datetime import datetime
from flask import Blueprint, render_template, send_from_directory, session, jsonify, redirect, request

materiales_bp = Blueprint(
    'materiales', __name__,
    url_prefix='/recursos/materiales',
)

MATERIALES_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static', 'recursos', 'materiales')
TOOL_KEY = 'materiales_progress'


def _is_admin():
    return session.get('is_admin', False)


def _admin_email():
    return session.get('admin_email', '')


def _require_admin(f):
    from functools import wraps
    @wraps(f)
    def wrapped(*args, **kwargs):
        if not _is_admin():
            return redirect('/panel')
        return f(*args, **kwargs)
    return wrapped


def _require_admin_api(f):
    from functools import wraps
    @wraps(f)
    def wrapped(*args, **kwargs):
        if not _is_admin():
            return jsonify({'error': 'No autorizado'}), 403
        return f(*args, **kwargs)
    return wrapped


@materiales_bp.route('/')
@_require_admin
def index():
    return render_template('recursos/materiales/index.html')


@materiales_bp.route('/archivo/<path:filepath>')
@_require_admin
def serve_file(filepath):
    return send_from_directory(MATERIALES_DIR, filepath)


# ==================== PROGRESS API ====================

@materiales_bp.route('/api/progress', methods=['GET'])
@_require_admin_api
def get_progress():
    from app import RecursoState
    email = _admin_email()
    row = RecursoState.query.filter_by(admin_email=email, tool=TOOL_KEY).first()
    if row and row.data:
        try:
            return jsonify(json.loads(row.data))
        except Exception:
            pass
    return jsonify({'done': []})


@materiales_bp.route('/api/progress', methods=['POST'])
@_require_admin_api
def save_progress():
    from app import db, RecursoState
    email = _admin_email()
    data = request.get_json(silent=True)
    if not data or 'done' not in data:
        return jsonify({'error': 'Datos inválidos'}), 400

    done_list = data['done']
    if not isinstance(done_list, list):
        return jsonify({'error': 'Formato inválido'}), 400

    payload = json.dumps({'done': done_list}, ensure_ascii=False)
    row = RecursoState.query.filter_by(admin_email=email, tool=TOOL_KEY).first()
    if row:
        row.data = payload
        row.updated_at = datetime.utcnow()
    else:
        row = RecursoState(admin_email=email, tool=TOOL_KEY, data=payload)
        db.session.add(row)
    db.session.commit()
    return jsonify({'ok': True})
