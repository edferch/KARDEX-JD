"""Protección de rutas de administrador (misma sesión que usa la pantalla /admin)."""
from functools import wraps
from flask import session, jsonify, flash, redirect, url_for


def admin_required(f):
    """Para endpoints AJAX: si no hay sesión de administrador, responde JSON 403."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('admin_logged_in'):
            return jsonify({'success': False, 'error': 'No autorizado. Debes iniciar sesión como administrador.'}), 403
        return f(*args, **kwargs)
    return decorated


def admin_required_form(f):
    """Para rutas que reciben un POST de un <form> normal (no AJAX) dentro de /admin:
    si no hay sesión de administrador, redirige al login en vez de responder JSON."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('admin_logged_in'):
            flash("Debes iniciar sesión como administrador.", "error")
            return redirect(url_for('admin_bp.admin'))
        return f(*args, **kwargs)
    return decorated
