"""Application factory de KARDEX JD.

La app se organiza en blueprints por área funcional (kardex, inventario,
movimientos, reportes, excel, admin, consultas). `app.py`, en la raíz del
proyecto, es el único punto de entrada: crea la app con `create_app()` y la
ejecuta.
"""
import os

from flask import Flask, request, session

from .inventarios import INVENTARIOS, obtener_inventario_actual

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Endpoints (ya con el prefijo de su blueprint) que NO deben cerrar la sesión
# de administrador al visitarlos: son llamadas AJAX de fondo hechas desde la
# propia pantalla de admin, o el cambio de inventario que regresa al mismo lugar.
_RUTAS_EXENTAS_DE_CIERRE_ADMIN = {
    'static',
    'admin_bp.admin',
    'admin_bp.admin_listar_movimientos',
    'admin_bp.admin_editar_movimiento',
    'admin_bp.admin_eliminar_movimiento',
    'admin_bp.eliminar_grupo',
    'admin_bp.eliminar_proveedor',
    'admin_bp.eliminar_fuente',
    'admin_bp.eliminar_ip',
    'kardex_bp.cambiar_inventario',
}


def create_app():
    app = Flask(
        __name__,
        template_folder=os.path.join(BASE_DIR, 'templates'),
        static_folder=os.path.join(BASE_DIR, 'static'),
    )
    # Clave secreta necesaria para los mensajes de éxito/error (flash) y la sesión.
    app.secret_key = 'mi_clave_secreta_kardex'

    from .routes.kardex_routes import kardex_bp
    from .routes.inventario_routes import inventario_bp
    from .routes.movimientos_routes import movimientos_bp
    from .routes.reportes_routes import reportes_bp
    from .routes.excel_routes import excel_bp
    from .routes.admin_routes import admin_bp
    from .routes.consultas_routes import consultas_bp

    app.register_blueprint(kardex_bp)
    app.register_blueprint(inventario_bp)
    app.register_blueprint(movimientos_bp)
    app.register_blueprint(reportes_bp)
    app.register_blueprint(excel_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(consultas_bp)

    @app.context_processor
    def inyectar_inventario_actual():
        return {'inventario_actual': obtener_inventario_actual(), 'inventarios_disponibles': INVENTARIOS}

    @app.before_request
    def cerrar_sesion_admin_al_salir():
        if session.get('admin_logged_in') and request.endpoint not in _RUTAS_EXENTAS_DE_CIERRE_ADMIN:
            session.pop('admin_logged_in', None)

    return app
