"""Punto de entrada de KARDEX JD. La aplicación en sí (rutas, lógica, base de
datos) vive en el paquete `kardex/`, organizada en blueprints por área
funcional. Ver `kardex/__init__.py` para el application factory."""
from kardex import create_app
from kardex.db import init_db

app = create_app()

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=3000, debug=True)
