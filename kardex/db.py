"""Conexión a la base de datos y esquema (SQLite)."""
import sqlite3

DB_PATH = 'kardex.db'


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


def init_db():
    conn = get_db_connection()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS grupos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL UNIQUE
        );
        CREATE TABLE IF NOT EXISTS fuentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL UNIQUE
        );
        CREATE TABLE IF NOT EXISTS proveedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT, nit TEXT, nombre TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS materiales (
            id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL, codigo TEXT, descripcion TEXT,
            tipo_material TEXT, numero_metrico TEXT, origen TEXT, empresa TEXT,
            presentacion TEXT, unidad TEXT, cantidad_inicial NUMERIC DEFAULT 0,
            precio_unitario NUMERIC DEFAULT 0.0, fuente TEXT, drive_link TEXT, costo_link TEXT,
            inventario TEXT NOT NULL DEFAULT 'A'
        );
        CREATE TABLE IF NOT EXISTS movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, material_id INTEGER REFERENCES materiales(id),
            tipo TEXT, cantidad NUMERIC, precio_unitario NUMERIC, fecha DATE,
            documento TEXT, numero_documento TEXT, fecha_factura DATE,
            departamento TEXT, solicitante TEXT
        );
        CREATE TABLE IF NOT EXISTS ips_autorizadas (
            id INTEGER PRIMARY KEY AUTOINCREMENT, ip_direccion TEXT NOT NULL, descripcion TEXT
        );
    ''')

    columnas_materiales = [fila['name'] for fila in conn.execute('PRAGMA table_info(materiales)').fetchall()]
    if 'codigo' not in columnas_materiales:
        conn.execute('ALTER TABLE materiales ADD COLUMN codigo TEXT')
    if 'inventario' not in columnas_materiales:
        conn.execute("ALTER TABLE materiales ADD COLUMN inventario TEXT NOT NULL DEFAULT 'A'")

    conn.commit()
    conn.close()
