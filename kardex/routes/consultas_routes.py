"""Vistas de solo consulta: Quote (para cotizaciones) y Consultor de stock
(pantalla pública para IPs no autorizadas)."""
from flask import Blueprint, render_template

from ..db import get_db_connection
from ..inventarios import obtener_inventario_actual
from ..logic import obtener_materiales_con_stock

consultas_bp = Blueprint('consultas_bp', __name__)


@consultas_bp.route('/quote')
def quote():
    conn = get_db_connection()
    cursor = conn.cursor()
    materiales = obtener_materiales_con_stock(cursor)
    cursor.close()
    conn.close()
    return render_template('quote.html', materiales=materiales)


@consultas_bp.route('/consultor')
def consultor():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('SELECT * FROM materiales WHERE inventario = ? ORDER BY nombre ASC', (obtener_inventario_actual(),))
    materiales_db = cursor.fetchall()
    stock_materiales = []

    for mat in materiales_db:
        mat_id = mat['id']
        cant_saldo = mat['cantidad_inicial']
        cursor.execute('SELECT tipo, cantidad FROM movimientos WHERE material_id = ?', (mat_id,))
        movimientos = cursor.fetchall()

        for mov in movimientos:
            if mov['tipo'] == 'entrada':
                cant_saldo += mov['cantidad']
            elif mov['tipo'] == 'salida':
                cant_saldo -= mov['cantidad']

        # Convertir la fila de la base de datos a un diccionario normal.
        material_info = dict(mat)
        # Añadir el stock calculado a este diccionario.
        material_info['stock'] = cant_saldo

        stock_materiales.append(material_info)

    cursor.close()
    conn.close()
    return render_template('consultor.html', materiales=stock_materiales)
