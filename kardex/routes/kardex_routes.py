"""Página principal: el Kardex completo (tabla tipo Excel) y el selector de inventario."""
import calendar
from datetime import datetime

from flask import Blueprint, request, redirect, url_for, render_template, session

from ..db import get_db_connection
from ..inventarios import INVENTARIOS, obtener_inventario_actual
from ..logic import es_ip_autorizada, preparar_datos_kardex

kardex_bp = Blueprint('kardex_bp', __name__)


@kardex_bp.route('/cambiar_inventario/<letra>')
def cambiar_inventario(letra):
    letra = letra.upper()
    if letra in INVENTARIOS:
        session['inventario_actual'] = letra
    destino = request.args.get('next') or url_for('kardex_bp.index')
    return redirect(destino)


@kardex_bp.route('/')
def index():
    if es_ip_autorizada():
        return _renderizar_kardex_completo()
    else:
        return redirect(url_for('consultas_bp.consultor'))


def _renderizar_kardex_completo():
    conn = get_db_connection()
    cursor = conn.cursor()

    # Obtener el mes desde la URL
    mes_filtro = request.args.get('mes')
    if not mes_filtro:
        mes_filtro = datetime.now().strftime('%Y-%m')

    cursor.execute('SELECT * FROM materiales WHERE inventario = ? ORDER BY nombre ASC', (obtener_inventario_actual(),))
    materiales_db = cursor.fetchall()

    movimientos_por_material = {}
    for mat in materiales_db:
        cursor.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat['id'],))
        movimientos_por_material[mat['id']] = cursor.fetchall()

    materiales_kardex, alertas_rojas, alertas_amarillas, totales = preparar_datos_kardex(
        materiales_db,
        movimientos_por_material,
        mes_filtro
    )

    hoy = datetime.now()
    try:
        _, ultimo_dia = calendar.monthrange(hoy.year, hoy.month)
        es_fin_de_mes = (ultimo_dia - hoy.day) <= 3
    except Exception:
        es_fin_de_mes = False

    cursor.execute('SELECT * FROM grupos ORDER BY nombre ASC')
    grupos = cursor.fetchall()
    cursor.close()
    conn.close()

    return render_template('index.html', materiales=materiales_kardex, grupos=grupos, mes_filtro=mes_filtro,
                           alertas_rojas=alertas_rojas, alertas_amarillas=alertas_amarillas,
                           es_fin_de_mes=es_fin_de_mes, totales=totales)
