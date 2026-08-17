"""Reportes: kardex histórico de un material (por mes o fecha) y búsqueda de
movimientos por orden/documento a través de todos los materiales."""
from datetime import datetime

from flask import Blueprint, request, render_template

from ..db import get_db_connection
from ..inventarios import obtener_inventario_actual

reportes_bp = Blueprint('reportes_bp', __name__)


@reportes_bp.route('/reporte')
def reporte():
    conn = get_db_connection()
    cursor = conn.cursor()
    inventario_actual = obtener_inventario_actual()

    cursor.execute('SELECT * FROM materiales WHERE inventario = ? ORDER BY nombre ASC', (inventario_actual,))
    materiales = cursor.fetchall()

    selected_material_id = request.args.get('material_id', type=int)
    mes_filtro = request.args.get('mes') or ''
    fecha_filtro = request.args.get('fecha') or ''
    orden_busqueda = (request.args.get('orden') or '').strip()

    if fecha_filtro:
        mes_filtro = ''  # la fecha específica tiene prioridad sobre el mes
    elif not mes_filtro:
        mes_filtro = datetime.now().strftime('%Y-%m')

    reporte_datos = None

    if selected_material_id:
        cursor.execute('SELECT * FROM materiales WHERE id = ? AND inventario = ?', (selected_material_id, inventario_actual))
        mat = cursor.fetchone()
        if mat:
            mat_id = mat['id']
            cant_saldo = mat['cantidad_inicial']
            precio_promedio = mat['precio_unitario']
            total_saldo = cant_saldo * precio_promedio

            cursor.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat_id,))
            movimientos = cursor.fetchall()

            if fecha_filtro:
                movs_anteriores = [m for m in movimientos if str(m['fecha']) < fecha_filtro]
                movs_actuales = [m for m in movimientos if str(m['fecha']) == fecha_filtro]
            elif mes_filtro != 'todos':
                movs_anteriores = [m for m in movimientos if str(m['fecha']) < f"{mes_filtro}-01"]
                movs_actuales = [m for m in movimientos if str(m['fecha']).startswith(mes_filtro)]
            else:
                movs_anteriores = []
                movs_actuales = movimientos

            for mov in movs_anteriores:
                if mov['tipo'] == 'entrada':
                    costo_movimiento = mov['cantidad'] * mov['precio_unitario']
                    cant_saldo += mov['cantidad']
                    total_saldo += costo_movimiento
                    if cant_saldo > 0: precio_promedio = total_saldo / cant_saldo
                elif mov['tipo'] == 'salida':
                    costo_movimiento = mov['cantidad'] * precio_promedio
                    cant_saldo -= mov['cantidad']
                    total_saldo -= costo_movimiento

            filas_kardex = []
            # Primera fila: El saldo inicial o anterior según el filtro
            if fecha_filtro:
                titulo_saldo = f'Saldo Anterior (antes del {fecha_filtro})'
            else:
                titulo_saldo = 'Saldo Inicial' if mes_filtro == 'todos' else f'Saldo Anterior ({mes_filtro})'
            filas_kardex.append({
                'fecha': '-', 'detalle': titulo_saldo,
                'ing_cant': '', 'ing_costo': '', 'ing_total': '',
                'sal_cant': '', 'sal_costo': '', 'sal_total': '',
                'saldo_cant': cant_saldo, 'saldo_costo': precio_promedio, 'saldo_total': total_saldo
            })

            for mov in movs_actuales:
                doc_info = ""
                if mov['documento'] and mov['numero_documento']:
                    doc_info = f" ({mov['documento']} #{mov['numero_documento']})"
                elif mov['documento']:
                    doc_info = f" ({mov['documento']})"

                if mov['tipo'] == 'entrada':
                    costo_movimiento = mov['cantidad'] * mov['precio_unitario']
                    cant_saldo += mov['cantidad']
                    total_saldo += costo_movimiento
                    if cant_saldo > 0:
                        precio_promedio = total_saldo / cant_saldo

                    filas_kardex.append({
                        'fecha': mov['fecha'], 'detalle': f"Entrada / Compra{doc_info}",
                        'ing_cant': mov['cantidad'], 'ing_costo': mov['precio_unitario'], 'ing_total': costo_movimiento,
                        'sal_cant': '', 'sal_costo': '', 'sal_total': '',
                        'saldo_cant': cant_saldo, 'saldo_costo': precio_promedio, 'saldo_total': total_saldo
                    })
                elif mov['tipo'] == 'salida':
                    costo_movimiento = mov['cantidad'] * precio_promedio
                    cant_saldo -= mov['cantidad']
                    total_saldo -= costo_movimiento

                    filas_kardex.append({
                        'fecha': mov['fecha'], 'detalle': f"Salida / Egreso{doc_info}",
                        'ing_cant': '', 'ing_costo': '', 'ing_total': '',
                        'sal_cant': mov['cantidad'], 'sal_costo': precio_promedio, 'sal_total': costo_movimiento,
                        'saldo_cant': cant_saldo, 'saldo_costo': precio_promedio, 'saldo_total': total_saldo
                    })

            reporte_datos = {'material': mat, 'filas': filas_kardex}

    resultados_orden = None
    if orden_busqueda:
        cursor.execute('''
            SELECT mov.*, mat.nombre AS material_nombre, mat.codigo AS material_codigo
            FROM movimientos mov
            JOIN materiales mat ON mat.id = mov.material_id
            WHERE mat.inventario = ? AND (mov.numero_documento LIKE ? OR mov.documento LIKE ?)
            ORDER BY mov.fecha ASC, mov.id ASC
        ''', (inventario_actual, f'%{orden_busqueda}%', f'%{orden_busqueda}%'))
        resultados_orden = [dict(m) for m in cursor.fetchall()]
        for r in resultados_orden:
            r['total'] = r['cantidad'] * r['precio_unitario']

    cursor.close()
    conn.close()
    return render_template('reporte.html', materiales=materiales, reporte_datos=reporte_datos,
                           selected_material_id=selected_material_id, mes_filtro=mes_filtro,
                           fecha_filtro=fecha_filtro, orden_busqueda=orden_busqueda,
                           resultados_orden=resultados_orden)
