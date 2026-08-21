"""Lógica de negocio del Kardex compartida entre varias rutas
(cálculo de saldos/movimientos, stock actual, etc.), sin depender de ninguna
vista en particular."""
from flask import request

from .db import get_db_connection
from .inventarios import obtener_inventario_actual


def _ip_autorizada_para(tipo):
    ip_cliente = request.remote_addr
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT 1 FROM ips_autorizadas WHERE ip_direccion = ? AND tipo = ?', (ip_cliente, tipo))
    autorizada = cursor.fetchone() is not None
    cursor.close()
    conn.close()
    return autorizada


def ip_autorizada_kardex():
    """IPs autorizadas para ver el Kardex completo (acceso total)."""
    return _ip_autorizada_para('kardex')


def ip_autorizada_quote():
    """IPs autorizadas para ver la pantalla de Quote. Quien tiene acceso al
    Kardex también puede ver Quote."""
    return ip_autorizada_kardex() or _ip_autorizada_para('quote')


def aplicar_movimiento(cant_saldo, precio_promedio, total_saldo, mov):
    """Aplica UN movimiento (entrada o salida) a un saldo corrido, usando costeo
    promedio ponderado (Kardex). Es el único lugar del código donde vive esta
    fórmula: tanto el Kardex del index como el Reporte por material la llaman,
    para que sus cifras nunca puedan divergir entre pantallas.

    Devuelve (cant_saldo, precio_promedio, total_saldo, costo_del_movimiento).
    """
    cantidad = float(mov['cantidad'])
    if mov['tipo'] == 'entrada':
        costo_movimiento = cantidad * float(mov['precio_unitario'])
        cant_saldo += cantidad
        total_saldo += costo_movimiento
        if cant_saldo > 0:
            precio_promedio = total_saldo / cant_saldo
    elif mov['tipo'] == 'salida':
        costo_movimiento = cantidad * precio_promedio
        cant_saldo -= cantidad
        total_saldo -= costo_movimiento
    else:
        costo_movimiento = 0.0
    return cant_saldo, precio_promedio, total_saldo, costo_movimiento


def preparar_datos_kardex(materiales_db, movimientos_por_material, mes_filtro):
    materiales_kardex = []
    alertas_rojas = []
    alertas_amarillas = []

    totales = {
        'ini_cant': 0.0, 'ini_total': 0.0,
        'ing_cant': 0.0, 'ing_total': 0.0,
        'sal_cant': 0.0, 'sal_total': 0.0,
        'fin_cant': 0.0, 'fin_total': 0.0
    }

    for mat in materiales_db:
        mat_id = mat['id']
        cant_saldo = float(mat['cantidad_inicial'])
        precio_promedio = float(mat['precio_unitario'])
        total_saldo = cant_saldo * precio_promedio
        movimientos = movimientos_por_material.get(mat_id, [])

        if mes_filtro != 'todos':
            movs_anteriores = [m for m in movimientos if str(m['fecha']) < f"{mes_filtro}-01"]
            movs_actuales = [m for m in movimientos if str(m['fecha']).startswith(mes_filtro)]
        else:
            movs_anteriores = []
            movs_actuales = movimientos

        for mov in movs_anteriores:
            cant_saldo, precio_promedio, total_saldo, _ = aplicar_movimiento(cant_saldo, precio_promedio, total_saldo, mov)

        ini_cant, ini_costo, ini_total = cant_saldo, precio_promedio, total_saldo

        acum_ingreso_cant, acum_ingreso_total = 0.0, 0.0
        acum_salida_cant, acum_salida_total = 0.0, 0.0

        for mov in movs_actuales:
            cant_saldo, precio_promedio, total_saldo, costo_movimiento = aplicar_movimiento(cant_saldo, precio_promedio, total_saldo, mov)
            if mov['tipo'] == 'entrada':
                acum_ingreso_cant += float(mov['cantidad'])
                acum_ingreso_total += costo_movimiento
            elif mov['tipo'] == 'salida':
                acum_salida_cant += float(mov['cantidad'])
                acum_salida_total += costo_movimiento

        avg_ingreso = acum_ingreso_total / acum_ingreso_cant if acum_ingreso_cant > 0 else 0
        avg_salida = acum_salida_total / acum_salida_cant if acum_salida_cant > 0 else 0

        # La vista del personal (index) ya no identifica materiales por nombre: se usa
        # el código, y si el material no tiene código asignado, su descripción.
        etiqueta_alerta = dict(mat).get('codigo') or dict(mat).get('descripcion') or 'Material sin código'
        if cant_saldo < 2:
            alertas_rojas.append({'nombre': etiqueta_alerta, 'stock': cant_saldo})
        elif cant_saldo < 5:
            alertas_amarillas.append({'nombre': etiqueta_alerta, 'stock': cant_saldo})

        materiales_kardex.append({
            'id': mat['id'],
            'nombre': mat['nombre'],
            'codigo': dict(mat).get('codigo', ''),
            'descripcion': dict(mat).get('descripcion', ''),
            'drive_link': mat['drive_link'],
            'costo_link': dict(mat).get('costo_link', ''),
            'tipo_material': mat['tipo_material'],
            'unidad': mat['unidad'],
            'ini_cant': ini_cant, 'ini_costo': ini_costo, 'ini_total': ini_total,
            'ing_cant': acum_ingreso_cant, 'ing_costo': avg_ingreso, 'ing_total': acum_ingreso_total,
            'sal_cant': acum_salida_cant, 'sal_costo': avg_salida, 'sal_total': acum_salida_total,
            'fin_cant': cant_saldo, 'fin_costo': precio_promedio, 'fin_total': total_saldo
        })

        totales['ini_cant'] += ini_cant
        totales['ini_total'] += ini_total
        totales['ing_cant'] += acum_ingreso_cant
        totales['ing_total'] += acum_ingreso_total
        totales['sal_cant'] += acum_salida_cant
        totales['sal_total'] += acum_salida_total
        totales['fin_cant'] += cant_saldo
        totales['fin_total'] += total_saldo

    return materiales_kardex, alertas_rojas, alertas_amarillas, totales


def obtener_materiales_con_stock(cursor):
    """Calcula stock actual, costo promedio actual y fecha de última entrada por material.

    Reutiliza preparar_datos_kardex (con mes_filtro='todos') para que el costo promedio
    y el stock coincidan siempre con los que muestra el Kardex del index: antes esta
    función promediaba el costo de TODO lo que había entrado alguna vez sin descontar
    las salidas, lo que daba un número distinto al costo promedio ponderado real en
    cuanto un material tenía más de un precio de compra y alguna salida.
    """
    cursor.execute('SELECT * FROM materiales WHERE inventario = ? ORDER BY nombre ASC', (obtener_inventario_actual(),))
    materiales_raw = cursor.fetchall()

    movimientos_por_material = {}
    for mat in materiales_raw:
        cursor.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat['id'],))
        movimientos_por_material[mat['id']] = cursor.fetchall()

    materiales_kardex, _, _, _ = preparar_datos_kardex(materiales_raw, movimientos_por_material, 'todos')
    kardex_por_id = {k['id']: k for k in materiales_kardex}

    materiales = []
    for mat in materiales_raw:
        m = dict(mat)
        kardex = kardex_por_id[m['id']]

        m['stock_actual'] = kardex['fin_cant']
        m['costo_promedio_actual'] = kardex['fin_costo']

        cursor.execute("SELECT MAX(fecha) as ultima_entrada FROM movimientos WHERE material_id = ? AND tipo = 'entrada'", (m['id'],))
        m['ultima_entrada'] = cursor.fetchone()['ultima_entrada']

        materiales.append(m)

    return materiales


def movimiento_a_dict(m):
    """Convierte una fila de movimientos (sqlite3.Row) a un dict serializable a JSON."""
    d = dict(m)
    d['cantidad'] = float(d['cantidad']) if d.get('cantidad') is not None else 0
    d['precio_unitario'] = float(d['precio_unitario']) if d.get('precio_unitario') is not None else 0
    d['fecha'] = str(d['fecha']) if d.get('fecha') else ''
    d['fecha_factura'] = str(d['fecha_factura']) if d.get('fecha_factura') else ''
    return d
