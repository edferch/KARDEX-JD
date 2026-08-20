"""Registro de movimientos de inventario: entradas (factura), devoluciones y salidas."""
from datetime import datetime

from flask import Blueprint, request, redirect, url_for, flash, jsonify

from ..db import get_db_connection

movimientos_bp = Blueprint('movimientos_bp', __name__)


@movimientos_bp.route('/agregar_entrada', methods=['POST'])
def agregar_entrada():
    # Entrada de inventario respaldada SIEMPRE por una factura de compra.
    if request.method == 'POST':
        material_id = int(request.form['material_id'])
        cantidad = float(request.form['cantidad'])

        precio_str = request.form.get('precio', '').strip()
        precio = float(precio_str) if precio_str else 0.0

        fecha = request.form.get('fecha')
        fecha_factura = request.form.get('fecha_factura', '')
        numero_documento = request.form.get('numero_documento', '').strip()

        if not fecha:
            fecha = datetime.now().strftime('%Y-%m-%d')

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO movimientos (material_id, tipo, cantidad, precio_unitario, fecha, documento, numero_documento, fecha_factura)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (material_id, 'entrada', cantidad, precio, fecha, 'Factura', numero_documento, fecha_factura))

        conn.commit()
        cursor.close()
        conn.close()

        flash("Éxito: Factura registrada correctamente.", "success")
        return redirect(url_for('kardex_bp.index'))


@movimientos_bp.route('/agregar_devolucion', methods=['POST'])
def agregar_devolucion():
    # Proceso independiente: devuelve al stock unidades de una salida previa,
    # usando el costo exacto con el que salieron (no requiere factura nueva).
    if request.method == 'POST':
        material_id = int(request.form['material_id'])
        cantidad = float(request.form['cantidad'])
        fecha = request.form.get('fecha')
        numero_orden = request.form.get('numero_documento', '').strip()

        if not fecha:
            fecha = datetime.now().strftime('%Y-%m-%d')

        conn = get_db_connection()
        cursor = conn.cursor()

        # Buscamos la salida cuya Orden (documento) coincide con la ingresada
        cursor.execute('''
            SELECT precio_unitario FROM movimientos
            WHERE material_id = ? AND tipo = 'salida' AND documento = ?
            ORDER BY id DESC
            LIMIT 1
        ''', (material_id, numero_orden))
        salida = cursor.fetchone()

        if not salida:
            cursor.close()
            conn.close()
            flash(f"Error: No se encontró una salida asociada a la Orden '{numero_orden}'.", "error")
            return redirect(url_for('kardex_bp.index'))

        precio = salida['precio_unitario']

        cursor.execute('''
            INSERT INTO movimientos (material_id, tipo, cantidad, precio_unitario, fecha, documento, numero_documento)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (material_id, 'entrada', cantidad, precio, fecha, 'Devolución', numero_orden))

        conn.commit()
        cursor.close()
        conn.close()

        flash("Éxito: Devolución registrada correctamente.", "success")
        return redirect(url_for('kardex_bp.index'))


@movimientos_bp.route('/api/precio_devolucion')
def api_precio_devolucion():
    material_id = request.args.get('material_id', type=int)
    orden = request.args.get('orden', '')

    if not material_id or not orden:
        return jsonify({'success': False, 'error': 'Faltan datos'})

    conn = get_db_connection()
    cursor = conn.cursor()
    # Busca el precio original basado en la orden (documento)
    cursor.execute('''
        SELECT precio_unitario FROM movimientos
        WHERE material_id = ? AND tipo = 'salida' AND documento = ?
        ORDER BY id DESC
        LIMIT 1
    ''', (material_id, orden))
    salida = cursor.fetchone()
    cursor.close()
    conn.close()

    if salida:
        return jsonify({'success': True, 'precio': salida['precio_unitario']})
    else:
        return jsonify({'success': False, 'error': 'No encontrada'})


@movimientos_bp.route('/agregar_salida', methods=['POST'])
def agregar_salida():
    if request.method == 'POST':
        material_id = int(request.form['material_id'])
        cantidad_a_sacar = float(request.form['cantidad'])
        fecha = request.form.get('fecha')
        # Ahora el documento es fijo "Orden"
        documento = request.form.get('documento', 'Orden')
        # El número de documento es el correlativo único
        numero_documento = request.form.get('numero_documento', '').strip()
        departamento = request.form.get('departamento', '')
        solicitante = request.form.get('solicitante', '')

        if not fecha:
            fecha = datetime.now().strftime('%Y-%m-%d')

        if not numero_documento:
            flash("Error: El correlativo es obligatorio.", "error")
            return redirect(url_for('kardex_bp.index'))

        conn = get_db_connection()
        cursor = conn.cursor()

        # 2. Validar existencias actuales
        cursor.execute('SELECT * FROM materiales WHERE id = ?', (material_id,))
        material = cursor.fetchone()
        cant_actual = material['cantidad_inicial']
        total_actual = material['cantidad_inicial'] * material['precio_unitario']
        precio_promedio = material['precio_unitario']

        cursor.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (material_id,))
        movimientos = cursor.fetchall()
        for mov in movimientos:
            if mov['tipo'] == 'entrada':
                cant_actual += mov['cantidad']
                total_actual += (mov['cantidad'] * mov['precio_unitario'])
                if cant_actual > 0: precio_promedio = total_actual / cant_actual
            elif mov['tipo'] == 'salida':
                cant_actual -= mov['cantidad']
                total_actual -= (mov['cantidad'] * precio_promedio)

        # BLOQUEO SI NO HAY STOCK SUFICIENTE
        if cantidad_a_sacar > cant_actual:
            cursor.close()
            conn.close()
            flash(f"Error: No puedes sacar {cantidad_a_sacar} unidades. Solo hay {cant_actual} disponibles del material '{material['nombre']}'.", "error")
            return redirect(url_for('kardex_bp.index'))

        # Si hay stock y correlativo único, registrar la salida
        cursor.execute('''
            INSERT INTO movimientos (material_id, tipo, cantidad, precio_unitario, fecha, documento, numero_documento, departamento, solicitante)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (material_id, 'salida', cantidad_a_sacar, precio_promedio, fecha, documento, numero_documento, departamento, solicitante))

        conn.commit()
        cursor.close()
        conn.close()

        flash(f"Éxito: Salida registrada correctamente con correlativo {numero_documento}.", "success")
        return redirect(url_for('kardex_bp.index'))
