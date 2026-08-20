"""Gestión de materiales del inventario activo, y las entidades de referencia
(grupos y proveedores) que se agregan/editan al vuelo desde esa misma pantalla."""
import sqlite3

from flask import Blueprint, request, redirect, url_for, render_template, flash, jsonify

from ..db import get_db_connection
from ..inventarios import obtener_inventario_actual
from ..logic import obtener_materiales_con_stock

inventario_bp = Blueprint('inventario_bp', __name__)


@inventario_bp.route('/inventario', methods=['GET', 'POST'])
def inventario():
    if request.method == 'POST':
        nombre = request.form['nombre']
        codigo = request.form.get('codigo', '')
        descripcion = request.form.get('descripcion', '')
        tipo_material = request.form['tipo_material']
        numero_metrico = request.form['numero_metrico']
        origen = request.form['origen']
        empresa = request.form['empresa']
        presentacion = request.form['presentacion']
        unidad = request.form['unidad']
        cantidad_inicial = float(request.form['cantidad_inicial'])
        precio_unitario = float(request.form['precio_unitario'])
        fuente = request.form.get('fuente', '')
        drive_link = request.form.get('drive_link', '')

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO materiales (nombre, codigo, descripcion, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente, drive_link, inventario)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (nombre, codigo, descripcion, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente, drive_link, obtener_inventario_actual()))
        conn.commit()
        cursor.close()
        conn.close()

        flash(f"Éxito: Material agregado correctamente al Inventario {obtener_inventario_actual()}.", "success")
        return redirect(url_for('inventario_bp.inventario'))

    conn = get_db_connection()
    cursor = conn.cursor()

    materiales = obtener_materiales_con_stock(cursor)

    cursor.execute('SELECT * FROM grupos ORDER BY nombre ASC')
    grupos = cursor.fetchall()
    cursor.execute('SELECT * FROM proveedores ORDER BY nombre ASC')
    proveedores = cursor.fetchall()
    cursor.execute('SELECT * FROM fuentes ORDER BY nombre ASC')
    fuentes = cursor.fetchall()

    cursor.close()
    conn.close()

    return render_template('inventario.html', materiales=materiales, grupos=grupos, proveedores=proveedores, fuentes=fuentes)


@inventario_bp.route('/editar_material', methods=['POST'])
def editar_material():
    if request.method == 'POST':
        id_material = int(request.form['id'])
        nombre = request.form['nombre']
        codigo = request.form.get('codigo', '')
        descripcion = request.form.get('descripcion', '')
        tipo_material = request.form['tipo_material']
        numero_metrico = request.form['numero_metrico']
        origen = request.form['origen']
        empresa = request.form['empresa']
        presentacion = request.form['presentacion']
        unidad = request.form['unidad']
        cantidad_inicial = float(request.form['cantidad_inicial'])
        precio_unitario = float(request.form['precio_unitario'])
        fuente = request.form.get('fuente', '')
        drive_link = request.form.get('drive_link', '')

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE materiales
            SET nombre = ?, codigo = ?, descripcion = ?, tipo_material = ?, numero_metrico = ?, origen = ?, empresa = ?, presentacion = ?, unidad = ?, cantidad_inicial = ?, precio_unitario = ?, fuente = ?, drive_link = ?
            WHERE id = ?
        ''', (nombre, codigo, descripcion, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente, drive_link, id_material))
        conn.commit()
        cursor.close()
        conn.close()

        flash("Éxito: Material actualizado correctamente.", "success")
        return redirect(url_for('inventario_bp.inventario'))


@inventario_bp.route('/eliminar_material/<int:id>', methods=['POST'])
def eliminar_material(id):
    if request.method == 'POST':
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM movimientos WHERE material_id = ?', (id,))
        cursor.execute('DELETE FROM materiales WHERE id = ?', (id,))
        conn.commit()
        cursor.close()
        conn.close()
        flash("Éxito: Material eliminado correctamente.", "success")
        return redirect(url_for('inventario_bp.inventario'))


@inventario_bp.route('/actualizar_vinculo_ajax', methods=['POST'])
def actualizar_vinculo_ajax():
    data = request.json
    material_id = data.get('material_id')
    link = data.get('link', '')
    tipo = data.get('tipo', 'nombre')  # 'nombre' o 'costo'

    if not material_id:
        return jsonify({'success': False, 'error': 'ID de material no proporcionado'})

    columna = 'drive_link' if tipo == 'nombre' else 'costo_link'

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(f'UPDATE materiales SET {columna} = ? WHERE id = ?', (link, material_id))
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': str(e)})


# --- GRUPOS (categorías): alta/edición/baja rápida desde Inventario ---

@inventario_bp.route('/agregar_grupo_ajax', methods=['POST'])
def agregar_grupo_ajax():
    nombre = request.json.get('nombre')
    if not nombre:
        return jsonify({'success': False, 'error': 'El nombre está vacío'})

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute('INSERT INTO grupos (nombre) VALUES (?) RETURNING id', (nombre,))
        nuevo_id = cursor.fetchone()[0]
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True, 'id': nuevo_id, 'nombre': nombre})
    except sqlite3.IntegrityError:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': 'El grupo ya existe'})


@inventario_bp.route('/editar_grupo_ajax', methods=['POST'])
def editar_grupo_ajax():
    data = request.json
    id = data.get('id')
    nombre = data.get('nombre')
    nombre_viejo = data.get('nombre_viejo')

    if not nombre:
        return jsonify({'success': False, 'error': 'El nombre está vacío'})

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Actualizar grupo
        cursor.execute('UPDATE grupos SET nombre = ? WHERE id = ?', (nombre, id))
        # Actualizar todos los materiales que usaban este grupo al nuevo nombre
        if nombre != nombre_viejo:
            cursor.execute('UPDATE materiales SET tipo_material = ? WHERE tipo_material = ?', (nombre, nombre_viejo))
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True})
    except sqlite3.IntegrityError:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': 'El grupo ya existe'})
    except Exception as e:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': str(e)})


@inventario_bp.route('/eliminar_grupo_ajax', methods=['POST'])
def eliminar_grupo_ajax():
    data = request.json
    id = data.get('id')
    pin = data.get('pin')

    if pin != '1234':
        return jsonify({'success': False, 'error': 'PIN incorrecto'})

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute('DELETE FROM grupos WHERE id = ?', (id,))
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': str(e)})


# --- PROVEEDORES: alta/edición/baja rápida desde Inventario ---

@inventario_bp.route('/agregar_proveedor_ajax', methods=['POST'])
def agregar_proveedor_ajax():
    nit = request.json.get('nit', '')
    nombre = request.json.get('nombre')
    if not nombre:
        return jsonify({'success': False, 'error': 'El nombre está vacío'})

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute('INSERT INTO proveedores (nit, nombre) VALUES (?, ?) RETURNING id', (nit, nombre))
        nuevo_id = cursor.fetchone()[0]
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True, 'id': nuevo_id, 'nombre': nombre})
    except Exception as e:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': str(e)})


@inventario_bp.route('/editar_proveedor_ajax', methods=['POST'])
def editar_proveedor_ajax():
    data = request.json
    id = data.get('id')
    nit = data.get('nit', '')
    nombre = data.get('nombre')
    nombre_viejo = data.get('nombre_viejo')

    if not nombre:
        return jsonify({'success': False, 'error': 'El nombre está vacío'})

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Actualizar proveedor
        cursor.execute('UPDATE proveedores SET nit = ?, nombre = ? WHERE id = ?', (nit, nombre, id))
        # Actualizar todos los materiales que usaban este proveedor al nuevo nombre
        if nombre != nombre_viejo:
            cursor.execute('UPDATE materiales SET empresa = ? WHERE empresa = ?', (nombre, nombre_viejo))
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': str(e)})


@inventario_bp.route('/eliminar_proveedor_ajax', methods=['POST'])
def eliminar_proveedor_ajax():
    data = request.json
    id = data.get('id')
    pin = data.get('pin')

    if pin != '1234':
        return jsonify({'success': False, 'error': 'PIN incorrecto'})

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute('DELETE FROM proveedores WHERE id = ?', (id,))
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': str(e)})
