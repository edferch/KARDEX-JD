"""Panel de administración: login, catálogos (grupos/proveedores/fuentes/IPs)
y gestión de movimientos históricos (editar/eliminar entradas y salidas)."""
import sqlite3

from flask import Blueprint, request, redirect, url_for, render_template, flash, session, jsonify

from ..auth import admin_required, admin_required_form
from ..db import get_db_connection
from ..inventarios import obtener_inventario_actual
from ..logic import movimiento_a_dict

admin_bp = Blueprint('admin_bp', __name__)


@admin_bp.route('/admin', methods=['GET', 'POST'])
def admin():
    # --- SISTEMA DE LOGIN PARA LA PANTALLA DE ADMIN ---
    if not session.get('admin_logged_in'):
        if request.method == 'POST':
            if request.form.get('admin_password') == 'laFabrica1':  # <- Contraseña de administrador
                session['admin_logged_in'] = True
                flash("Acceso concedido.", "success")
                return redirect(url_for('admin_bp.admin'))
            elif request.form.get('admin_password'):
                flash("Error: Contraseña incorrecta.", "error")
        return render_template('admin.html', login_required=True)

    conn = get_db_connection()
    cursor = conn.cursor()

    if request.method == 'POST':
        accion = request.form.get('accion')

        if accion == 'logout':
            session.pop('admin_logged_in', None)
            flash("Sesión de administrador cerrada.", "success")
            return redirect(url_for('kardex_bp.index'))

        if accion == 'grupo':
            try:
                cursor.execute('INSERT INTO grupos (nombre) VALUES (?)', (request.form['nombre_grupo'],))
                flash("Éxito: Grupo agregado correctamente.", "success")
            except sqlite3.IntegrityError:
                conn.rollback()
                flash("Error: El grupo ya existe.", "error")

        elif accion == 'proveedor':
            cursor.execute('INSERT INTO proveedores (nit, nombre) VALUES (?, ?)',
                           (request.form['nit'], request.form['nombre']))
            flash("Éxito: Proveedor agregado correctamente.", "success")

        elif accion == 'fuente':
            try:
                cursor.execute('INSERT INTO fuentes (nombre) VALUES (?)', (request.form['nombre_fuente'],))
                flash("Éxito: Fuente agregada correctamente.", "success")
            except sqlite3.IntegrityError:
                conn.rollback()
                flash("Error: La fuente ya existe.", "error")

        elif accion == 'agregar_ip':
            tipo_ip = request.form.get('tipo_ip', 'kardex')
            if tipo_ip not in ('kardex', 'quote'):
                tipo_ip = 'kardex'
            cursor.execute('INSERT INTO ips_autorizadas (ip_direccion, descripcion, tipo) VALUES (?, ?, ?)',
                           (request.form['nueva_ip'], request.form['desc_ip'], tipo_ip))
            conn.commit()
            flash("IP agregada a la lista blanca.", "success")

        conn.commit()
        return redirect(url_for('admin_bp.admin'))

    cursor.execute('SELECT * FROM grupos ORDER BY nombre ASC')
    grupos = cursor.fetchall()

    cursor.execute('SELECT * FROM proveedores ORDER BY nombre ASC')
    proveedores = cursor.fetchall()

    cursor.execute('SELECT * FROM fuentes ORDER BY nombre ASC')
    fuentes = cursor.fetchall()

    cursor.execute('SELECT id, nombre FROM materiales WHERE inventario = ? ORDER BY nombre ASC', (obtener_inventario_actual(),))
    materiales = cursor.fetchall()

    cursor.execute('SELECT * FROM ips_autorizadas ORDER BY tipo ASC, id DESC')
    ips = cursor.fetchall()

    cursor.close()
    conn.close()
    return render_template('admin.html', grupos=grupos, proveedores=proveedores, fuentes=fuentes, materiales=materiales, ips=ips)


# --- GESTIÓN DE MOVIMIENTOS (ENTRADAS/SALIDAS) DESDE EL ADMIN ---
# Solo accesible con sesión de administrador iniciada (admin_required).

@admin_bp.route('/admin/movimientos')
@admin_required
def admin_listar_movimientos():
    material_id = request.args.get('material_id', type=int)
    tipo = request.args.get('tipo')
    mes = request.args.get('mes')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = '''
        SELECT mov.*, mat.nombre AS material_nombre, mat.unidad AS material_unidad
        FROM movimientos mov
        JOIN materiales mat ON mat.id = mov.material_id
        WHERE mat.inventario = ?
    '''
    params = [obtener_inventario_actual()]
    if material_id:
        query += ' AND mov.material_id = ?'
        params.append(material_id)
    if tipo in ('entrada', 'salida'):
        query += ' AND mov.tipo = ?'
        params.append(tipo)
    if mes:
        query += " AND strftime('%Y-%m', mov.fecha) = ?"
        params.append(mes)
    query += ' ORDER BY mov.fecha DESC, mov.id DESC LIMIT 300'

    cursor.execute(query, params)
    movimientos = [movimiento_a_dict(m) for m in cursor.fetchall()]
    cursor.close()
    conn.close()
    return jsonify({'success': True, 'movimientos': movimientos})


@admin_bp.route('/admin/movimiento/editar', methods=['POST'])
@admin_required
def admin_editar_movimiento():
    data = request.json or {}
    id_mov = data.get('id')

    if not id_mov:
        return jsonify({'success': False, 'error': 'Movimiento no especificado.'})

    try:
        cantidad = float(data.get('cantidad'))
        precio_unitario = float(data.get('precio_unitario'))
    except (TypeError, ValueError):
        return jsonify({'success': False, 'error': 'Cantidad o precio unitario inválidos.'})

    fecha = data.get('fecha') or None
    if not fecha:
        return jsonify({'success': False, 'error': 'La fecha es obligatoria.'})

    documento = data.get('documento', '')
    numero_documento = (data.get('numero_documento') or '').strip()
    fecha_factura = data.get('fecha_factura') or None
    departamento = data.get('departamento', '')
    solicitante = data.get('solicitante', '')

    if not numero_documento:
        return jsonify({'success': False, 'error': 'El correlativo/documento es obligatorio.'})

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # El correlativo debe seguir siendo único, excluyendo el propio registro
        cursor.execute('SELECT id FROM movimientos WHERE numero_documento = ? AND id != ?', (numero_documento, id_mov))
        if cursor.fetchone():
            cursor.close()
            conn.close()
            return jsonify({'success': False, 'error': f"El correlativo '{numero_documento}' ya está en uso por otro movimiento."})

        cursor.execute('''
            UPDATE movimientos
            SET cantidad = ?, precio_unitario = ?, fecha = ?, documento = ?,
                numero_documento = ?, fecha_factura = ?, departamento = ?, solicitante = ?
            WHERE id = ?
        ''', (cantidad, precio_unitario, fecha, documento, numero_documento, fecha_factura, departamento, solicitante, id_mov))
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': str(e)})


@admin_bp.route('/admin/movimiento/eliminar', methods=['POST'])
@admin_required
def admin_eliminar_movimiento():
    data = request.json or {}
    id_mov = data.get('id')
    pin = data.get('pin')

    if pin != '1234':
        return jsonify({'success': False, 'error': 'PIN incorrecto.'})
    if not id_mov:
        return jsonify({'success': False, 'error': 'Movimiento no especificado.'})

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute('DELETE FROM movimientos WHERE id = ?', (id_mov,))
        conn.commit()
        cursor.close()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.rollback()
        cursor.close()
        conn.close()
        return jsonify({'success': False, 'error': str(e)})


# --- ELIMINACIÓN DE CATÁLOGOS (grupos/proveedores/fuentes/IPs) DESDE LA LISTA DEL ADMIN ---

@admin_bp.route('/eliminar_grupo/<int:id>', methods=['POST'])
@admin_required_form
def eliminar_grupo(id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM grupos WHERE id = ?', (id,))
    conn.commit()
    cursor.close()
    conn.close()
    flash("Éxito: Grupo eliminado correctamente.", "success")
    return redirect(url_for('admin_bp.admin'))


@admin_bp.route('/eliminar_proveedor/<int:id>', methods=['POST'])
@admin_required_form
def eliminar_proveedor(id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM proveedores WHERE id = ?', (id,))
    conn.commit()
    cursor.close()
    conn.close()
    flash("Éxito: Proveedor eliminado correctamente.", "success")
    return redirect(url_for('admin_bp.admin'))


@admin_bp.route('/eliminar_fuente/<int:id>', methods=['POST'])
@admin_required_form
def eliminar_fuente(id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM fuentes WHERE id = ?', (id,))
    conn.commit()
    cursor.close()
    conn.close()
    flash("Éxito: Fuente eliminada correctamente.", "success")
    return redirect(url_for('admin_bp.admin'))


@admin_bp.route('/eliminar_ip/<int:id>', methods=['POST'])
@admin_required_form
def eliminar_ip(id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM ips_autorizadas WHERE id = ?', (id,))
    conn.commit()
    cursor.close()
    conn.close()
    flash("Éxito: IP eliminada de la lista blanca.", "success")
    return redirect(url_for('admin_bp.admin'))
