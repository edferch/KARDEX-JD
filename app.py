import sqlite3
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, Response, session
from datetime import datetime
import calendar
import csv
from io import StringIO, BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = Flask(__name__)
# Clave secreta necesaria para los mensajes de éxito/error (flash)
app.secret_key = 'mi_clave_secreta_kardex'

# --- FUNCIÓN: Inicializar la base de datos ---
def inicializar_db():
    conn = sqlite3.connect('kardex.db')
    cursor = conn.cursor()

    # Tabla de Materiales
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS materiales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            tipo_material TEXT,
            numero_metrico TEXT,
            origen TEXT,
            empresa TEXT,
            presentacion TEXT,
            unidad TEXT,
            cantidad_inicial INTEGER DEFAULT 0,
            precio_unitario REAL DEFAULT 0.0
        )
    ''')

    # Tabla de Movimientos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            material_id INTEGER,
            tipo TEXT,
            cantidad INTEGER,
            precio_unitario REAL,
            fecha DATE,
            FOREIGN KEY(material_id) REFERENCES materiales(id)
        )
    ''')

    # Agregar columnas documento a movimientos si no existen (Actualización segura)
    try:
        cursor.execute('ALTER TABLE movimientos ADD COLUMN documento TEXT')
        cursor.execute('ALTER TABLE movimientos ADD COLUMN numero_documento TEXT')
    except sqlite3.OperationalError:
        pass

    # Agregar columna de fecha de factura a movimientos si no existe
    try:
        cursor.execute('ALTER TABLE movimientos ADD COLUMN fecha_factura DATE')
    except sqlite3.OperationalError:
        pass

    # Tabla de Proveedores
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS proveedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nit TEXT,
            nombre TEXT NOT NULL
        )
    ''')

    # Tabla de Fuentes (Quien paga)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fuentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL UNIQUE
        )
    ''')
    
    # Agregar la columna fuente a la tabla materiales si no existe (Actualización segura)
    try:
        cursor.execute('ALTER TABLE materiales ADD COLUMN fuente TEXT')
    except sqlite3.OperationalError:
        pass

    # Agregar la columna descripcion a la tabla materiales si no existe
    try:
        cursor.execute('ALTER TABLE materiales ADD COLUMN descripcion TEXT')
    except sqlite3.OperationalError:
        pass

    # Agregar la columna de hipervínculo a la tabla materiales si no existe
    try:
        cursor.execute('ALTER TABLE materiales ADD COLUMN drive_link TEXT')
    except sqlite3.OperationalError:
        pass

    # Tabla de Grupos (Categorías/Fuentes)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS grupos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL UNIQUE
        )
    ''')
    

    conn.commit()
    conn.close()

def get_db_connection():
    conn = sqlite3.connect('kardex.db')
    conn.row_factory = sqlite3.Row 
    return conn

# --- RUTAS DE FLASK ---

@app.route('/')
def index():
    conn = get_db_connection()
    
    # Obtener el mes desde la URL, si no hay, usar el mes actual
    mes_filtro = request.args.get('mes')
    if not mes_filtro:
        mes_filtro = datetime.now().strftime('%Y-%m')
        
    materiales_db = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    materiales_kardex = []
    
    alertas_rojas = []
    alertas_amarillas = []
    
    # Detectar si estamos a fin de mes (últimos 3 días) para sugerir descarga
    hoy = datetime.now()
    try:
        _, ultimo_dia = calendar.monthrange(hoy.year, hoy.month)
        es_fin_de_mes = (ultimo_dia - hoy.day) <= 3
    except Exception:
        es_fin_de_mes = False
        
    # Lógica de Costo Promedio Ponderado
    for mat in materiales_db:
        mat_id = mat['id']
        
        cant_saldo = mat['cantidad_inicial']
        precio_promedio = mat['precio_unitario']
        total_saldo = cant_saldo * precio_promedio
        
        movimientos = conn.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat_id,)).fetchall()
        
        # --- DIVIDIR MOVIMIENTOS: ANTERIORES VS ACTUALES ---
        if mes_filtro != 'todos':
            movs_anteriores = [m for m in movimientos if m['fecha'] < f"{mes_filtro}-01"]
            movs_actuales = [m for m in movimientos if m['fecha'].startswith(mes_filtro)]
        else:
            movs_anteriores = []
            movs_actuales = movimientos
            
        # 1. Procesar históricos para obtener el "Saldo Inicial" del mes seleccionado
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
        
        ini_cant = cant_saldo
        ini_costo = precio_promedio
        ini_total = total_saldo
        
        # 2. Procesar únicamente los movimientos del mes seleccionado
        acum_ingreso_cant = 0
        acum_ingreso_total = 0
        acum_salida_cant = 0
        acum_salida_total = 0
        
        for mov in movs_actuales:
            if mov['tipo'] == 'entrada':
                costo_movimiento = mov['cantidad'] * mov['precio_unitario']
                cant_saldo += mov['cantidad']
                total_saldo += costo_movimiento
                acum_ingreso_cant += mov['cantidad']
                acum_ingreso_total += costo_movimiento
                if cant_saldo > 0:
                    precio_promedio = total_saldo / cant_saldo
                    
            elif mov['tipo'] == 'salida':
                costo_movimiento = mov['cantidad'] * precio_promedio
                cant_saldo -= mov['cantidad']
                total_saldo -= costo_movimiento
                acum_salida_cant += mov['cantidad']
                acum_salida_total += costo_movimiento

        avg_ingreso = acum_ingreso_total / acum_ingreso_cant if acum_ingreso_cant > 0 else 0
        avg_salida = acum_salida_total / acum_salida_cant if acum_salida_cant > 0 else 0

        # --- LÓGICA DE ALERTAS ---
        if cant_saldo < 2:
            alertas_rojas.append({'nombre': mat['nombre'], 'stock': cant_saldo})
        elif cant_saldo < 5:
            alertas_amarillas.append({'nombre': mat['nombre'], 'stock': cant_saldo})

        materiales_kardex.append({
            'id': mat['id'],
            'nombre': mat['nombre'],
            'drive_link': mat['drive_link'],
            'tipo_material': mat['tipo_material'],
            'unidad': mat['unidad'],
            'ini_cant': ini_cant,
            'ini_costo': ini_costo,
            'ini_total': ini_total,
            'ing_cant': acum_ingreso_cant,
            'ing_costo': avg_ingreso,
            'ing_total': acum_ingreso_total,
            'sal_cant': acum_salida_cant,
            'sal_costo': avg_salida,
            'sal_total': acum_salida_total,
            'fin_cant': cant_saldo,
            'fin_costo': precio_promedio,
            'fin_total': total_saldo
        })
        
    grupos = conn.execute('SELECT * FROM grupos ORDER BY nombre ASC').fetchall()
    conn.close()
    
    return render_template('index.html', materiales=materiales_kardex, grupos=grupos, mes_filtro=mes_filtro, alertas_rojas=alertas_rojas, alertas_amarillas=alertas_amarillas, es_fin_de_mes=es_fin_de_mes)

@app.route('/inventario', methods=['GET', 'POST'])
def inventario():
    if request.method == 'POST':
        nombre = request.form['nombre']
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
        conn.execute('''
            INSERT INTO materiales (nombre, descripcion, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente, drive_link)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (nombre, descripcion, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente, drive_link))
        conn.commit()
        conn.close()

        flash("Éxito: Material agregado correctamente al Inventario.", "success")
        return redirect(url_for('inventario'))

    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    grupos = conn.execute('SELECT * FROM grupos ORDER BY nombre ASC').fetchall()
    proveedores = conn.execute('SELECT * FROM proveedores ORDER BY nombre ASC').fetchall()
    fuentes = conn.execute('SELECT * FROM fuentes ORDER BY nombre ASC').fetchall()
    conn.close()
    return render_template('inventario.html', materiales=materiales, grupos=grupos, proveedores=proveedores, fuentes=fuentes)

@app.route('/editar_material', methods=['POST'])
def editar_material():
    if request.method == 'POST':
        id_material = int(request.form['id'])
        nombre = request.form['nombre']
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
        conn.execute('''
            UPDATE materiales 
            SET nombre = ?, descripcion = ?, tipo_material = ?, numero_metrico = ?, origen = ?, empresa = ?, presentacion = ?, unidad = ?, cantidad_inicial = ?, precio_unitario = ?, fuente = ?, drive_link = ?
            WHERE id = ?
        ''', (nombre, descripcion, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente, drive_link, id_material))
        conn.commit()
        conn.close()

        flash("Éxito: Material actualizado correctamente.", "success")
        return redirect(url_for('inventario'))

@app.route('/agregar_grupo_ajax', methods=['POST'])
def agregar_grupo_ajax():
    nombre = request.json.get('nombre')
    if not nombre:
        return jsonify({'success': False, 'error': 'El nombre está vacío'})
    
    conn = get_db_connection()
    try:
        cursor = conn.execute('INSERT INTO grupos (nombre) VALUES (?)', (nombre,))
        conn.commit()
        nuevo_id = cursor.lastrowid
        conn.close()
        return jsonify({'success': True, 'id': nuevo_id, 'nombre': nombre})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'success': False, 'error': 'El grupo ya existe'})

@app.route('/editar_grupo_ajax', methods=['POST'])
def editar_grupo_ajax():
    data = request.json
    id = data.get('id')
    nombre = data.get('nombre')
    nombre_viejo = data.get('nombre_viejo')
    
    if not nombre:
        return jsonify({'success': False, 'error': 'El nombre está vacío'})
        
    conn = get_db_connection()
    try:
        # Actualizar grupo
        conn.execute('UPDATE grupos SET nombre = ? WHERE id = ?', (nombre, id))
        # Actualizar todos los materiales que usaban este grupo al nuevo nombre
        if nombre != nombre_viejo:
            conn.execute('UPDATE materiales SET tipo_material = ? WHERE tipo_material = ?', (nombre, nombre_viejo))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'success': False, 'error': 'El grupo ya existe'})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/eliminar_grupo_ajax', methods=['POST'])
def eliminar_grupo_ajax():
    data = request.json
    id = data.get('id')
    pin = data.get('pin')
    
    if pin != '1234':
        return jsonify({'success': False, 'error': 'PIN incorrecto'})
        
    conn = get_db_connection()
    try:
        conn.execute('DELETE FROM grupos WHERE id = ?', (id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/actualizar_vinculo_ajax', methods=['POST'])
def actualizar_vinculo_ajax():
    data = request.json
    material_id = data.get('material_id')
    link = data.get('link', '')

    if not material_id:
        return jsonify({'success': False, 'error': 'ID de material no proporcionado'})

    conn = get_db_connection()
    try:
        conn.execute('UPDATE materiales SET drive_link = ? WHERE id = ?', (link, material_id))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/agregar_proveedor_ajax', methods=['POST'])
def agregar_proveedor_ajax():
    nit = request.json.get('nit', '')
    nombre = request.json.get('nombre')
    if not nombre:
        return jsonify({'success': False, 'error': 'El nombre está vacío'})
    
    conn = get_db_connection()
    try:
        cursor = conn.execute('INSERT INTO proveedores (nit, nombre) VALUES (?, ?)', (nit, nombre))
        conn.commit()
        nuevo_id = cursor.lastrowid
        conn.close()
        return jsonify({'success': True, 'id': nuevo_id, 'nombre': nombre})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/editar_proveedor_ajax', methods=['POST'])
def editar_proveedor_ajax():
    data = request.json
    id = data.get('id')
    nit = data.get('nit', '')
    nombre = data.get('nombre')
    nombre_viejo = data.get('nombre_viejo')
    
    if not nombre:
        return jsonify({'success': False, 'error': 'El nombre está vacío'})
        
    conn = get_db_connection()
    try:
        # Actualizar proveedor
        conn.execute('UPDATE proveedores SET nit = ?, nombre = ? WHERE id = ?', (nit, nombre, id))
        # Actualizar todos los materiales que usaban este proveedor al nuevo nombre
        if nombre != nombre_viejo:
            conn.execute('UPDATE materiales SET empresa = ? WHERE empresa = ?', (nombre, nombre_viejo))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/eliminar_proveedor_ajax', methods=['POST'])
def eliminar_proveedor_ajax():
    data = request.json
    id = data.get('id')
    pin = data.get('pin')
    
    if pin != '1234':
        return jsonify({'success': False, 'error': 'PIN incorrecto'})
        
    conn = get_db_connection()
    try:
        conn.execute('DELETE FROM proveedores WHERE id = ?', (id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/agregar_entrada', methods=['POST'])
def agregar_entrada():
    if request.method == 'POST':
        material_id = int(request.form['material_id'])
        cantidad = float(request.form['cantidad'])
        precio = float(request.form['precio'])
        fecha = request.form.get('fecha')
        fecha_factura = request.form.get('fecha_factura', '')
        documento = request.form.get('documento', '')
        numero_documento = request.form.get('numero_documento', '')

        if not fecha:
            fecha = datetime.now().strftime('%Y-%m-%d')

        conn = get_db_connection()
        
        # --- LÓGICA DE DEVOLUCIÓN ---
        doc_lower = documento.strip().lower()
        if 'devolucion' in doc_lower or 'devolución' in doc_lower:
            # Buscar si existe una salida con ese número de documento para este material
            salida = conn.execute('''
                SELECT precio_unitario FROM movimientos
                WHERE material_id = ? AND tipo = 'salida' AND numero_documento = ?
                ORDER BY id DESC LIMIT 1
            ''', (material_id, numero_documento)).fetchone()
            
            if salida:
                precio = salida['precio_unitario']
            else:
                # Si no encuentra el doc exacto, intentar con la última salida del material
                ultima_salida = conn.execute('''
                    SELECT precio_unitario FROM movimientos
                    WHERE material_id = ? AND tipo = 'salida'
                    ORDER BY id DESC LIMIT 1
                ''', (material_id,)).fetchone()
                if ultima_salida:
                    precio = ultima_salida['precio_unitario']

        conn.execute('''
            INSERT INTO movimientos (material_id, tipo, cantidad, precio_unitario, fecha, documento, numero_documento, fecha_factura)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (material_id, 'entrada', cantidad, precio, fecha, documento, numero_documento, fecha_factura))
        conn.commit()
        conn.close()
        
        flash("Éxito: Entrada registrada correctamente.", "success")
        if request.form.get('origen') == 'vista_entradas':
            return redirect(url_for('entradas'))
        return redirect(url_for('index'))

@app.route('/agregar_salida', methods=['POST'])
def agregar_salida():
    if request.method == 'POST':
        material_id = int(request.form['material_id'])
        cantidad_a_sacar = float(request.form['cantidad'])
        fecha = request.form.get('fecha')
        documento = request.form.get('documento', '')
        numero_documento = request.form.get('numero_documento', '')

        if not fecha:
            fecha = datetime.now().strftime('%Y-%m-%d')

        conn = get_db_connection()
        
        # Validar existencias actuales
        material = conn.execute('SELECT * FROM materiales WHERE id = ?', (material_id,)).fetchone()
        cant_actual = material['cantidad_inicial']
        total_actual = material['cantidad_inicial'] * material['precio_unitario']
        precio_promedio = material['precio_unitario']
        
        movimientos = conn.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (material_id,)).fetchall()
        for mov in movimientos:
            if mov['tipo'] == 'entrada':
                cant_actual += mov['cantidad']
                total_actual += (mov['cantidad'] * mov['precio_unitario'])
                precio_promedio = total_actual / cant_actual
            elif mov['tipo'] == 'salida':
                cant_actual -= mov['cantidad']
                total_actual -= (mov['cantidad'] * precio_promedio)

        # BLOQUEO SI NO HAY STOCK SUFICIENTE
        if cantidad_a_sacar > cant_actual:
            conn.close()
            flash(f"Error: No puedes sacar {cantidad_a_sacar} unidades. Solo hay {cant_actual} disponibles del material '{material['nombre']}'.", "error")
            
            if request.form.get('origen') == 'vista_salidas':
                return redirect(url_for('salidas'))
            return redirect(url_for('index'))

        # Si hay stock, registrar la salida
        conn.execute('''
            INSERT INTO movimientos (material_id, tipo, cantidad, precio_unitario, fecha, documento, numero_documento)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (material_id, 'salida', cantidad_a_sacar, precio_promedio, fecha, documento, numero_documento))
        conn.commit()
        conn.close()
        
        flash("Éxito: Salida registrada correctamente.", "success")
        if request.form.get('origen') == 'vista_salidas':
            return redirect(url_for('salidas'))
        return redirect(url_for('index'))

@app.route('/eliminar_material/<int:id>', methods=['POST'])
def eliminar_material(id):
    if request.method == 'POST':
        conn = get_db_connection()
        conn.execute('DELETE FROM movimientos WHERE material_id = ?', (id,))
        conn.execute('DELETE FROM materiales WHERE id = ?', (id,))
        conn.commit()
        conn.close()
        flash("Éxito: Material eliminado correctamente.", "success")
        return redirect(url_for('inventario'))

@app.route('/entradas')
def entradas():
    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    conn.close()
    return render_template('entradas.html', materiales=materiales)

@app.route('/salidas')
def salidas():
    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    conn.close()
    return render_template('salidas.html', materiales=materiales)

@app.route('/reporte')
def reporte():
    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    
    selected_material_id = request.args.get('material_id', type=int)
    mes_filtro = request.args.get('mes')
    if not mes_filtro:
        mes_filtro = datetime.now().strftime('%Y-%m')
        
    reporte_datos = None
    
    if selected_material_id:
        mat = conn.execute('SELECT * FROM materiales WHERE id = ?', (selected_material_id,)).fetchone()
        if mat:
            mat_id = mat['id']
            cant_saldo = mat['cantidad_inicial']
            precio_promedio = mat['precio_unitario']
            total_saldo = cant_saldo * precio_promedio
            
            movimientos = conn.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat_id,)).fetchall()
            
            if mes_filtro != 'todos':
                movs_anteriores = [m for m in movimientos if m['fecha'] < f"{mes_filtro}-01"]
                movs_actuales = [m for m in movimientos if m['fecha'].startswith(mes_filtro)]
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
    conn.close()
    return render_template('reporte.html', materiales=materiales, reporte_datos=reporte_datos, selected_material_id=selected_material_id, mes_filtro=mes_filtro)

# --- RUTAS DE EXPORTACIÓN A EXCEL (CSV) ---
@app.route('/exportar_inventario')
def exportar_inventario():
    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    conn.close()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inventario"
    
    # Estilos basados en la interfaz (CSS)
    fill_hdr_gris = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    font_hdr_gris = Font(color="475569", bold=True)
    alignment_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    alignment_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    border_thin = Border(left=Side(style='thin', color='E2E8F0'), 
                         right=Side(style='thin', color='E2E8F0'), 
                         top=Side(style='thin', color='E2E8F0'), 
                         bottom=Side(style='thin', color='E2E8F0'))

    headers = ['Nombre', 'Descripción', 'Grupo', 'No. Metrico', 'Origen', 'Fuente', 'Proveedor', 'Presentacion', 'Unidad', 'Existencia Inicial', 'Costo Unitario (Q)']
    
    ws.append(headers)
    for col_num, cell in enumerate(ws[1], 1):
        cell.fill = fill_hdr_gris
        cell.font = font_hdr_gris
        # Alinear a la izquierda todas las columnas de texto
        cell.alignment = alignment_left if col_num <= 9 else alignment_right
        cell.border = border_thin

    for idx, mat in enumerate(materiales, 1):
        row = [mat['nombre'], mat['descripcion'], mat['tipo_material'], mat['numero_metrico'], mat['origen'], mat['fuente'], mat['empresa'], mat['presentacion'], mat['unidad'], mat['cantidad_inicial'], round(mat['precio_unitario'], 2)]
        ws.append(row)
        for col_num, cell in enumerate(ws[ws.max_row], 1):
            cell.alignment = alignment_left if col_num <= 9 else alignment_right
            cell.border = border_thin
    # Ajustar ancho de las columnas automáticamente
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column].width = max_length + 2

    output = BytesIO()
    wb.save(output)
    
    return Response(output.getvalue(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': 'attachment; filename=Plantilla_Inventario.xlsx'})

@app.route('/exportar_kardex')
def exportar_kardex():
    mes_filtro = request.args.get('mes')
    if not mes_filtro:
        mes_filtro = datetime.now().strftime('%Y-%m')
        
    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    
    wb = openpyxl.Workbook()
    
    # Hoja 1: Resumen
    ws_resumen = wb.active
    ws_resumen.title = "Kardex Resumido"
    
    # Hoja 2: Detallado
    ws_detallado = wb.create_sheet(title="Kardex Detallado")

    # Estilos basados en la interfaz (CSS)
    fill_hdr_gris = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    font_hdr_gris = Font(color="475569", bold=True)
    fill_hdr_verde = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    font_hdr_verde = Font(color="065F46", bold=True)
    fill_hdr_naranja = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
    font_hdr_naranja = Font(color="9A3412", bold=True)
    fill_hdr_azul = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    font_hdr_azul = Font(color="1E3A8A", bold=True)
    
    fill_celda_verde = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    fill_celda_naranja = PatternFill(start_color="FFF7ED", end_color="FFF7ED", fill_type="solid")
    fill_celda_azul = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    font_naranja_bold = Font(color="9A3412", bold=True)

    alignment_left = Alignment(horizontal="left", vertical="center")
    alignment_right = Alignment(horizontal="right", vertical="center")
    border_thin = Border(left=Side(style='thin', color='E2E8F0'), right=Side(style='thin', color='E2E8F0'), top=Side(style='thin', color='E2E8F0'), bottom=Side(style='thin', color='E2E8F0'))

    # --- CONFIGURAR HOJA: KARDEX RESUMIDO ---
    headers_resumen = ['Producto', 'Grupo', 'Entradas', 'Salidas', 'Existencia', 'Costo Prom. (Q)', 'Total (Q)']
    ws_resumen.append(headers_resumen)
    for col_num, cell in enumerate(ws_resumen[1], 1):
        if col_num == 3: cell.fill, cell.font = fill_hdr_verde, font_hdr_verde
        elif col_num == 4: cell.fill, cell.font = fill_hdr_naranja, font_hdr_naranja
        elif col_num >= 5: cell.fill, cell.font = fill_hdr_azul, font_hdr_azul
        else: cell.fill, cell.font = fill_hdr_gris, font_hdr_gris
        cell.alignment = alignment_left if col_num <= 2 else alignment_right
        cell.border = border_thin

    # --- CONFIGURAR HOJA: KARDEX DETALLADO ---
    headers_detallado = ['Producto', 'Grupo', 'Fecha', 'Detalle', 'Documento', 'No. Documento', 
                 'Entrada Cantidad', 'Entrada Costo (Q)', 'Entrada Total (Q)',
                 'Salida Cantidad', 'Salida Costo (Q)', 'Salida Total (Q)',
                 'Saldo Cantidad', 'Saldo Costo Prom (Q)', 'Saldo Total (Q)']
                 
    ws_detallado.append(headers_detallado)
    for col_num, cell in enumerate(ws_detallado[1], 1):
        if 7 <= col_num <= 9: cell.fill, cell.font = fill_hdr_verde, font_hdr_verde
        elif 10 <= col_num <= 12: cell.fill, cell.font = fill_hdr_naranja, font_hdr_naranja
        elif 13 <= col_num <= 15: cell.fill, cell.font = fill_hdr_azul, font_hdr_azul
        else: cell.fill, cell.font = fill_hdr_gris, font_hdr_gris
        cell.alignment = alignment_left if col_num <= 6 else alignment_right
        cell.border = border_thin

    # Variables para totales del Kardex Resumido
    total_entradas_res = 0
    total_salidas_res = 0
    total_existencia_res = 0
    gran_total_res = 0

    for idx, mat in enumerate(materiales, 1):
        mat_id = mat['id']
        cant_saldo = mat['cantidad_inicial']
        precio_promedio = mat['precio_unitario']
        total_saldo = cant_saldo * precio_promedio
        
        movimientos = conn.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat_id,)).fetchall()
        
        if mes_filtro != 'todos':
            movs_anteriores = [m for m in movimientos if m['fecha'] < f"{mes_filtro}-01"]
            movs_actuales = [m for m in movimientos if m['fecha'].startswith(mes_filtro)]
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

        ini_cant = cant_saldo
        acum_ingresos = 0
        acum_salidas = 0
        
        # Variables para totales del Kardex Detallado por producto
        sum_entrada_cant = 0
        sum_entrada_total = 0
        sum_salida_cant = 0
        sum_salida_total = 0
        
        titulo_saldo = 'Saldo Inicial' if mes_filtro == 'todos' else f'Saldo Anterior ({mes_filtro})'
        row_det = [
            mat['nombre'], mat['tipo_material'], '-', titulo_saldo, '', '',
            '', '', '', '', '', '',
            cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
        ]
        ws_detallado.append(row_det)
        for col_num, cell in enumerate(ws_detallado[ws_detallado.max_row], 1):
            if 13 <= col_num <= 15: cell.fill = fill_celda_azul
            cell.alignment = alignment_left if col_num <= 6 else alignment_right
            cell.border = border_thin
        
        for mov in movs_actuales:
            doc = mov['documento'] or ''
            num_doc = mov['numero_documento'] or ''
            
            if mov['tipo'] == 'entrada':
                costo_mov = mov['cantidad'] * mov['precio_unitario']
                cant_saldo += mov['cantidad']
                total_saldo += costo_mov
                acum_ingresos += mov['cantidad']
                sum_entrada_cant += mov['cantidad']
                sum_entrada_total += costo_mov
                if cant_saldo > 0: precio_promedio = total_saldo / cant_saldo
                row_det = [
                    mat['nombre'], mat['tipo_material'], mov['fecha'], 'Entrada / Compra', doc, num_doc,
                    mov['cantidad'], round(mov['precio_unitario'], 2), round(costo_mov, 2),
                    '', '', '',
                    cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
                ]
                ws_detallado.append(row_det)
                for col_num, cell in enumerate(ws_detallado[ws_detallado.max_row], 1):
                    if 7 <= col_num <= 9: cell.fill = fill_celda_verde
                    elif 13 <= col_num <= 15: cell.fill = fill_celda_azul
                    cell.alignment = alignment_left if col_num <= 6 else alignment_right
                    cell.border = border_thin

            elif mov['tipo'] == 'salida':
                costo_mov = mov['cantidad'] * precio_promedio
                cant_saldo -= mov['cantidad']
                total_saldo -= costo_mov
                acum_salidas += mov['cantidad']
                sum_salida_cant += mov['cantidad']
                sum_salida_total += costo_mov
                row_det = [
                    mat['nombre'], mat['tipo_material'], mov['fecha'], 'Salida / Egreso', doc, num_doc,
                    '', '', '',
                    mov['cantidad'], round(precio_promedio, 2), round(costo_mov, 2),
                    cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
                ]
                ws_detallado.append(row_det)
                for col_num, cell in enumerate(ws_detallado[ws_detallado.max_row], 1):
                    if 10 <= col_num <= 12: 
                        cell.fill = fill_celda_naranja
                        if col_num == 10: cell.font = font_naranja_bold
                    elif 13 <= col_num <= 15: cell.fill = fill_celda_azul
                    cell.alignment = alignment_left if col_num <= 6 else alignment_right
                    cell.border = border_thin

        # FILA DE TOTALES POR PRODUCTO (DETALLADO)
        row_det_total = [
            '', '', '', 'TOTALES PERIODO', '', '',
            sum_entrada_cant, '', round(sum_entrada_total, 2),
            sum_salida_cant, '', round(sum_salida_total, 2),
            '', '', ''
        ]
        ws_detallado.append(row_det_total)
        for col_num, cell in enumerate(ws_detallado[ws_detallado.max_row], 1):
            cell.font = Font(bold=True)
            if 7 <= col_num <= 9: cell.fill = fill_hdr_verde
            elif 10 <= col_num <= 12: cell.fill = fill_hdr_naranja
            cell.alignment = alignment_left if col_num <= 6 else alignment_right
            cell.border = border_thin

        # FILA EN BLANCO PARA SEPARAR PRODUCTOS EN EL DETALLADO
        ws_detallado.append([])

        # AGREGAR FILA AL KARDEX RESUMIDO
        row_res = [
            mat['nombre'], mat['tipo_material'],
            acum_ingresos, acum_salidas,
            cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
        ]
        ws_resumen.append(row_res)
        for col_num, cell in enumerate(ws_resumen[ws_resumen.max_row], 1):
            if col_num == 3: cell.fill = fill_celda_verde
            elif col_num == 4: cell.fill = fill_celda_naranja
            elif col_num >= 5: cell.fill = fill_celda_azul
            cell.alignment = alignment_left if col_num <= 2 else alignment_right
            cell.border = border_thin
            
        # Sumar al Gran Total del Kardex Resumido
        total_entradas_res += acum_ingresos
        total_salidas_res += acum_salidas
        total_existencia_res += cant_saldo
        gran_total_res += total_saldo

    # AGREGAR FILA DE TOTAL GENERAL AL KARDEX RESUMIDO AL FINAL DE TODO
    row_res_tot = [
        'TOTAL GENERAL', '',
        total_entradas_res, total_salidas_res,
        total_existencia_res, '', round(gran_total_res, 2)
    ]
    ws_resumen.append(row_res_tot)
    for col_num, cell in enumerate(ws_resumen[ws_resumen.max_row], 1):
        cell.font = Font(bold=True)
        cell.fill = fill_hdr_gris
        cell.alignment = alignment_left if col_num <= 2 else alignment_right
        cell.border = border_thin

    # Ajustar ancho de las columnas automáticamente en AMBAS hojas
    for sheet in [ws_resumen, ws_detallado]:
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            sheet.column_dimensions[column].width = max_length + 2

    conn.close()
    
    output = BytesIO()
    wb.save(output)
    nombre_archivo = f'Kardex_General_{mes_filtro}.xlsx' if mes_filtro != 'todos' else 'Kardex_General_Completo.xlsx'
    return Response(output.getvalue(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': f'attachment; filename={nombre_archivo}'})

@app.route('/cargar_excel', methods=['GET', 'POST'])
def cargar_excel():
    if request.method == 'POST':
        if 'archivo_excel' not in request.files:
            flash('Error: No se encontró el archivo en la solicitud.', 'error')
            return redirect(request.url)
        
        file = request.files['archivo_excel']
        
        if file.filename == '':
            flash('Error: No se seleccionó ningún archivo.', 'error')
            return redirect(request.url)

        if file and file.filename.endswith('.xlsx'):
            try:
                conn = get_db_connection()
                cursor = conn.cursor()
                workbook = openpyxl.load_workbook(file)
                sheet = workbook.active
                
                rows_processed = 0
                rows_imported = 0
                rows_skipped = 0
                
                sql_insert = '''
                    INSERT INTO materiales (nombre, descripcion, tipo_material, numero_metrico, origen, fuente, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, drive_link)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                '''
                
                for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                    rows_processed += 1
                    try:
                        if len(row) < 11:
                            rows_skipped += 1
                            continue

                        (nombre_raw, descripcion, tipo_material_raw, numero_metrico, origen, 
                         fuente_raw, empresa_raw, presentacion, unidad, cantidad_inicial_raw, precio_unitario_raw) = row[:11]

                        # --- VALIDACIÓN DE DATOS OBLIGATORIOS ---
                        if not all([nombre_raw, tipo_material_raw, fuente_raw, empresa_raw, cantidad_inicial_raw is not None, precio_unitario_raw is not None]):
                            rows_skipped += 1
                            continue

                        # --- LIMPIEZA Y CREACIÓN AUTOMÁTICA DE ENTIDADES ---
                        nombre = str(nombre_raw).strip()
                        tipo_material = str(tipo_material_raw).strip()
                        fuente = str(fuente_raw).strip()
                        empresa = str(empresa_raw).strip()

                        # Si el grupo, fuente o proveedor no existen, se crean.
                        cursor.execute('INSERT OR IGNORE INTO grupos (nombre) VALUES (?)', (tipo_material,))
                        cursor.execute('INSERT OR IGNORE INTO fuentes (nombre) VALUES (?)', (fuente,))
                        prov_exists = cursor.execute('SELECT id FROM proveedores WHERE nombre = ?', (empresa,)).fetchone()
                        if not prov_exists:
                            cursor.execute('INSERT INTO proveedores (nombre, nit) VALUES (?, ?)', (empresa, ''))

                        cantidad_inicial = float(cantidad_inicial_raw)
                        precio_unitario = float(precio_unitario_raw)

                        values_to_insert = (nombre, descripcion, tipo_material, numero_metrico, origen, fuente, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, '')
                        cursor.execute(sql_insert, values_to_insert)
                        rows_imported += 1

                    except (ValueError, TypeError):
                        rows_skipped += 1
                        continue
                
                conn.commit()
                conn.close()
                
                flash_message = f"Éxito: Carga completada. Se importaron {rows_imported} materiales."
                if rows_skipped > 0:
                    flash_message += f" Se omitieron {rows_skipped} filas por datos faltantes o formato incorrecto."
                flash(flash_message, "success")

            except Exception as e:
                flash(f"Error: Ocurrió un problema al procesar el archivo Excel: {e}", "error")
            return redirect(url_for('inventario'))

    return render_template('carga_masiva.html')

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    # --- SISTEMA DE LOGIN PARA LA PANTALLA DE ADMIN ---
    if not session.get('admin_logged_in'):
        if request.method == 'POST':
            if request.form.get('admin_password') == 'admin123': # <- Contraseña de administrador
                session['admin_logged_in'] = True
                flash("Acceso concedido.", "success")
                return redirect(url_for('admin'))
            elif request.form.get('admin_password'):
                flash("Error: Contraseña incorrecta.", "error")
        return render_template('admin.html', login_required=True)

    conn = get_db_connection()
    if request.method == 'POST':
        accion = request.form.get('accion')
        
        if accion == 'logout':
            session.pop('admin_logged_in', None)
            flash("Sesión de administrador cerrada.", "success")
            return redirect(url_for('index'))
            
        if accion == 'grupo':
            try:
                conn.execute('INSERT INTO grupos (nombre) VALUES (?)', (request.form['nombre_grupo'],))
                flash("Éxito: Grupo agregado correctamente.", "success")
            except sqlite3.IntegrityError:
                flash("Error: El grupo ya existe.", "error")
                
        elif accion == 'proveedor':
            conn.execute('INSERT INTO proveedores (nit, nombre) VALUES (?, ?)', 
                         (request.form['nit'], request.form['nombre']))
            flash("Éxito: Proveedor agregado correctamente.", "success")
            
        elif accion == 'fuente':
            try:
                conn.execute('INSERT INTO fuentes (nombre) VALUES (?)', (request.form['nombre_fuente'],))
                flash("Éxito: Fuente agregada correctamente.", "success")
            except sqlite3.IntegrityError:
                flash("Error: La fuente ya existe.", "error")
                
        conn.commit()
        return redirect(url_for('admin'))
        
    grupos = conn.execute('SELECT * FROM grupos ORDER BY nombre ASC').fetchall()
    proveedores = conn.execute('SELECT * FROM proveedores ORDER BY nombre ASC').fetchall()
    fuentes = conn.execute('SELECT * FROM fuentes ORDER BY nombre ASC').fetchall()
    conn.close()
    return render_template('admin.html', grupos=grupos, proveedores=proveedores, fuentes=fuentes)

@app.route('/eliminar_grupo/<int:id>', methods=['POST'])
def eliminar_grupo(id):
    conn = get_db_connection()
    conn.execute('DELETE FROM grupos WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    flash("Éxito: Grupo eliminado correctamente.", "success")
    return redirect(url_for('admin'))

@app.route('/eliminar_proveedor/<int:id>', methods=['POST'])
def eliminar_proveedor(id):
    conn = get_db_connection()
    conn.execute('DELETE FROM proveedores WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    flash("Éxito: Proveedor eliminado correctamente.", "success")
    return redirect(url_for('admin'))

@app.route('/eliminar_fuente/<int:id>', methods=['POST'])
def eliminar_fuente(id):
    conn = get_db_connection()
    conn.execute('DELETE FROM fuentes WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    flash("Éxito: Fuente eliminada correctamente.", "success")
    return redirect(url_for('admin'))

@app.route('/consultor')
def consultor():
    conn = get_db_connection()
    materiales_db = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    stock_materiales = []

    for mat in materiales_db:
        mat_id = mat['id']
        cant_saldo = mat['cantidad_inicial']
        movimientos = conn.execute('SELECT tipo, cantidad FROM movimientos WHERE material_id = ?', (mat_id,)).fetchall()
        
        for mov in movimientos:
            if mov['tipo'] == 'entrada':
                cant_saldo += mov['cantidad']
            elif mov['tipo'] == 'salida':
                cant_saldo -= mov['cantidad']
                
        # Convertir la fila de la base de datos (sqlite3.Row) a un diccionario normal.
        material_info = dict(mat)
        # Añadir el stock calculado a este diccionario.
        material_info['stock'] = cant_saldo
        
        stock_materiales.append(material_info)
    conn.close()
    return render_template('consultor.html', materiales=stock_materiales)

if __name__ == '__main__':
    inicializar_db()
    app.run(host='0.0.0.0', port=3000, debug=True)