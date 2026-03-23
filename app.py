import sqlite3
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, Response, session
from datetime import datetime
import calendar
import csv
from io import StringIO, BytesIO
from openpyxl import Workbook
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
        tipo_material = request.form['tipo_material']
        numero_metrico = request.form['numero_metrico']
        origen = request.form['origen']
        empresa = request.form['empresa']
        presentacion = request.form['presentacion']
        unidad = request.form['unidad']
        cantidad_inicial = int(request.form['cantidad_inicial'])
        precio_unitario = float(request.form['precio_unitario'])
        fuente = request.form.get('fuente', '')

        conn = get_db_connection()
        conn.execute('''
            INSERT INTO materiales (nombre, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (nombre, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente))
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

@app.route('/agregar_entrada', methods=['POST'])
def agregar_entrada():
    if request.method == 'POST':
        material_id = int(request.form['material_id'])
        cantidad = int(request.form['cantidad'])
        precio = float(request.form['precio'])
        fecha = request.form.get('fecha')
        fecha_factura = request.form.get('fecha_factura', '')
        documento = request.form.get('documento', '')
        numero_documento = request.form.get('numero_documento', '')

        if not fecha:
            fecha = datetime.now().strftime('%Y-%m-%d')

        conn = get_db_connection()
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
        cantidad_a_sacar = int(request.form['cantidad'])
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
        return redirect(url_for('index'))

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
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"
    
    # Estilos basados en la interfaz (CSS)
    fill_hdr_gris = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    font_hdr_gris = Font(color="475569", bold=True)
    alignment_left = Alignment(horizontal="left", vertical="center")
    alignment_right = Alignment(horizontal="right", vertical="center")
    border_thin = Border(left=Side(style='thin', color='E2E8F0'), 
                         right=Side(style='thin', color='E2E8F0'), 
                         top=Side(style='thin', color='E2E8F0'), 
                         bottom=Side(style='thin', color='E2E8F0'))

    headers = ['No.', 'Nombre', 'Grupo', 'No. Metrico', 'Origen', 'Fuente', 'Proveedor', 'Presentacion', 'Unidad', 'Existencia Inicial', 'Costo Unitario (Q)']
    
    ws.append(headers)
    for col_num, cell in enumerate(ws[1], 1):
        cell.fill = fill_hdr_gris
        cell.font = font_hdr_gris
        cell.alignment = alignment_left if col_num <= 9 else alignment_right
        cell.border = border_thin

    for idx, mat in enumerate(materiales, 1):
        row = [
            idx, mat['nombre'], mat['tipo_material'], mat['numero_metrico'],
            mat['origen'], mat['fuente'], mat['empresa'], mat['presentacion'],
            mat['unidad'], mat['cantidad_inicial'], round(mat['precio_unitario'], 2)
        ]
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
    
    return Response(output.getvalue(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': 'attachment; filename=Inventario_General.xlsx'})

@app.route('/exportar_kardex')
def exportar_kardex():
    mes_filtro = request.args.get('mes')
    if not mes_filtro:
        mes_filtro = datetime.now().strftime('%Y-%m')
        
    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    
    wb = Workbook()
    
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
    headers_resumen = ['No.', 'Producto', 'Grupo', 'Unidad', 'Exist. Inicial', 'Entradas', 'Salidas', 'Exist. Final', 'Costo Prom. (Q)', 'Total (Q)']
    ws_resumen.append(headers_resumen)
    for col_num, cell in enumerate(ws_resumen[1], 1):
        if col_num == 6: cell.fill, cell.font = fill_hdr_verde, font_hdr_verde
        elif col_num == 7: cell.fill, cell.font = fill_hdr_naranja, font_hdr_naranja
        elif col_num >= 8: cell.fill, cell.font = fill_hdr_azul, font_hdr_azul
        else: cell.fill, cell.font = fill_hdr_gris, font_hdr_gris
        cell.alignment = alignment_left if col_num <= 4 else alignment_right
        cell.border = border_thin

    # --- CONFIGURAR HOJA: KARDEX DETALLADO ---
    headers_detallado = ['Producto', 'Grupo', 'Fecha', 'Semana del Mes', 'Detalle', 'Documento', 'No. Documento', 
                 'Entrada Cantidad', 'Entrada Costo (Q)', 'Entrada Total (Q)',
                 'Salida Cantidad', 'Salida Costo (Q)', 'Salida Total (Q)',
                 'Saldo Cantidad', 'Saldo Costo Prom (Q)', 'Saldo Total (Q)']
                 
    ws_detallado.append(headers_detallado)
    for col_num, cell in enumerate(ws_detallado[1], 1):
        if 8 <= col_num <= 10: cell.fill, cell.font = fill_hdr_verde, font_hdr_verde
        elif 11 <= col_num <= 13: cell.fill, cell.font = fill_hdr_naranja, font_hdr_naranja
        elif 14 <= col_num <= 16: cell.fill, cell.font = fill_hdr_azul, font_hdr_azul
        else: cell.fill, cell.font = fill_hdr_gris, font_hdr_gris
        cell.alignment = alignment_left if col_num <= 7 else alignment_right
        cell.border = border_thin

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
        
        titulo_saldo = 'Saldo Inicial' if mes_filtro == 'todos' else f'Saldo Anterior ({mes_filtro})'
        row_det = [
            mat['nombre'], mat['tipo_material'], '-', '-', titulo_saldo, '', '',
            '', '', '', '', '', '',
            cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
        ]
        ws_detallado.append(row_det)
        for col_num, cell in enumerate(ws_detallado[ws_detallado.max_row], 1):
            if 14 <= col_num <= 16: cell.fill = fill_celda_azul
            cell.alignment = alignment_left if col_num <= 7 else alignment_right
            cell.border = border_thin
        
        for mov in movs_actuales:
            doc = mov['documento'] or ''
            num_doc = mov['numero_documento'] or ''
            
            fecha_obj = datetime.strptime(mov['fecha'], '%Y-%m-%d')
            semana = f"Semana {(fecha_obj.day - 1) // 7 + 1}"
            
            if mov['tipo'] == 'entrada':
                costo_mov = mov['cantidad'] * mov['precio_unitario']
                cant_saldo += mov['cantidad']
                total_saldo += costo_mov
                acum_ingresos += mov['cantidad']
                if cant_saldo > 0: precio_promedio = total_saldo / cant_saldo
                row_det = [
                    mat['nombre'], mat['tipo_material'], mov['fecha'], semana, 'Entrada / Compra', doc, num_doc,
                    mov['cantidad'], round(mov['precio_unitario'], 2), round(costo_mov, 2),
                    '', '', '',
                    cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
                ]
                ws_detallado.append(row_det)
                for col_num, cell in enumerate(ws_detallado[ws_detallado.max_row], 1):
                    if 8 <= col_num <= 10: cell.fill = fill_celda_verde
                    elif 14 <= col_num <= 16: cell.fill = fill_celda_azul
                    cell.alignment = alignment_left if col_num <= 7 else alignment_right
                    cell.border = border_thin

            elif mov['tipo'] == 'salida':
                costo_mov = mov['cantidad'] * precio_promedio
                cant_saldo -= mov['cantidad']
                total_saldo -= costo_mov
                acum_salidas += mov['cantidad']
                row_det = [
                    mat['nombre'], mat['tipo_material'], mov['fecha'], semana, 'Salida / Egreso', doc, num_doc,
                    '', '', '',
                    mov['cantidad'], round(precio_promedio, 2), round(costo_mov, 2),
                    cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
                ]
                ws_detallado.append(row_det)
                for col_num, cell in enumerate(ws_detallado[ws_detallado.max_row], 1):
                    if 11 <= col_num <= 13: 
                        cell.fill = fill_celda_naranja
                        if col_num == 11: cell.font = font_naranja_bold
                    elif 14 <= col_num <= 16: cell.fill = fill_celda_azul
                    cell.alignment = alignment_left if col_num <= 7 else alignment_right
                    cell.border = border_thin

        # FILA EN BLANCO PARA SEPARAR PRODUCTOS EN EL DETALLADO
        ws_detallado.append([])

        # AGREGAR FILA AL KARDEX RESUMIDO
        row_res = [
            idx, mat['nombre'], mat['tipo_material'], mat['unidad'],
            ini_cant, acum_ingresos, acum_salidas,
            cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
        ]
        ws_resumen.append(row_res)
        for col_num, cell in enumerate(ws_resumen[ws_resumen.max_row], 1):
            if col_num == 6: cell.fill = fill_celda_verde
            elif col_num == 7: cell.fill = fill_celda_naranja
            elif col_num >= 8: cell.fill = fill_celda_azul
            cell.alignment = alignment_left if col_num <= 4 else alignment_right
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

if __name__ == '__main__':
    inicializar_db()
    app.run(host='0.0.0.0', port=3000, debug=True)