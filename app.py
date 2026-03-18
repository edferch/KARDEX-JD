import sqlite3
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, Response
from datetime import datetime
import csv
from io import StringIO

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
    materiales_db = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    materiales_kardex = []
    
    # Lógica de Costo Promedio Ponderado
    for mat in materiales_db:
        mat_id = mat['id']
        
        cant_saldo = mat['cantidad_inicial']
        precio_promedio = mat['precio_unitario']
        total_saldo = cant_saldo * precio_promedio
        
        acum_ingreso_cant = 0
        acum_ingreso_total = 0
        acum_salida_cant = 0
        acum_salida_total = 0
        
        movimientos = conn.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat_id,)).fetchall()
        
        for mov in movimientos:
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

        materiales_kardex.append({
            'id': mat['id'],
            'nombre': mat['nombre'],
            'tipo_material': mat['tipo_material'],
            'unidad': mat['unidad'],
            'ini_cant': mat['cantidad_inicial'],
            'ini_costo': mat['precio_unitario'],
            'ini_total': mat['cantidad_inicial'] * mat['precio_unitario'],
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
    
    return render_template('index.html', materiales=materiales_kardex, grupos=grupos)

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
        documento = request.form.get('documento', '')
        numero_documento = request.form.get('numero_documento', '')

        if not fecha:
            fecha = datetime.now().strftime('%Y-%m-%d')

        conn = get_db_connection()
        conn.execute('''
            INSERT INTO movimientos (material_id, tipo, cantidad, precio_unitario, fecha, documento, numero_documento)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (material_id, 'entrada', cantidad, precio, fecha, documento, numero_documento))
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
    reporte_datos = None
    
    if selected_material_id:
        mat = conn.execute('SELECT * FROM materiales WHERE id = ?', (selected_material_id,)).fetchone()
        if mat:
            mat_id = mat['id']
            cant_saldo = mat['cantidad_inicial']
            precio_promedio = mat['precio_unitario']
            total_saldo = cant_saldo * precio_promedio
            
            filas_kardex = []
            # Primera fila: El saldo inicial
            filas_kardex.append({
                'fecha': '-', 'detalle': 'Saldo Inicial',
                'ing_cant': '', 'ing_costo': '', 'ing_total': '',
                'sal_cant': '', 'sal_costo': '', 'sal_total': '',
                'saldo_cant': cant_saldo, 'saldo_costo': precio_promedio, 'saldo_total': total_saldo
            })
            
            movs = conn.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat_id,)).fetchall()
            for mov in movs:
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
    return render_template('reporte.html', materiales=materiales, reporte_datos=reporte_datos, selected_material_id=selected_material_id)

# --- RUTAS DE EXPORTACIÓN A EXCEL (CSV) ---
@app.route('/exportar_inventario')
def exportar_inventario():
    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    conn.close()
    
    si = StringIO()
    cw = csv.writer(si)
    # Fila de Encabezados
    cw.writerow(['No.', 'Nombre', 'Grupo', 'No. Metrico', 'Origen', 'Fuente', 'Proveedor', 'Presentacion', 'Unidad', 'Existencia Inicial', 'Costo Unitario (Q)'])
    
    for idx, mat in enumerate(materiales, 1):
        cw.writerow([
            idx, mat['nombre'], mat['tipo_material'], mat['numero_metrico'],
            mat['origen'], mat['fuente'], mat['empresa'], mat['presentacion'],
            mat['unidad'], mat['cantidad_inicial'], round(mat['precio_unitario'], 2)
        ])
        
    # utf-8-sig asegura que Excel lea correctamente los caracteres como tildes
    output = si.getvalue().encode('utf-8-sig')
    return Response(output, mimetype='text/csv', headers={'Content-Disposition': 'attachment; filename=Inventario_General.csv'})

@app.route('/exportar_kardex')
def exportar_kardex():
    conn = get_db_connection()
    materiales = conn.execute('SELECT * FROM materiales ORDER BY nombre ASC').fetchall()
    
    si = StringIO()
    cw = csv.writer(si)
    # Fila de Encabezados
    cw.writerow(['Producto', 'Grupo', 'Fecha', 'Semana del Mes', 'Detalle', 'Documento', 'No. Documento', 
                 'Entrada Cantidad', 'Entrada Costo (Q)', 'Entrada Total (Q)',
                 'Salida Cantidad', 'Salida Costo (Q)', 'Salida Total (Q)',
                 'Saldo Cantidad', 'Saldo Costo Prom (Q)', 'Saldo Total (Q)'])
                 
    for mat in materiales:
        mat_id = mat['id']
        cant_saldo = mat['cantidad_inicial']
        precio_promedio = mat['precio_unitario']
        total_saldo = cant_saldo * precio_promedio
        
        # Saldo Inicial
        cw.writerow([
            mat['nombre'], mat['tipo_material'], '-', '-', 'Saldo Inicial', '', '',
            '', '', '', '', '', '',
            cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
        ])
        
        movs = conn.execute('SELECT * FROM movimientos WHERE material_id = ? ORDER BY fecha ASC, id ASC', (mat_id,)).fetchall()
        for mov in movs:
            doc = mov['documento'] or ''
            num_doc = mov['numero_documento'] or ''
            
            # Calcular la semana del mes (Días 1-7 = Sem 1, Días 8-14 = Sem 2...)
            fecha_obj = datetime.strptime(mov['fecha'], '%Y-%m-%d')
            semana = f"Semana {(fecha_obj.day - 1) // 7 + 1}"
            
            if mov['tipo'] == 'entrada':
                costo_mov = mov['cantidad'] * mov['precio_unitario']
                cant_saldo += mov['cantidad']
                total_saldo += costo_mov
                if cant_saldo > 0: precio_promedio = total_saldo / cant_saldo
                cw.writerow([
                    mat['nombre'], mat['tipo_material'], mov['fecha'], semana, 'Entrada / Compra', doc, num_doc,
                    mov['cantidad'], round(mov['precio_unitario'], 2), round(costo_mov, 2),
                    '', '', '',
                    cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
                ])
            elif mov['tipo'] == 'salida':
                costo_mov = mov['cantidad'] * precio_promedio
                cant_saldo -= mov['cantidad']
                total_saldo -= costo_mov
                cw.writerow([
                    mat['nombre'], mat['tipo_material'], mov['fecha'], semana, 'Salida / Egreso', doc, num_doc,
                    '', '', '',
                    mov['cantidad'], round(precio_promedio, 2), round(costo_mov, 2),
                    cant_saldo, round(precio_promedio, 2), round(total_saldo, 2)
                ])
    conn.close()
    output = si.getvalue().encode('utf-8-sig')
    return Response(output, mimetype='text/csv', headers={'Content-Disposition': 'attachment; filename=Kardex_General_Completo.csv'})

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    conn = get_db_connection()
    if request.method == 'POST':
        accion = request.form.get('accion')
        
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