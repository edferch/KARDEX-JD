import sqlite3
import psycopg2

def migrar_base_de_datos():
    print("Iniciando migración de SQLite a PostgreSQL...")

    # 1. Conectar a SQLite (Base de datos origen)
    sqlite_conn = sqlite3.connect('kardex.db')
    sqlite_cursor = sqlite_conn.cursor()

    # 2. Conectar a PostgreSQL (Base de datos destino)
    # Reemplaza 'tu_usuario' y 'tu_contraseña' con tus credenciales de PostgreSQL
    pg_conn = psycopg2.connect(
        dbname="kardex_jd", 
        user="postgres", 
        password="fc17181931", 
        host="localhost"
    )
    pg_cursor = pg_conn.cursor()

    # 3. Crear las tablas en PostgreSQL con la sintaxis correcta (SERIAL para autoincremento)
    tablas_pg = '''
    CREATE TABLE IF NOT EXISTS grupos (
        id SERIAL PRIMARY KEY, nombre TEXT NOT NULL UNIQUE
    );
    CREATE TABLE IF NOT EXISTS fuentes (
        id SERIAL PRIMARY KEY, nombre TEXT NOT NULL UNIQUE
    );
    CREATE TABLE IF NOT EXISTS proveedores (
        id SERIAL PRIMARY KEY, nit TEXT, nombre TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS materiales (
        id SERIAL PRIMARY KEY, nombre TEXT NOT NULL, descripcion TEXT,
        tipo_material TEXT, numero_metrico TEXT, origen TEXT, empresa TEXT,
        presentacion TEXT, unidad TEXT, cantidad_inicial NUMERIC DEFAULT 0,
        precio_unitario NUMERIC DEFAULT 0.0, fuente TEXT, drive_link TEXT, costo_link TEXT
    );
    CREATE TABLE IF NOT EXISTS movimientos (
        id SERIAL PRIMARY KEY, material_id INTEGER REFERENCES materiales(id),
        tipo TEXT, cantidad NUMERIC, precio_unitario NUMERIC, fecha DATE,
        documento TEXT, numero_documento TEXT, fecha_factura DATE,
        departamento TEXT, solicitante TEXT
    );
    '''
    pg_cursor.execute(tablas_pg)
    pg_conn.commit()
    print("Tablas creadas en PostgreSQL exitosamente.")

    # 4. Función genérica para mover datos de una tabla a otra
    def transferir_tabla(nombre_tabla, columnas):
        # Leer de SQLite
        sqlite_cursor.execute(f"SELECT {columnas} FROM {nombre_tabla}")
        filas = sqlite_cursor.fetchall()
        
        if filas:
            # Crear los placeholders (%s, %s, %s...) para PostgreSQL
            placeholders = ', '.join(['%s'] * len(filas[0]))
            query_insert = f"INSERT INTO {nombre_tabla} ({columnas}) VALUES ({placeholders})"
            
            # Insertar en PostgreSQL
            pg_cursor.executemany(query_insert, filas)
            
            # Sincronizar el contador de IDs (Fundamental para que no haya errores al crear nuevos registros)
            pg_cursor.execute(f"SELECT setval(pg_get_serial_sequence('{nombre_tabla}', 'id'), coalesce(max(id), 1), max(id) IS NOT null) FROM {nombre_tabla};")
            
            print(f"✅ Se migraron {len(filas)} registros de la tabla '{nombre_tabla}'.")
        else:
            print(f"⚠️ La tabla '{nombre_tabla}' está vacía. Saltando...")

    # 5. Ejecutar la transferencia en orden (Primero las tablas independientes, luego las que tienen relaciones)
    try:
        transferir_tabla('grupos', 'id, nombre')
        transferir_tabla('fuentes', 'id, nombre')
        transferir_tabla('proveedores', 'id, nit, nombre')
        
        # Ojo: El orden de las columnas debe coincidir exactamente con lo que hay en tu kardex.db actual
        transferir_tabla('materiales', 'id, nombre, tipo_material, numero_metrico, origen, empresa, presentacion, unidad, cantidad_inicial, precio_unitario, fuente, descripcion, drive_link, costo_link')
        
        transferir_tabla('movimientos', 'id, material_id, tipo, cantidad, precio_unitario, fecha, documento, numero_documento, fecha_factura, departamento, solicitante')
        
        pg_conn.commit()
        print("\n🎉 ¡Migración completada con éxito!")

    except Exception as e:
        pg_conn.rollback()
        print(f"\n❌ Error durante la migración: {e}")
        
    finally:
        # Cerrar conexiones
        sqlite_conn.close()
        pg_conn.close()

if __name__ == '__main__':
    migrar_base_de_datos()