"""
Script de prueba para verificar la conexión a la base de datos
y la función get_stock_data antes de usar xlwings.
"""

import os
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Database configuration
DB_CONFIG = {
    "host": os.getenv("DB_HOST_URL"),
    "database": os.getenv("DB_NAME"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "port": 5432,
    "connect_timeout": 20,
}


def test_database_connection():
    """
    Prueba la conexión a la base de datos.
    """
    print("=" * 60)
    print("PRUEBA DE CONEXIÓN A LA BASE DE DATOS")
    print("=" * 60)
    
    print("\nConfiguracion de la base de datos:")
    print(f"  Host: {DB_CONFIG['host']}")
    print(f"  Database: {DB_CONFIG['database']}")
    print(f"  User: {DB_CONFIG['user']}")
    print(f"  Port: {DB_CONFIG['port']}")
    print(f"  Timeout: {DB_CONFIG['connect_timeout']}s")
    
    try:
        # Crear cadena de conexión
        connection_string = (
            f"postgresql+psycopg2://{DB_CONFIG['user']}:{DB_CONFIG['password']}"
            f"@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
            f"?connect_timeout={DB_CONFIG['connect_timeout']}"
        )
        
        print("\n✓ Cadena de conexión creada correctamente")
        
        # Crear engine
        engine = create_engine(connection_string)
        print("✓ SQLAlchemy engine creado correctamente")
        
        # Probar conexión
        with engine.connect() as connection:
            result = connection.execute(text("SELECT version()"))
            version = result.fetchone()
            if version:
                version = version[0]
            print("✓ Conexión exitosa a PostgreSQL")
            print(f"\nVersión de PostgreSQL:")
            print(f"  {version}")
        
        engine.dispose()
        print("\n✓ Conexión cerrada correctamente")
        return True
        
    except Exception as e:
        print(f"\n✗ Error al conectar a la base de datos:")
        print(f"  {type(e).__name__}: {str(e)}")
        return False


def test_stock_function(p_empavi, p_erpcodave, p_fch):
    """
    Prueba la función de stock con parámetros específicos.
    
    Args:
        p_empavi (str): Company/warehouse identifier
        p_erpcodave (str): ERP product code
        p_fch (str): Date in format YYYY-MM-DD
    """
    print("\n" + "=" * 60)
    print("PRUEBA DE LA FUNCIÓN api_xls.f_pla_qty_stock")
    print("=" * 60)
    
    print("\nParámetros de prueba:")
    print(f"  p_empavi: {p_empavi}")
    print(f"  p_erpcodave: {p_erpcodave}")
    print(f"  p_fch: {p_fch}")
    
    try:
        # Crear engine
        connection_string = (
            f"postgresql+psycopg2://{DB_CONFIG['user']}:{DB_CONFIG['password']}"
            f"@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
            f"?connect_timeout={DB_CONFIG['connect_timeout']}"
        )
        engine = create_engine(connection_string)
        
        # Ejecutar la función
        with engine.connect() as connection:
            query = text("SELECT api_xls.f_pla_qty_stock(:p_empavi, :p_erpcodave, :p_fch)")
            result = connection.execute(
                query,
                {
                    "p_empavi": p_empavi,
                    "p_erpcodave": p_erpcodave,
                    "p_fch": p_fch
                }
            )
            
            row = result.fetchone()
            
            if row is not None:
                stock_qty = row[0]
                print(f"\n✓ Función ejecutada correctamente")
                print(f"\nResultado:")
                print(f"  Stock Quantity: {stock_qty}")
                print(f"  Tipo: {type(stock_qty)}")
                return stock_qty
            else:
                print("\n⚠ La función no devolvió ningún resultado")
                return None
        
        engine.dispose()
        
    except Exception as e:
        print(f"\n✗ Error al ejecutar la función:")
        print(f"  {type(e).__name__}: {str(e)}")
        return None


def main():
    """
    Función principal para ejecutar todas las pruebas.
    """
    print("\n" + "█" * 60)
    print("  TEST DE INTEGRACIÓN EXCEL-POSTGRESQL")
    print("█" * 60)
    
    # Verificar que las variables de entorno estén cargadas
    if not all([DB_CONFIG['host'], DB_CONFIG['database'], 
                DB_CONFIG['user'], DB_CONFIG['password']]):
        print("\n✗ Error: Variables de entorno no configuradas correctamente")
        print("  Verifica que el archivo .env existe y contiene:")
        print("    - DB_HOST_URL")
        print("    - DB_NAME")
        print("    - DB_USER")
        print("    - DB_PASSWORD")
        return
    
    # Test 1: Conexión a la base de datos
    connection_ok = test_database_connection()
    
    if not connection_ok:
        print("\n" + "=" * 60)
        print("Las pruebas se han detenido debido a error de conexión")
        print("=" * 60)
        return
    
    # Test 2: Función de stock (con parámetros de ejemplo)
    p_empavi = "PDX"
    p_erpcodave = "BRAM"
    p_fch = "22-09-2025"
    
    test_stock_function(p_empavi, p_erpcodave, p_fch)
    
    print("\n" + "=" * 60)
    print("PRUEBAS COMPLETADAS")
    print("=" * 60)


if __name__ == "__main__":
    main()

