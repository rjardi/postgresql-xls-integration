"""
Excel xlwings custom function to retrieve stock data from PostgreSQL database.
"""

import os
import xlwings as xw
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


def get_database_engine():
    """
    Create and return a SQLAlchemy engine for PostgreSQL connection.
    
    Returns:
        sqlalchemy.engine.Engine: Database engine instance
    """
    connection_string = (
        f"postgresql+psycopg2://{DB_CONFIG['user']}:{DB_CONFIG['password']}"
        f"@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
        f"?connect_timeout={DB_CONFIG['connect_timeout']}"
    )
    return create_engine(connection_string)


@xw.func
def get_stock_data(p_empavi, p_erpcodave, p_fch):
    """
    Excel custom function to retrieve stock quantity from database.
    
    Args:
        p_empavi (str): Company/warehouse identifier
        p_erpcodave (str): ERP product code
        p_fch (str): Date in appropriate format (YYYY-MM-DD)
    
    Returns:
        int: Stock quantity or error message
    """
    try:
        # Create database engine
        engine = get_database_engine()
        
        # Execute query calling the database function
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
            
            # Fetch the result
            row = result.fetchone()
            
            if row is not None:
                return row[0]
            else:
                return "No data found"
    
    except Exception as e:
        # Return error message to Excel
        return f"Error: {str(e)}"
    
    finally:
        # Dispose of the engine to close connections
        if 'engine' in locals():
            engine.dispose()

