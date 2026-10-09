import mysql.connector
import pandas as pd
from mysql.connector import Error
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def get_pharmacy_data():
    try:
        connection = mysql.connector.connect(
            host='localhost',
            database='bayer',
            user='admin2',
            password='admin',
        )
        if connection.is_connected():
            query = query = """
                SELECT 
                    t.CHAIN_NAME,
                    t.PHARMACY_NAME, 
                    t.BRAND, 
                    t.MERCH,
                    COALESCE(p.Price, 0) AS Price
                FROM calc_Merch_KA t
                LEFT JOIN LA_Price p ON t.BRAND = p.Brand
            """
            df = pd.read_sql(query, connection)
            return df
    except Error as e:
        print(f'Помилка MySQL: {e}')
        return None
    finally:
        if 'connection' in locals() and connection.is_connected():
            connection.close()


def get_requisites_data():
    return load_requisites_file(PROJECT_ROOT / 'rekvizity.xlsx')


def load_requisites_file(filename):
    """Load requisites from a user-selected CSV or Excel file."""
    path = Path(filename)
    if path.suffix.lower() == '.csv':
        df_req = pd.read_csv(path, encoding='utf-8', dtype=str)
    elif path.suffix.lower() in {'.xlsx', '.xls'}:
        df_req = pd.read_excel(path, dtype=str)
    else:
        raise ValueError('Підтримуються тільки файли CSV, XLSX або XLS.')

    # Очищуємо заголовки колонок від зайвих пробілів
    df_req.columns = df_req.columns.astype(str).str.strip()
    return df_req.fillna('')