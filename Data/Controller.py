import mysql.connector
import pandas as pd
from mysql.connector import Error

def get_pharmacy_data():
    try:
        connection = mysql.connector.connect(
            host='localhost',
            database='bayer',
            user='admin2',
            password='admin'
        )

        if connection.is_connected():
            # Робимо JOIN з таблицею LA_Price, щоб отримати ціну для кожного бренду
            query = """
                SELECT 
                    t.PERIOD, 
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
        print(f"Помилка при підключенні до MySQL: {e}")
        return None

    finally:
        if 'connection' in locals() and connection.is_connected():
            connection.close()