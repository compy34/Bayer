import mysql.connector
import pandas as pd
from mysql.connector import Error


def get_pharmacy_data():
  try:
    connection = mysql.connector.connect(
        host='localhost',
        database='bayer',
        user='admin2',
        password='admin',
    )

    if connection.is_connected():
      # Замість t.PERIOD беремо тільки те, що реально є в таблиці
      query = """
                SELECT 
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
    print(f'Помилка при підключенні до MySQL: {e}')
    return None

  finally:
    if 'connection' in locals() and connection.is_connected():
      connection.close()


def get_requisites_data():
  try:
    # Завантажуємо файл реквізитів з кореня проєкту
    df_req = pd.read_csv('реквізити.csv', encoding='utf-8')
    return df_req
  except Exception as e:
    print(f'Помилка завантаження файлу реквізитів: {e}')
    return None