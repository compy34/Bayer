from convertor import generate_excel_report
from Data.Controller import get_pharmacy_data, get_requisites_data


def main():
  print('Отримання даних з бази даних та реквізитів...')
  df = get_pharmacy_data()
  df_req = get_requisites_data()

  print('Генерація фінального звіту в Excel...')
  generate_excel_report(df, df_req, 'marketing_report.xlsx')


if __name__ == '__main__':
  main()