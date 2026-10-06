from Data.Controller import get_pharmacy_data
from convertor import generate_excel_report


def main():
    print("Отримання даних з бази даних MySQL...")
    df = get_pharmacy_data()

    print("Генерація звіту в Excel з урахуванням прайсу...")
    generate_excel_report(df, 'marketing_report3.xlsx')


if __name__ == "__main__":
    main()