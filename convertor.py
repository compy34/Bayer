import pandas as pd
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


def generate_excel_report(
    df, df_req, output_filename='marketing_report.xlsx'
):
  if df is None or df.empty:
    print('Немає даних для обробки.')
    return

  # 1. Створюємо зведену таблицю (бренди — стовпчики, аптеки — рядки)
  pivot_df = df.pivot_table(
      index='PHARMACY_NAME',
      columns='BRAND',
      values='MERCH',
      aggfunc='sum',
      fill_value=0,
  ).reset_index()

  # 2. Формуємо словник цін для підрахунку суми
  price_dict = (
      df.drop_duplicates(subset=['BRAND']).set_index('BRAND')['Price'].to_dict()
  )

  # 3. Додаємо стовпчик "Сума до виплати, без ПДВ"
  brand_columns = pivot_df.columns[1:]
  total_sum = 0
  for brand in brand_columns:
    brand_price = price_dict.get(brand, 0)
    total_sum += pivot_df[brand] * brand_price
  pivot_df['Сума до виплати, без ПДВ'] = total_sum

  # 4. Записуємо базовий датафрейм в Excel (починаючи з 10 рядка для шапки)
  with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
    pivot_df.to_excel(writer, index=False, startrow=9, sheet_name='Звіт')

  import openpyxl

  wb = openpyxl.load_workbook(output_filename)
  ws = wb['Звіт']

  # Видаляємо стандартний рядок заголовків від pandas, який з'явився на 10 рядку
  ws.delete_rows(10, 2)

  # Стилі
  thin_border = Border(
      left=Side(style='thin', color='000000'),
      right=Side(style='thin', color='000000'),
      top=Side(style='thin', color='000000'),
      bottom=Side(style='thin', color='000000'),
  )
  center_alignment = Alignment(
      horizontal='center', vertical='center', wrap_text=True
  )
  left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
  right_alignment = Alignment(
      horizontal='right', vertical='center', wrap_text=True
  )

  # Визначаємо кількість брендів та загальну ширину таблиці
  num_brands = len(brand_columns)
  sum_col_idx = 3 + num_brands  # № п/п (1), Назва (2), Бренди (3...), Сума

  # Діапазон колонок для компактного центрування шапки зверху
  header_start_col = 3
  header_end_col = max(3 + num_brands - 1, 8)

  start_let = get_column_letter(header_start_col)
  end_let = get_column_letter(header_end_col)

  # Отримуємо реквізити (якщо є)
  contract_no = '260326'
  contract_date = '«01» березня 2026 року'
  executor_name = 'ПРИВАТНЕ АКЦІОНЕРНЕ ТОВАРИСТВО «АПТЕКИ ЗАПОРІЖЖЯ»'

  if df_req is not None and not df_req.empty:
    row_req = df_req.iloc[0]
    executor_name = str(row_req.get('Виконавець', executor_name))
    contract_no = str(row_req.get('№ договора ', contract_no))
    contract_date = str(row_req.get('Дата договора', contract_date))

  period_str = 'з "01" серпня по "31" серпня 2026 року.'
  end_date_str = '31.08.2026'

  # --- КОМПАКТНА ШАПКА ПОСЕРЕДИНІ ЗВЕРХУ (Рядки 1 - 6) ---

  # Рядок 1: Додаток
  ws.merge_cells(f'{start_let}1:{end_let}1')
  ws[f'{start_let}1'] = 'Додаток № 1 до Акта ЛІК № ________ про надані послуги'
  ws[f'{start_let}1'].font = Font(bold=True, size=10)
  ws[f'{start_let}1'].alignment = Alignment(horizontal='center', vertical='center')

  # Рядок 2: Договір
  ws.merge_cells(f'{start_let}2:{end_let}2')
  ws[f'{start_let}2'] = (
      f'відповідно до Договору про маркетингові послуги № {contract_no} від'
      f' {contract_date} (надалі – «Договір»)'
  )
  ws[f'{start_let}2'].font = Font(bold=True, size=10)
  ws[f'{start_let}2'].alignment = Alignment(horizontal='center', vertical='center')

  # Рядок 3: Звіт щодо послуг
  ws.merge_cells(f'{start_let}3:{end_let}3')
  ws[f'{start_let}3'] = 'Звіт щодо Послуг, передбачених Договором за період'
  ws[f'{start_let}3'].font = Font(bold=True, size=10)
  ws[f'{start_let}3'].alignment = Alignment(horizontal='center', vertical='center')

  # Рядок 4: Період
  ws.merge_cells(f'{start_let}4:{end_let}4')
  ws[f'{start_let}4'] = f'{period_str}'
  ws[f'{start_let}4'].font = Font(bold=True, size=10)
  ws[f'{start_let}4'].alignment = Alignment(horizontal='center', vertical='center')

  # Рядок 6: Місто зліва, дата справа в межах блоку
  ws.cell(row=6, column=header_start_col, value='м. Київ').font = Font(size=10)
  ws.cell(row=6, column=header_start_col).alignment = left_alignment

  ws.cell(row=6, column=header_end_col, value=end_date_str).font = Font(size=10)
  ws.cell(row=6, column=header_end_col).alignment = right_alignment

  # Рядок 8: Юридична підстава — розтягуємо на всю ширину таблиці від A до кінця
  ws.merge_cells(start_row=8, start_column=1, end_row=8, end_column=sum_col_idx)
  ws.cell(
      row=8,
      column=1,
      value=(
          f'{executor_name} підтверджує, що відповідно до Договору Замовнику були'
          ' надані маркетингові послуги.\n1. Розміщення Товару Замовника на'
          ' полицях аптечних закладів Аптечної мережі відповідно до встановленого'
          ' Замовником плану викладки згідно з Планограмою та адресною'
          ' програмою:'
      ),
  )
  ws.cell(row=8, column=1).font = Font(size=10)
  ws.cell(row=8, column=1).alignment = Alignment(
      horizontal='left', vertical='center', wrap_text=True
  )
  ws.row_dimensions[8].height = 45

  # --- ТАБЛИЧНА ШАПКА (Рядки 10 та 11) ---
  header_row_1 = 10
  header_row_2 = 11

  ws.merge_cells(
      start_row=header_row_1, start_column=1, end_row=header_row_2, end_column=1
  )
  ws.cell(row=header_row_1, column=1, value='№ п/п')

  ws.merge_cells(
      start_row=header_row_1, start_column=2, end_row=header_row_2, end_column=2
  )
  ws.cell(
      row=header_row_1,
      column=2,
      value=(
          'Назва та адреса Аптеки/Аптечного пункту, де були надані маркетингові'
          ' послуги'
      ),
  )

  start_col = 3
  end_col = start_col + num_brands - 1
  start_col_letter = get_column_letter(start_col)
  end_col_letter = get_column_letter(end_col)

  ws.merge_cells(f'{start_col_letter}{header_row_1}:{end_col_letter}{header_row_1}')
  ws[f'{start_col_letter}{header_row_1}'] = (
      'Бренди Товару, щодо яких надавалися маркетингові послуги, кількість'
      ' викладок'
  )

  for i, brand in enumerate(brand_columns):
    col_idx = start_col + i
    ws.cell(row=header_row_2, column=col_idx, value=brand)

  sum_col_letter = get_column_letter(sum_col_idx)
  ws.merge_cells(
      start_row=header_row_1,
      start_column=sum_col_idx,
      end_row=header_row_2,
      end_column=sum_col_idx,
  )
  ws.cell(row=header_row_1, column=sum_col_idx, value='Сума до виплати, без ПДВ')

  ws.row_dimensions[header_row_1].height = 30
  ws.row_dimensions[header_row_2].height = 25

  # --- ЗАПОВНЕННЯ ДАНИХ (починаючи з 12 рядка) ---
  start_data_row = 12
  last_data_row = start_data_row + len(pivot_df) - 1

  for idx, row_data in enumerate(pivot_df.iterrows(), start=start_data_row):
    row_values = row_data[1]

    # № п/п
    cell_no = ws.cell(row=idx, column=1, value=idx - start_data_row + 1)
    cell_no.alignment = center_alignment
    cell_no.border = thin_border

    # Назва аптеки
    cell_name = ws.cell(row=idx, column=2, value=row_values['PHARMACY_NAME'])
    cell_name.alignment = left_alignment
    cell_name.border = thin_border

    # Бренди (0 або 1)
    for i, brand in enumerate(brand_columns):
      cell_brand = ws.cell(
          row=idx, column=start_col + i, value=row_values[brand]
      )
      cell_brand.alignment = center_alignment
      cell_brand.border = thin_border

    # Сума по рядку
    cell_sum = ws.cell(
        row=idx, column=sum_col_idx, value=row_values['Сума до виплати, без ПДВ']
    )
    cell_sum.alignment = center_alignment
    cell_sum.border = thin_border
    cell_sum.number_format = '#,##0.00'

  # Стилі для шапки таблиці
  for row in ws.iter_rows(
      min_row=header_row_1,
      max_row=header_row_2,
      min_col=1,
      max_col=sum_col_idx,
  ):
    for cell in row:
      cell.alignment = center_alignment
      cell.border = thin_border
      cell.font = Font(bold=True, size=9)

  # --- ПІДСУМКОВИЙ БЛОК ВНИЗУ ---
  start_row_bottom = last_data_row + 2
  col_label = sum_col_idx - 1
  col_value = sum_col_idx

  # Всього без ПДВ
  ws.cell(
      row=start_row_bottom, column=col_label, value='Всього без ПДВ'
  ).alignment = right_alignment
  ws.cell(row=start_row_bottom, column=col_label).font = Font(bold=True)
  ws.cell(row=start_row_bottom, column=col_label).border = thin_border

  val1 = ws.cell(
      row=start_row_bottom,
      column=col_value,
      value=f'=SUM({sum_col_letter}{start_data_row}:{sum_col_letter}{last_data_row})',
  )
  val1.alignment = right_alignment
  val1.font = Font(bold=True)
  val1.border = thin_border
  val1.number_format = '#,##0.00'

  # ПДВ (20%)
  row_vat = start_row_bottom + 1
  ws.cell(row=row_vat, column=col_label, value='ПДВ').alignment = (
      right_alignment
  )
  ws.cell(row=row_vat, column=col_label).font = Font(bold=True)
  ws.cell(row=row_vat, column=col_label).border = thin_border

  val2 = ws.cell(
      row=row_vat,
      column=col_value,
      value=f'={sum_col_letter}{start_row_bottom} * 0.2',
  )
  val2.alignment = right_alignment
  val2.font = Font(bold=True)
  val2.border = thin_border
  val2.number_format = '#,##0.00'

  # Разом з ПДВ
  row_total = start_row_bottom + 2
  ws.cell(
      row=row_total, column=col_label, value='Разом з ПДВ'
  ).alignment = right_alignment
  ws.cell(row=row_total, column=col_label).font = Font(bold=True)
  ws.cell(row=row_total, column=col_label).border = thin_border

  val3 = ws.cell(
      row=row_total,
      column=col_value,
      value=(
          f'={sum_col_letter}{start_row_bottom} +'
          f' {sum_col_letter}{row_vat}'
      ),
  )
  val3.alignment = right_alignment
  val3.font = Font(bold=True)
  val3.border = thin_border
  val3.number_format = '#,##0.00'

  # Ширина колонок
  for col in ws.columns:
    col_letter = get_column_letter(col[0].column)
    if start_col <= col[0].column < sum_col_idx:
      ws.column_dimensions[col_letter].width = 16
    elif col[0].column == 1:
      ws.column_dimensions[col_letter].width = 8
    elif col[0].column == 2:
      ws.column_dimensions[col_letter].width = 50
    else:
      max_len = max(len(str(cell.value or '')) for cell in col)
      ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

  wb.save(output_filename)
  print(f'Звіт успішно збережено у файл: {output_filename}')