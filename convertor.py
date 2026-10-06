import pandas as pd
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


def generate_excel_report(df, output_filename='marketing_report.xlsx'):
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

  # 4. Записуємо базовий датафрейм в Excel (починаючи з 3-го рядка для шапки)
  with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
    pivot_df.to_excel(writer, index=False, startrow=2, sheet_name='Звіт')

  # 5. Відкриваємо файл через openpyxl для оформлення
  import openpyxl

  wb = openpyxl.load_workbook(output_filename)
  ws = wb['Звіт']

  # Видаляємо стандартний рядок заголовків від pandas
  ws.delete_rows(1, 2)

  # --- ШАПКА ТАБЛИЦІ ---
  ws.merge_cells('A1:A2')
  ws['A1'] = '№ п/п'

  ws.merge_cells('B1:B2')
  ws['B1'] = (
      'Назва та адреса Аптеки/Аптечного пункту, де були надані маркетингові'
      ' послуги'
  )

  num_brands = len(brand_columns)
  start_col = 3  # Колонки брендів починаються з C
  end_col = start_col + num_brands - 1

  start_col_letter = get_column_letter(start_col)
  end_col_letter = get_column_letter(end_col)

  ws.merge_cells(f'{start_col_letter}1:{end_col_letter}1')
  ws[f'{start_col_letter}1'] = (
      'Бренди Товару, щодо яких надавалися маркетингові послуги, кількість'
      ' викладок'
  )

  for i, brand in enumerate(brand_columns):
    col_idx = start_col + i
    ws.cell(row=2, column=col_idx, value=brand)

  sum_col_idx = end_col + 1
  sum_col_letter = get_column_letter(sum_col_idx)
  ws.merge_cells(f'{sum_col_letter}1:{sum_col_letter}2')
  ws[f'{sum_col_letter}1'] = 'Сума до виплати, без ПДВ'

  # --- СТИЛІ ---
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

  ws.row_dimensions[1].height = 30
  ws.row_dimensions[2].height = 25

  # Заповнюємо дані рядків
  last_data_row = 2 + len(pivot_df)
  for row_idx, row_data in enumerate(pivot_df.iterrows(), start=3):
    row_values = row_data[1]

    # № п/п
    cell_no = ws.cell(row=row_idx, column=1, value=row_idx - 2)
    cell_no.alignment = center_alignment
    cell_no.border = thin_border

    # Назва аптеки
    cell_name = ws.cell(row=row_idx, column=2, value=row_values['PHARMACY_NAME'])
    cell_name.alignment = left_alignment
    cell_name.border = thin_border

    # Бренди (0 або 1)
    for i, brand in enumerate(brand_columns):
      cell_brand = ws.cell(
          row=row_idx, column=start_col + i, value=row_values[brand]
      )
      cell_brand.alignment = center_alignment
      cell_brand.border = thin_border

    # Сума до виплати по рядку
    cell_sum = ws.cell(
        row=row_idx,
        column=sum_col_idx,
        value=row_values['Сума до виплати, без ПДВ'],
    )
    cell_sum.alignment = center_alignment
    cell_sum.border = thin_border
    cell_sum.number_format = '#,##0.00'

  # Оформлення шапки
  for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=sum_col_idx):
    for cell in row:
      cell.alignment = center_alignment
      cell.border = thin_border
      cell.font = Font(bold=True, size=9)

  # --- ПІДСУМКОВИЙ БЛОК ВНИЗУ ---
  start_row_bottom = last_data_row + 2
  col_label = sum_col_idx - 1  # Колонка перед сумою для тексту
  col_value = sum_col_idx  # Колонка для формули/значення

  # 1. Всього без ПДВ
  ws.merge_cells(
      start_row=start_row_bottom,
      start_column=col_label,
      end_row=start_row_bottom,
      end_column=col_label,
  )
  lbl1 = ws.cell(row=start_row_bottom, column=col_label, value='Всього без ПДВ')
  lbl1.alignment = right_alignment
  lbl1.font = Font(bold=True)
  lbl1.border = thin_border

  val1 = ws.cell(
      row=start_row_bottom,
      column=col_value,
      value=f'=SUM({sum_col_letter}3:{sum_col_letter}{last_data_row})',
  )
  val1.alignment = right_alignment
  val1.font = Font(bold=True)
  val1.border = thin_border
  val1.number_format = '#,##0.00'

  # 2. ПДВ (20%)
  row_vat = start_row_bottom + 1
  lbl2 = ws.cell(row=row_vat, column=col_label, value='ПДВ')
  lbl2.alignment = right_alignment
  lbl2.font = Font(bold=True)
  lbl2.border = thin_border

  val2 = ws.cell(
      row=row_vat,
      column=col_value,
      value=f'={sum_col_letter}{start_row_bottom} * 0.2',
  )
  val2.alignment = right_alignment
  val2.font = Font(bold=True)
  val2.border = thin_border
  val2.number_format = '#,##0.00'

  # 3. Разом з ПДВ
  row_total = start_row_bottom + 2
  lbl3 = ws.cell(row=row_total, column=col_label, value='Разом з ПДВ')
  lbl3.alignment = right_alignment
  lbl3.font = Font(bold=True)
  lbl3.border = thin_border

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

  # Автопідбір ширини колонок
  for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

  ws.column_dimensions['B'].width = 50

  wb.save(output_filename)
  print(
      'Звіт із дворядковою шапкою та підсумками успішно збережено у файл:'
      f' {output_filename}'
  )