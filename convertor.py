import pandas as pd
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from act_generator import generate_act_sheet

# Списки брендів із шаблону звітів.
LIK_BRANDS = [
    'Елевіт-3 600',
    'Елевіт-3 600 ПОСМ',
    'Бепантен',
    'Антифлу',
    'Еріус',
    'Реліф',
    'Супрадин-10__',
    'Ренні',
    'Терафлекс + Тералів',
    'Терафлекс + Тералів ПОСМ',
    'Бепантен Плюс',
    'Тералів',
    'Кальцемін',
    'Бепантен Сенсідерм-4__',
    'Терафлекс',
]

DD_BRANDS = [
    'Елевіт',
    'Елевіт ПОСМ',
    'Елевіт-3',
    'Супрадин-13',
    'Супрадин-10',
    'Супрадин-10__',
    'Бепантен Сенсідерм-4',
    'Бепантен Сенсідерм-2',
    'Бепантен Сенсідерм-2 ПОСМ',
    'Бепантен Плюс__',
]

MHL_LIK_SKU = [
    'Еріус сироп 0.5 мг/мл 60мл',
    'Еріус таблетки 5 мг № 10',
    'Антифлу Кідс, порошок пакетики №5',
    'Аспірин С, таб. № 10',
    'Кальцемін Адванс №60',
    'Кальцемін Сільвер №60',
    'Бепантен, крем 5%, туба 30г №1',
    'Бепантен, мазь 5%, туба 30г №1',
    'Бепантен, мазь 5%, туба 100г №1',
    'Бепантен Плюс, крем, туба 30г №1',
    'Елевіт Пронаталь, т.в/о №100',
    'Реліф, супозиторїї ректальні №12',
    'Реліф Адванс, супозиторїї ректальні №12',
    'Реліф Про, супозиторїї ректальні №10',
    'Реліф Ультра, супозиторїї ректальні №12',
    'Ренні, табл. жув., мент. смак. №24',
    "Ренні б/ц м'ятні,табл. жув.№24",
    'Ренні з апельсиновим смаком, табл. жув. №24',
    'Тералів, табл. 220мг №12',
    'Терафлекс, капсули №60',
    'Терафлекс Адванс, капсули №60',
    'Терафлекс, капсули №120',
]

MHL_DD_SKU = [
    'Супрадин Актив, табл. №30',
    'Супрадин Актив, табл. шип., №15',
    'Супрадин Ведмежуйки №30',
    'Супрадин Ведмежуйки №60',
    'Супрадин Бустер, табл. шип., №15',
    'Бепантен Сенсідерм',
    'Елевіт-Комплекс 2, капс. № 30',
    'Елевіт-Комплекс 3, капс. № 30',
]




def build_report_dataframe(df, brand_filter_list=None):
    """Build the pharmacy/brand summary filtered by category."""
    if df is None or df.empty:
        return None

    work_df = df.copy()
    if brand_filter_list is not None:
        work_df = work_df[work_df['BRAND'].isin(brand_filter_list)]

    if work_df.empty:
        return None

    pivot_df = work_df.pivot_table(
        index='PHARMACY_NAME',
        columns='BRAND',
        values='MERCH',
        aggfunc='sum',
        fill_value=0,
    ).reset_index()

    price_dict = (
        work_df.drop_duplicates(subset=['BRAND']).set_index('BRAND')['Price'].to_dict()
    )
    brand_columns = pivot_df.columns[1:]
    total_sum = 0
    for brand in brand_columns:
        total_sum += pivot_df[brand] * price_dict.get(brand, 0)
    pivot_df['Загальна сума'] = total_sum
    return pivot_df


def build_mhl_dataframe(df, brand_filter_list):
    """Aggregate MHL data for one product category over the selected period."""
    if df is None or df.empty:
        return pd.DataFrame(
            columns=[
                'DRUG_ID', 'Назва товару', 'BRAND', 'PENETRATION',
                'PRICE_CP', 'SUPPLY', 'Сума закупленого товару',
            ]
        )

    work_df = df.copy()
    allowed_values = {
        str(value).strip().casefold() for value in brand_filter_list
    }
    sku_values = work_df['SKU_UA'].fillna('').astype(str).str.strip().str.casefold()
    work_df = work_df[sku_values.isin(allowed_values)].copy()
    if work_df.empty:
        return pd.DataFrame(
            columns=[
                'DRUG_ID', 'Назва товару', 'BRAND', 'PENETRATION',
                'PRICE_CP', 'SUPPLY', 'Сума закупленого товару',
            ]
        )

    work_df['_date'] = pd.to_datetime(work_df['DATE'], errors='coerce')
    for column in ['PENETRATION', 'PRICE_CP', 'SUPPLY', 'PENETRATION_TARGET']:
        work_df[column] = pd.to_numeric(work_df[column], errors='coerce').fillna(0)
    work_df = work_df.sort_values('_date')
    group_columns = ['DRUG_ID', 'SKU_UA', 'BRAND']
    result = work_df.groupby(group_columns, dropna=False).agg(
        PENETRATION=('PENETRATION', 'last'),
        PRICE_CP=('PRICE_CP', 'last'),
        SUPPLY=('SUPPLY', 'sum'),
        PENETRATION_TARGET=('PENETRATION_TARGET', 'last'),
    ).reset_index()
    result['Сума закупленого товару'] = (
        result['SUPPLY'] * result['PRICE_CP']
    ).where(
        result['PENETRATION'] >= result['PENETRATION_TARGET'],
        0,
    )
    return result.rename(columns={'SKU_UA': 'Назва товару'})


def mhl_brand_counts(df):
    """Return category counts for diagnostics before Excel generation."""
    if df is None or df.empty or 'BRAND' not in df.columns:
        return 0, 0, 0
    sku_names = df['SKU_UA'].fillna('').astype(str).str.strip().str.casefold()
    lik = {str(value).strip().casefold() for value in MHL_LIK_SKU}
    dd = {str(value).strip().casefold() for value in MHL_DD_SKU}
    lik_match = sku_names.isin(lik)
    dd_match = sku_names.isin(dd)
    return int(lik_match.sum()), int(dd_match.sum()), int(
        (~(lik_match | dd_match)).sum()
    )


def mhl_unknown_samples(df, limit=8):
    """Return real database values that are not classified by the lists."""
    if df is None or df.empty:
        return []
    sku_names = df['SKU_UA'].fillna('').astype(str).str.strip()
    known = {
        str(value).strip().casefold()
        for value in MHL_LIK_SKU + MHL_DD_SKU
    }
    unknown = df.loc[
        ~sku_names.str.casefold().isin(known),
        ['BRAND', 'SKU_UA'],
    ].drop_duplicates()
    return [
        f'{brand} / {sku}'
        for brand, sku in unknown.head(limit).itertuples(index=False, name=None)
    ]


def fill_mhl_worksheet(ws, mhl_df, period_start, period_end, contract_no, contract_date):
    """Write one MHL sheet with totals and VAT."""
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right = Alignment(horizontal='right', vertical='center', wrap_text=True)
    columns = [
        '№ п/п', 'Назва товару', 'PENETRATION', 'PRICE_CP',
        'SUPPLY', 'Сума закупленого товару',
    ]
    headers = [
        '№ п/п',
        'Назва товару',
        'Відсоток аптек мережі, які мають товар',
        'Розрахункова ціна за одиницю товару, грн',
        'Кількість упаковок товару, закуплена аптечною мережею',
        'Сума закупленого товару, грн',
    ]
    end_col = len(headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws.cell(1, 1, f'MHL за період з {period_start} по {period_end}')
    ws.cell(1, 1).font = Font(bold=True, size=12)
    ws.cell(1, 1).alignment = center
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    ws.cell(
        2, 1,
        'Винагорода за забезпечення наявності у аптечному закладі '
        f'Товарів із Рекомендованого переліку Товару Замовника, згідно '
        f'пункту 2.2.2 Договору про маркетингові послуги № {contract_no} '
        f'від {contract_date}',
    )
    ws.cell(2, 1).alignment = center
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(4, col_idx, header)
        cell.font = Font(bold=True, size=9)
        cell.alignment = center
        cell.border = thin_border

    for row_idx, (_, row) in enumerate(mhl_df.iterrows(), 5):
        values = [
            row_idx - 4, row['Назва товару'], row['PENETRATION'],
            row['PRICE_CP'], row['SUPPLY'], row['Сума закупленого товару'],
        ]
        for col_idx, value in enumerate(values, 1):
            cell = ws.cell(row_idx, col_idx, value)
            cell.alignment = center
            cell.border = thin_border
            if col_idx >= 4:
                cell.number_format = '#,##0.00'

    first_data_row = 5
    last_data_row = max(first_data_row, 4 + len(mhl_df))
    total_row = last_data_row + 2
    ws.cell(total_row, end_col - 1, 'Всього закуплено товару, грн.')
    ws.cell(total_row, end_col - 1).alignment = right
    ws.cell(total_row, end_col - 1).font = Font(bold=True)
    total_cell = ws.cell(
        total_row, end_col,
        f'=SUM({get_column_letter(end_col)}{first_data_row}:'
        f'{get_column_letter(end_col)}{last_data_row})',
    )
    total_cell.number_format = '#,##0.00'
    total_cell.alignment = right
    total_cell.font = Font(bold=True)
    total_cell.border = thin_border

    vat_row = total_row + 1
    ws.cell(vat_row, end_col - 1, 'ПДВ 20%')
    ws.cell(vat_row, end_col - 1).alignment = right
    vat_cell = ws.cell(vat_row, end_col, f'={get_column_letter(end_col)}{total_row}*0.2')
    vat_cell.number_format = '#,##0.00'
    vat_cell.alignment = right
    total_with_vat_row = vat_row + 1
    ws.cell(total_with_vat_row, end_col - 1, 'Всього з ПДВ')
    ws.cell(total_with_vat_row, end_col - 1).alignment = right
    total_with_vat = ws.cell(
        total_with_vat_row, end_col,
        f'={get_column_letter(end_col)}{total_row}+'
        f'{get_column_letter(end_col)}{vat_row}',
    )
    total_with_vat.number_format = '#,##0.00'
    total_with_vat.alignment = right
    for column, width in enumerate([8, 42, 18, 18, 20, 22], 1):
        ws.column_dimensions[get_column_letter(column)].width = width


def fill_worksheet_data(ws, pivot_df, brand_columns, period_start, period_end, executor_name, contract_no,
                        contract_date):
    """Допоміжна функція для заповнення структури звіту на аркуші Excel"""
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000'),
    )
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right_alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)

    num_brands = len(brand_columns)
    sum_col_idx = 3 + num_brands
    header_start_col = 3
    header_end_col = max(3 + num_brands - 1, 8)

    start_let = get_column_letter(header_start_col)
    end_let = get_column_letter(header_end_col)

    period_str = f'за період з "{period_start}" по "{period_end}".'
    end_date_str = period_end

    ws.merge_cells(f'{start_let}1:{end_let}1')
    ws[f'{start_let}1'] = 'ДОДАТКОВА УГОДА № ________'
    ws[f'{start_let}1'].font = Font(bold=True, size=10)
    ws[f'{start_let}1'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells(f'{start_let}2:{end_let}2')
    ws[f'{start_let}2'] = (
        f'до договору про надання послуг № {contract_no} від '
        f'{contract_date} (Маркетингові послуги)'
    )
    ws[f'{start_let}2'].font = Font(bold=True, size=10)
    ws[f'{start_let}2'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells(f'{start_let}3:{end_let}3')
    ws[f'{start_let}3'] = 'м. Київ'
    ws[f'{start_let}3'].font = Font(bold=True, size=10)
    ws[f'{start_let}3'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells(f'{start_let}4:{end_let}4')
    ws[f'{start_let}4'] = f'{period_str}'
    ws[f'{start_let}4'].font = Font(bold=True, size=10)
    ws[f'{start_let}4'].alignment = Alignment(horizontal='center', vertical='center')

    ws.cell(row=6, column=header_start_col, value='м. Київ').font = Font(size=10)
    ws.cell(row=6, column=header_start_col).alignment = left_alignment
    ws.cell(row=6, column=header_end_col, value=end_date_str).font = Font(size=10)
    ws.cell(row=6, column=header_end_col).alignment = right_alignment

    ws.merge_cells(start_row=8, start_column=1, end_row=8, end_column=sum_col_idx)
    ws.cell(
        row=8,
        column=1,
        value=(
            f'{executor_name} в особі представника діє на підставі статуту, '
            'та Замовник уклали цю додаткову угоду про наступне:\n1. '
            'Сторони домовилися затвердити Маркетинговий звіт за звітний період.'
        ),
    )
    ws.cell(row=8, column=1).font = Font(size=10)
    ws.cell(row=8, column=1).alignment = Alignment(
        horizontal='left', vertical='center', wrap_text=True
    )
    ws.row_dimensions[8].height = 45

    header_row_1 = 10
    header_row_2 = 11

    ws.merge_cells(
        start_row=header_row_1, start_column=1, end_row=header_row_2, end_column=1
    )
    ws.cell(row=header_row_1, column=1, value='№ з/п')

    ws.merge_cells(
        start_row=header_row_1, start_column=2, end_row=header_row_2, end_column=2
    )
    ws.cell(
        row=header_row_1,
        column=2,
        value=('Найменування аптечного закладу (Аптеки / Мережі)'),
    )

    start_col = 3
    end_col = start_col + num_brands - 1
    start_col_letter = get_column_letter(start_col)
    end_col_letter = get_column_letter(end_col)

    ws.merge_cells(f'{start_col_letter}{header_row_1}:{end_col_letter}{header_row_1}')
    ws[f'{start_col_letter}{header_row_1}'] = 'Бренди'

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
    ws.cell(row=header_row_1, column=sum_col_idx, value='Загальна сума')

    ws.row_dimensions[header_row_1].height = 30
    ws.row_dimensions[header_row_2].height = 25

    start_data_row = 12
    last_data_row = start_data_row + len(pivot_df) - 1

    for idx, row_data in enumerate(pivot_df.iterrows(), start=start_data_row):
        row_values = row_data[1]
        cell_no = ws.cell(row=idx, column=1, value=idx - start_data_row + 1)
        cell_no.alignment = center_alignment
        cell_no.border = thin_border

        cell_name = ws.cell(row=idx, column=2, value=row_values['PHARMACY_NAME'])
        cell_name.alignment = left_alignment
        cell_name.border = thin_border

        for i, brand in enumerate(brand_columns):
            cell_brand = ws.cell(
                row=idx, column=start_col + i, value=row_values[brand]
            )
            cell_brand.alignment = center_alignment
            cell_brand.border = thin_border

        cell_sum = ws.cell(
            row=idx, column=sum_col_idx, value=row_values['Загальна сума']
        )
        cell_sum.alignment = center_alignment
        cell_sum.border = thin_border
        cell_sum.number_format = '#,##0.00'

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

    start_row_bottom = last_data_row + 2
    col_label = sum_col_idx - 1
    col_value = sum_col_idx

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

    row_vat = start_row_bottom + 1
    ws.cell(row=row_vat, column=col_label, value='ПДВ 20%').alignment = right_alignment
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

    row_total = start_row_bottom + 2
    ws.cell(
        row=row_total, column=col_label, value='Всього з ПДВ'
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


def generate_excel_report(
        df,
        df_req,
        output_filename='marketing_report.xlsx',
        period_start=None,
        period_end=None,
        selected_pharmacy='Усі аптеки',
        mhl_df=None,
):
    if (df is None or df.empty) and (mhl_df is None or mhl_df.empty):
        return

    # Розділяємо датасет на два звіти за списками з шаблону.
    if df is not None and not df.empty:
        available_brands = df['BRAND'].dropna().astype(str).str.strip().unique()
        lik_list = [brand for brand in LIK_BRANDS if brand in available_brands]
        dd_list = [brand for brand in DD_BRANDS if brand in available_brands]
        pivot_lik = build_report_dataframe(df, lik_list)
        pivot_dd = build_report_dataframe(df, dd_list)
    else:
        pivot_lik = None
        pivot_dd = None
    mhl_lik = build_mhl_dataframe(mhl_df, MHL_LIK_SKU)
    mhl_dd = build_mhl_dataframe(mhl_df, MHL_DD_SKU)

    period_start = period_start or 'No date'
    period_end = period_end or 'No date'

    contract_no = '260326'
    contract_date = '01 березня 2026 року'
    executor_name = 'ПрАТ «Аптеки Запоріжжя»'

    if df_req is not None and not df_req.empty:
        matched_row = None
        clean_selected = str(selected_pharmacy).strip()
        target_col = None
        for col_candidate in ['chain_name', 'Клиент из базы данных', 'Мережа', 'Назва']:
            if col_candidate in df_req.columns:
                target_col = col_candidate
                break

        if clean_selected and clean_selected != 'Усі аптеки' and target_col:
            filtered_req = df_req[
                df_req[target_col].astype(str).str.strip().str.lower()
                == clean_selected.lower()
                ]
            if not filtered_req.empty:
                matched_row = filtered_req.iloc[0]

        if matched_row is None:
            matched_row = df_req.iloc[0]

        executor_name = str(
            matched_row.get('Виконавець', matched_row.get('Юридическое лицо', executor_name))
        )
        contract_no = str(matched_row.get('№ договора', contract_no))
        contract_date = str(matched_row.get('Дата договора', contract_date))

    import openpyxl
    wb = openpyxl.Workbook()
    # Видаляємо дефолтний аркуш
    default_sheet = wb.active

    # 1. Аркуш Звіт(ЛІК)
    if pivot_lik is not None and not pivot_lik.empty:
        ws_lik = wb.create_sheet(title='Звіт(ЛІК)')
        fill_worksheet_data(ws_lik, pivot_lik, pivot_lik.columns[1:], period_start, period_end, executor_name,
                            contract_no, contract_date)
        sum_lik = pivot_lik['Загальна сума'].sum()
    else:
        sum_lik = 0.0

    # 2. Аркуш Акт(ЛІК)
    ws_akt_lik = wb.create_sheet(title='Акт(ЛІК)')
    generate_act_sheet(
        ws=ws_akt_lik,
        df_req=df_req,
        total_sum_without_vat=sum_lik,
        period_start=period_start,
        period_end=period_end,
        selected_pharmacy=selected_pharmacy,
        act_number='1',
        contract_no=contract_no,
        contract_date=contract_date,
        executor_name=executor_name
    )

    # 3. Аркуш Звіт(ДД)
    if pivot_dd is not None and not pivot_dd.empty:
        ws_dd = wb.create_sheet(title='Звіт(ДД)')
        fill_worksheet_data(ws_dd, pivot_dd, pivot_dd.columns[1:], period_start, period_end, executor_name, contract_no,
                            contract_date)
        sum_dd = pivot_dd['Загальна сума'].sum()
    else:
        sum_dd = 0.0

    # 4. Аркуш Акт(ДД)
    ws_akt_dd = wb.create_sheet(title='Акт(ДД)')
    generate_act_sheet(
        ws=ws_akt_dd,
        df_req=df_req,
        total_sum_without_vat=sum_dd,
        period_start=period_start,
        period_end=period_end,
        selected_pharmacy=selected_pharmacy,
        act_number='2',
        contract_no=contract_no,
        contract_date=contract_date,
        executor_name=executor_name
    )

    ws_mhl_lik = wb.create_sheet(title='MHL(ЛІК)')
    fill_mhl_worksheet(
        ws_mhl_lik, mhl_lik, period_start, period_end, contract_no, contract_date
    )
    ws_mhl_dd = wb.create_sheet(title='MHL(ДД)')
    fill_mhl_worksheet(
        ws_mhl_dd, mhl_dd, period_start, period_end, contract_no, contract_date
    )

    # Видаляємо дефолтний аркуш, якщо він залишився порожнім
    if default_sheet.title in wb.sheetnames and len(wb.sheetnames) > 1:
        # Перевіримо, чи це саме той початковий дефолтний аркуш ("Sheet")
        if default_sheet.title == 'Sheet':
            wb.remove(default_sheet)

    wb.save(output_filename)
    return output_filename