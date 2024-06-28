import re
import pdfplumber
import pandas as pd
import os
import argparse


def fill_cells(table: list[list]) -> list[list]:

    num_rows = len(table)
    num_cols = len(table[0])

    for col in range(num_cols):
        current_value = None
        for row in range(num_rows):
            if table[row][col] is not None and table[row][col] != '':
                current_value = table[row][col]
            else:
                table[row][col] = current_value
    return table


def find_patterns(text: str) -> [str, str]:

    pattern_conditions = r'\b[А-ЯЁ]{4}\.\d+\.\d+\sТУ\b'
    pattern_type = r'\b[А-ЯЁ]\d+\-\d+\b'

    match_conditions = re.search(pattern_conditions, text)
    match_type = re.search(pattern_type, text)
    return (match_conditions.group() if match_conditions else None,
            match_type.group() if match_type else None)


def get_data_from_pdf(path: str) -> [list[list], str, str]:

    tables_from_pdf = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            match_conditions, match_type = find_patterns(page.extract_text())
            extracted_tables = page.extract_tables()
            for table in extracted_tables:
                table = fill_cells(table)
                tables_from_pdf.append(table)

    return tables_from_pdf, match_conditions, match_type


def split_cell(cell: str) -> [list, bool]:
    gap = False
    if '…' in cell or ';' in cell:
        value = re.split(r'[…;]', cell)
        gap = True
    else:
        value = [cell]
    return value, gap


def format_number(value: float) -> str:
    return f"{value:.2f}".rstrip('0').rstrip('.') if value % 1 != 0 else f"{value:.0f}"


def get_e24_values_in_range(min_value: str, max_value: str) -> list:

    e24_values = [
        10, 11, 12, 13, 15, 16, 18, 20, 22, 24,
        27, 30, 33, 36, 39, 43, 47, 51, 56, 62,
        68, 75, 82, 91
    ]
    results = []

    if float(max_value) < float(min_value):
        max_value = float(max_value) * 1000000

    for n in range(-12, 12):
        factor = 10 ** n
        for value in e24_values:
            e24_value = value * factor
            if float(min_value) <= e24_value <= float(max_value):
                results.append(format_number(e24_value))
    return results


def get_e6_values_in_range(min_value: str, max_value: str) -> list:

    e6_values = [10, 15, 22, 33, 47, 68]
    results = []

    if float(max_value) < float(min_value):
        max_value = float(max_value) * 1000000

    for value in e6_values:
        e6_value = value
        while e6_value <= float(max_value):
            if e6_value >= float(min_value):
                results.append(format_number(e6_value))
            e6_value *= 10
    return results


def parse_data(table: list[list]) -> list[list]:

    new_list = []
    pattern_for_pf = r'^[\d.,\s\W]*[пФ]*[\d.,\s\W]*$'
    pattern_for_mkf = r'^[\d.,\s\W]*[мкФ]*[\d.,\s\W]*$'
    first_loop = True

    for tbl in table[1:]:
        for row in tbl[3:len(tbl) - 1]:
            items, is_gap = split_cell(row[1])
            for i in range(len(items)):
                cleaned_data = ''.join(filter(lambda x: x.isdigit() or x in ',', items[i])).replace(',', '.')
                if cleaned_data.count('.') > 1:
                    parts = cleaned_data.split('.')
                    cleaned_data = '.'.join(parts[1:])
                if re.match(pattern_for_pf, row[1]):
                    items[i] = cleaned_data
                elif re.match(pattern_for_mkf, row[1]):
                    items[i] = str(int(float(cleaned_data) * 1000000))
                else:
                    if 'пФ' in items[i]:
                        items[i] = cleaned_data
                    elif 'мкФ' in items[i]:
                        items[i] = str(int(float(cleaned_data) * 1000000))
            if is_gap is True:
                if first_loop is True:
                    items = get_e24_values_in_range(min_value=items[0], max_value=items[1])
                else:
                    items = get_e6_values_in_range(min_value=items[0], max_value=items[1])
            new_list.append(items)
        first_loop = False
    return new_list


def get_formatted_data(table: list[list], match_conditions: str, match_type: str) -> list[list]:

    new_table = parse_data(table)
    updated_table = []
    formatted_data = []

    for tbl in table[1:]:
        for row in tbl[3:]:
            if row[0].isdigit():
                formatted_row = [row[0] + ' В', row[1], tbl[0][0], row[3], row[4], row[5], row[6]]
                updated_table.append(formatted_row)

    for row in updated_table:
        if row[2] == 'МП0':
            row.insert(2, '±5%')
        elif row[2] == 'Н30' or row[2] == 'H90':
            row.insert(2, '±20%')

    for i, row in enumerate(updated_table):
        replace_values = new_table[i]
        for value in replace_values:
            new_row = row[:]
            new_row[1] = value + ' пФ'
            formatted_data.append(new_row[:])

    formatted_data = [[match_type, match_conditions] + row for row in formatted_data]
    for row in formatted_data:
        row.append(f"{row[0]}-{row[2]}-{row[3]}{row[4]} {row[5]} {row[1]}")

    return formatted_data


def save_tables_to_excel(tables_to_excel: list[list], path: str) -> None:

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df = pd.DataFrame(tables_to_excel, columns=[
            'Тип', 'ТУ', 'Напряжение', 'Ёмкость', 'Допуск', 'ТКЕ', 'L', 'B', 'H', 'Масса', 'Полная запись'])
        df.to_excel(writer, sheet_name=f"Конденсаторы", index=False)


def run():
    parser = argparse.ArgumentParser(description="Обработка PDF-файла и вывод данных в Excel-таблицу")
    parser.add_argument('-in', '--input', type=str, required=True, help="Путь к входному файлу PDF")
    parser.add_argument('-out', '--output', type=str, default='output_data.xlsx', help="Имя выходного файла Excel")

    args = parser.parse_args()

    pdf_path = args.input
    excel_path = os.path.join(os.getcwd(), args.output)

    tables, conditions, capacitor_type = get_data_from_pdf(pdf_path)
    data = get_formatted_data(tables, conditions, capacitor_type)
    save_tables_to_excel(data, excel_path)
    print(f"Таблица сохранена по пути {excel_path}")


if __name__ == "__main__":

    run()
