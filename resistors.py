import pdfplumber
import pandas as pd
import re
import os
import argparse


def fill_cels(table: list[list]) -> list[list]:

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


def get_data_from_pdf(path_to_pdf: str) -> list[list]:

    table = []
    with pdfplumber.open(path_to_pdf) as pdf:
        extract_table = pdf.pages[0].extract_table()
        for data_from_table in extract_table[2:len(extract_table) - 1]:
            table.append(data_from_table)
    table = fill_cels(table)

    return table


def transform_ranges(table: list[list]) -> list[list]:

    pattern = r'(?:От|Св\.)\s*(\d+(?:.\d+)?)\s+до\s+(\d+(?:.\d+)?(?:e\d+)?)\s+включ\.?'

    for row in table:
        row[4] = re.sub(r'(\d+)\s*×\s*10(\d+)', lambda match: f'{match.group(1)}e{match.group(2)}', row[4])
        input_str = re.sub(pattern, r'\1;\2', row[4]).replace(',', '.')
        row[4] = re.findall(r'\d+[.;e\d]*', input_str)

    for row in table:
        for i, item in enumerate(row):
            if isinstance(item, str) and '±' in item:
                parts = item.split('\n')
                combined_list = []
                for part in parts:
                    subparts = part.split(';')
                    combined_parts = []
                    for subpart in subparts:
                        number = subpart.strip()
                        if number:
                            combined_parts.append(number)
                    combined_list.append(';'.join(combined_parts))
                row[i] = combined_list
    return table


def parse_data(table: list[list]) -> list[list]:

    new_table = []
    results = []

    for row in table:
        base_part = row[:4] + [row[6]]
        list1 = row[4]
        list2 = row[5]
        for l1, l2 in zip(list1, list2):
            new_table.append(base_part + [l1, l2])

    for row in new_table:
        tolerances = row[-1].split(';')
        if len(tolerances) > 1:
            for tol in tolerances:
                new_row = row.copy()
                new_row[-1] = tol
                results.append(new_row)
        else:
            results.append(row)

    return results


def format_number(value: float) -> str:
    return f"{value:.2f}".rstrip('0').rstrip('.') if value % 1 != 0 else f"{value:.0f}"


def get_values_in_range(min_value: str, max_value: str, e_values: list) -> list[str]:

    results = []
    for n in range(-12, 12):
        factor = 10 ** n
        for value in e_values:
            e48_value = value * factor
            if float(min_value) <= e48_value <= float(max_value):
                results.append(format_number(e48_value))
    return results


def deploy_data(table: list[list]) -> list[list]:

    e24_values = [
        10, 11, 12, 13, 15, 16, 18, 20, 22, 24,
        27, 30, 33, 36, 39, 43, 47, 51, 56, 62,
        68, 75, 82, 91
    ]
    e48_values = [
        1.0, 1.05, 1.10, 1.15, 1.21, 1.27, 1.33, 1.40, 1.47, 1.54,
        1.62, 1.69, 1.78, 1.87, 1.96, 2.05, 2.15, 2.26, 2.37, 2.49,
        2.61, 2.74, 2.87, 3.01, 3.16, 3.32, 3.48, 3.65, 3.83, 4.02,
        4.22, 4.42, 4.64, 4.87, 5.11, 5.36, 5.62, 5.90, 6.19, 6.49,
        6.81, 7.15, 7.50, 7.87, 8.25, 8.66, 9.09, 9.53
    ]
    e96_values = [
        1.00, 1.02, 1.05, 1.07, 1.10, 1.13, 1.15, 1.18, 1.21, 1.24, 1.27, 1.30,
        1.33, 1.37, 1.40, 1.43, 1.47, 1.50, 1.54, 1.58, 1.62, 1.65, 1.69, 1.74,
        1.78, 1.82, 1.87, 1.91, 1.96, 2.00, 2.05, 2.10, 2.16, 2.21, 2.26, 2.32,
        2.37, 2.43, 2.49, 2.55, 2.61, 2.67, 2.74, 2.80, 2.87, 2.94, 3.01, 3.09,
        3.16, 3.24, 3.32, 3.40, 3.48, 3.57, 3.65, 3.74, 3.83, 3.92, 4.02, 4.12,
        4.22, 4.32, 4.42, 4.53, 4.64, 4.75, 4.87, 4.99, 5.11, 5.23, 5.36, 5.49,
        5.62, 5.76, 5.90, 6.04, 6.19, 6.34, 6.49, 6.65, 6.81, 6.98, 7.15, 7.32,
        7.50, 7.68, 7.87, 8.06, 8.25, 8.45, 8.66, 8.87, 9.09, 9.31, 9.53, 9.76
    ]
    e192_values = [
        1.00, 1.01, 1.02, 1.04, 1.05, 1.06, 1.07, 1.09, 1.10, 1.11, 1.13, 1.14, 1.15, 1.17, 1.18, 1.20,
        1.21, 1.23, 1.24, 1.26, 1.27, 1.29, 1.30, 1.32, 1.33, 1.35, 1.37, 1.38, 1.40, 1.42, 1.43, 1.45,
        1.47, 1.49, 1.50, 1.52, 1.54, 1.56, 1.58, 1.60, 1.62, 1.64, 1.65, 1.68, 1.70, 1.72, 1.74, 1.76,
        1.78, 1.80, 1.82, 1.84, 1.87, 1.89, 1.91, 1.93, 1.96, 1.98, 2.00, 2.03, 2.05, 2.08, 2.10, 2.13,
        2.15, 2.18, 2.21, 2.23, 2.26, 2.29, 2.32, 2.34, 2.37, 2.40, 2.43, 2.46, 2.49, 2.52, 2.55, 2.58,
        2.61, 2.64, 2.67, 2.71, 2.74, 2.77, 2.80, 2.84, 2.87, 2.91, 2.94, 2.98, 3.01, 3.05, 3.09, 3.12,
        3.16, 3.20, 3.24, 3.28, 3.32, 3.36, 3.40, 3.44, 3.48, 3.52, 3.57, 3.61, 3.65, 3.70, 3.74, 3.79,
        3.83, 3.88, 3.92, 3.97, 4.02, 4.07, 4.12, 4.17, 4.22, 4.27, 4.32, 4.37, 4.42, 4.48, 4.53, 4.59,
        4.64, 4.70, 4.75, 4.81, 4.87, 4.93, 4.99, 5.05, 5.11, 5.17, 5.23, 5.3, 5.39, 5.46, 5.53, 5.6,
        5.69, 5.76, 5.84, 5.91, 6.04, 6.11, 6.19, 6.26, 6.34, 6.42, 6.49, 6.57, 6.65, 6.73, 6.81, 6.9,
        7.0, 7.09, 7.17, 7.26, 7.35, 7.45, 7.55, 7.65, 7.74, 7.84, 7.94, 8.06, 8.16, 8.25, 8.36, 8.45,
        8.56, 8.66, 8.76, 8.87, 8.98, 9.09, 9.21, 9.31, 9.42, 9.53, 9.76, 9.76, 9.76
    ]

    new_list = []
    item = ''
    for row in table:
        floor_border, ceil_border = row[-2].split(';')
        if row[-1] == '±5':
            item = get_values_in_range(min_value=floor_border, max_value=ceil_border, e_values=e24_values)
        elif row[-1] == '±2':
            item = get_values_in_range(min_value=floor_border, max_value=ceil_border, e_values=e48_values)
        elif row[-1] == '±1':
            item = get_values_in_range(min_value=floor_border, max_value=ceil_border, e_values=e96_values)
        elif row[-1] == '±0,5':
            item = get_values_in_range(min_value=floor_border, max_value=ceil_border, e_values=e192_values)
        new_list.append(item)

    formatted_data = []
    for i, row in enumerate(table):
        replace_values = new_list[i]
        for value in replace_values:
            new_row = row[:]
            new_row[-2] = value
            formatted_data.append(new_row[:])

    return formatted_data


def add_to_list_new_data(table: list[list]) -> list[list]:

    new_list = []
    for row in table:
        item = ''
        if (row[1] == 'Р1-12-0,062 ум.' or row[1] == 'Р1-12-0,062' or
                row[1] == 'Р1-12-0,1 ум.' or row[1] == 'Р1-12-0,125') and 100 < float(row[-2]) <= 1e7:
            item += 'К'
        if 100 < float(row[-2]) <= 1e7:
            item += 'Л'
        if 1 <= float(row[-2]) <= 2.7e7:
            item += 'МН'
        if float(row[-2]) < 1:
            item = ' '
        new_list.append(item)

    formatted_data = []
    for i, row in enumerate(table):
        replace_values = new_list[i]
        for value in replace_values:
            new_row = row[:]
            new_row.append(value)
            formatted_data.append(new_row[:])

    for row in formatted_data:
        if float(row[-3]) < 1000:
            value = format_number(float(row[-3])) + ' Ом'
        elif 1000 <= float(row[-3]) < 1000000:
            value = format_number(float(row[-3]) / 1000) + ' кОм'
        elif float(row[-3]) >= 1000000:
            value = format_number(float(row[-3]) / 1000000) + ' МОм'
        row[-3] = value.replace('.', ',')
    print(*formatted_data, sep='\n')

    return formatted_data


def save_data_to_excel(table: list[list], path_to_excel: str) -> None:
    with pd.ExcelWriter(path_to_excel, engine='openpyxl') as writer:
        df = pd.DataFrame(table)
        df.to_excel(writer, header=False, index=False)


if __name__ == "__main__":

    pdf_path = "R1_12-SHKAB.434110.021TU.pdf"
    excel_path = os.path.join(os.getcwd(), "data.xlsx")

    data = get_data_from_pdf(pdf_path)
    transformed_data = transform_ranges(data)
    parsed_data = parse_data(transformed_data)
    deployed_data = deploy_data(parsed_data)
    updated_data = add_to_list_new_data(deployed_data)
    save_data_to_excel(updated_data, excel_path)
