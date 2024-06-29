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

    return table


def save_data_to_excel(table: list[list], path_to_excel: str) -> None:
    with pd.ExcelWriter(path_to_excel, engine='openpyxl') as writer:
        df = pd.DataFrame(table)
        df.to_excel(writer, header=False, index=False)


if __name__ == "__main__":

    pdf_path = "R1_12-SHKAB.434110.021TU.pdf"
    excel_path = os.path.join(os.getcwd(), "data.xlsx")
    data = get_data_from_pdf(pdf_path)
    transform_data = transform_ranges(data)
    print(data)

    # save_data_to_excel(data, excel_path)
