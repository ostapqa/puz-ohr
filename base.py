import os
import pprint
import random

import openpyxl
import pandas as pd
from pathlib import Path


number_of_experts = 5
filename = "table.xlsx"


def read_file(path_to_xlsx):
    p = Path(path_to_xlsx)
    df = pd.read_excel(p)

    return df


def select_random_experts():
    numbers = list(range(1, 1001))
    random_numbers = random.sample(numbers, number_of_experts)

    sorted_keys = [f"E{i}" for i in range(1, number_of_experts)]

    result = dict(zip(sorted_keys, random_numbers))

    return result


def output_random_experts():
    pprint.pprint("Случайно выбранные эксперты")
    pp = pprint.PrettyPrinter(sort_dicts=False)
    pp.pprint(select_random_experts())


# function for development assistance
def fill_the_cells(filepath):
    wb = openpyxl.load_workbook(filepath)

    sheet_name = 'Исходные данные'
    sheet = wb[sheet_name]

    # Generate random data and fill the table
    for row_num in range(2, number_of_experts+2):

        max_value = random.randint(8, 10)
        sheet.cell(row=row_num, column=4, value=max_value)

        avg_value = random.randint(5, 7)
        sheet.cell(row=row_num, column=3, value=avg_value)

        min_value = random.randint(1, 4)
        sheet.cell(row=row_num, column=2, value=min_value)

    # Save the workbook to a file
    wb.save(filepath)

    return pd.read_excel(filepath)


def create_source_file(filepath):
    wb = openpyxl.Workbook()

    default_sheet = wb.active
    wb.remove(default_sheet)

    sheet_name = 'Исходные данные'

    sheet = wb.create_sheet(title=sheet_name)

    column_names = ['минимально', 'среднее', 'максимально']

    # generating rows with experts
    for row_num in range(1, number_of_experts + 1):
        cell_value = f'E{row_num}'
        sheet.cell(row=row_num + 1, column=1, value=cell_value)

    # generating columns with values
    for col_num, column_name in enumerate(column_names, start=1):
        sheet.cell(row=1, column=col_num + 1, value=column_name)

    # generating feedback
    for reason in range(len(column_names)):
        sheet.cell(row=1, column=reason + len(column_names) + 3, value=f"Объяснение {column_names[reason]}")

    wb.save(filepath)

    return pd.read_excel(path_to_file)


def create_list(filepath, step_number):
    wb = openpyxl.load_workbook(filepath)

    sheet_name = f'Вычисления {step_number} шага'
    wb.create_sheet(title=sheet_name)

    wb.save(filepath)

    return pd.read_excel(filepath)


source_dir = input("Выберите директорию для расположения исходного файла.\n")
path_to_file = os.path.join(source_dir, filename)

file_xlsx = create_source_file(path_to_file)

filled_file = fill_the_cells(path_to_file)

print(filled_file)

list1 = create_list(path_to_file, 1)
print(list1)
