import pandas as pd
from datetime import date, datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from docx import Document
import random

black_fill = PatternFill(fill_type="solid", fgColor="000000")
holidays_all = 0

def word_to_excel(path):
    doc = Document(path)
    table = doc.tables[0]
    data = []
    for rows in table.rows:
        data.append([cell.text for cell in rows.cells])
    df = pd.DataFrame(data[1:], columns=data[0])
    df.to_excel("data/samples/tz.xlsx", index=False)

def draw_dates(first_day_str):
    first_day = datetime.strptime(first_day_str, "%d/%m/%Y").date()
    path = "data/samples/scan.xlsx"
    wb = load_workbook(path)
    pos = wb["Посещения"]
    knig = wb["Книгов"]  
    pos.column_dimensions['A'].width = 22
    knig.column_dimensions['A'].width = 22

    dates = [first_day + timedelta(days=i) for i in range(91)]

    for i, dat in enumerate(dates):
        с = pos.cell(row = 5 + i, column = 1, value = dat)
        с.number_format = "DD MMMM YYYY"
        b = knig.cell(row = 6 + i, column = 1, value = dat)
        b.number_format = "DD MMMM YYYY"

    wb.save("data/output/test.xlsx")
    return dates

def draw_holidays(selected):
    global holidays_all
    holidays_all = len(selected)

    path = "data/output/test.xlsx"
    wb = load_workbook(path)

    fill_rows(wb["Посещения"], selected, black_fill)
    fill_rows(wb["Книгов"], selected, black_fill, start_row=6, max_col=25) 

    wb.save("data/output/test.xlsx")

def fill_rows(sheet, selected, black_fill, start_row=5, end_row=97, max_col=27):
    for row in range(start_row, end_row):
        cell = sheet.cell(row=row, column=1)
        value = cell.value
        if value is None:
            continue

        value_date = value.date() if hasattr(value, "date") else value

        if value_date in selected:
            for c in sheet.iter_rows(min_row=row, max_row=row, min_col=1, max_col=max_col):
                for x in c:
                    x.fill = black_fill
                    x.value = 0       


def list_of_values(numb):

    list_before_delta = []
    global holidays_all
    iterations = 91 - holidays_all

    for i in range(iterations):
        mu = numb / iterations 
        sigma = mu * 0.3
        float_value = random.gauss(mu, sigma)
        float_value = max(0, float_value)
        value = round(float_value)
        list_before_delta.append(value)

    total = sum(list_before_delta)
    delta = numb - total
    last_cell = len(list_before_delta) - 1
    plus = delta > 0

    while delta != 0:
        cell_index = random.randint(0, last_cell)
        cell = list_before_delta[cell_index]
        if plus:
            list_before_delta[cell_index] += 1
            delta -= 1
        else:
            if cell == 0:
                continue
            else:
                list_before_delta[cell_index] -= 1
                delta += 1

    return list_before_delta

def center():
    path_to_doc = "data/output/test.xlsx"
    path_to_tz = "data/samples/tz.xlsx"

    wb = load_workbook(path_to_doc)
    pos = wb["Посещения"]
    knig = wb["Книгов"]

    tz1 = load_workbook(path_to_tz)
    tz = tz1.active

    # Число пользователей:
    all_users = list_of_values(int(tz['E2'].value))
    insert_into_collumn(5, 'B', pos, all_users)

    wb.save(path_to_doc)


def insert_into_collumn(first_row_index, column, sheet, list_of_values):
    for i in range(91):
        cell_adress = f"{column}{i + first_row_index}"
        if sheet[cell_adress].value != 0:
            sheet[cell_adress] = list_of_values[i]
center()