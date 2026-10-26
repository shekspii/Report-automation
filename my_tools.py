import pandas as pd
from datetime import date, datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

black_fill = PatternFill(fill_type="solid", fgColor="000000")


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
                    x.value = None        