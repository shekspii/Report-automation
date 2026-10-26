import pandas as pd
from datetime import date, datetime, timedelta
from openpyxl import load_workbook

def draw_dates(first_day_str):
    first_day = datetime.strptime(first_day_str, "%d/%m/%Y").date()
    path = "data/samples/scan.xlsx"
    wb = load_workbook(path)
    pos = wb["Посещения"]  
    pos.column_dimensions['A'].width = 22

    dates = [first_day + timedelta(days=i) for i in range(91)]

    for i, dat in enumerate(dates):
        с = pos.cell(row = 5 + i, column = 1, value = dat)
        с.number_format = "DD MMMM YYYY"

    wb.save("data/output/test.xlsx")