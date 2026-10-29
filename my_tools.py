import pandas as pd
from datetime import date, datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from docx import Document
import random

black_fill = PatternFill(fill_type="solid", fgColor="000000")
holidays_all = 0

def word_to_excel(number, path):
    doc = Document(path)
    table = doc.tables[0]
    data = []
    for rows in table.rows:
        data.append([cell.text for cell in rows.cells])
    df = pd.DataFrame(data[1:], columns=data[0])
    df.to_excel(f"data/samples/tz{number}.xlsx", index=False)

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

    center()

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
    # print(holidays_all, iterations, numb, holidays_all)
    mu = numb / iterations 
    sigma = mu * 0.3

    for i in range(iterations):
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
    path_to_tz = "data/samples/tz0.xlsx"
    path_to_tz1 = "data/samples/tz1.xlsx"

    wb = load_workbook(path_to_doc)
    pos = wb["Посещения"]
    knig = wb["Книгов"]

    tz1ist = load_workbook(path_to_tz)
    tz = tz1ist.active

    tz_small = load_workbook(path_to_tz1)
    tz1 = tz_small.active

    ## Число пользователей
    # - обслуженных в стационарных условиях 
    all_users = list_of_values(int(tz['E2'].value))
    insert_into_collumn(5, 'B', pos, all_users)

    # - дети до 14 лет включительно
    human_14 = list_of_values(int(tz['E3'].value))
    insert_into_collumn(5, 'J', pos, human_14)

    # "- молодежь 15-17 лет"
    young_humans = list_of_values(int(tz['E4'].value))
    insert_into_collumn(5, 'K', pos, young_humans)

    # "- молодежь18-35 лет"
    mid_humans = list_of_values(int(tz['E5'].value))
    insert_into_collumn(5, 'L', pos, mid_humans)

    ## Общее кол-во книговыдач за отчетный период
    # -в стационарных условиях 
    all_books = list_of_values(int(tz['E8'].value))
    insert_into_collumn(6, 'B', knig, all_books)

    # -детям до 14 лет включительно
    for_human_14_books = list_of_values(int(tz['E9'].value))
    insert_into_collumn(6, 'L', knig, for_human_14_books)

    # "-молодежи 15-17 лет"
    for_young_humans_books = list_of_values(int(tz['E10'].value))
    insert_into_collumn(6, 'M', knig, for_young_humans_books)

    # "молодежи 18-35 лет"
    for_mid_humans_books = list_of_values(int(tz['E11'].value))
    insert_into_collumn(6, 'N', knig, for_mid_humans_books)

    ## Общее количество посещений 
    # - в стационарных условиях 
    all_visits = list_of_values(int(tz['D14'].value))
    insert_into_collumn(5, 'U', pos, all_visits)

    # - дети до 14 лет включительно
    human_14_visits = list_of_values(int(tz['D15'].value))
    insert_into_collumn(5, 'X', pos, human_14_visits)

    # Число обращений удаленных пользователей
    request_of_delete_users = list_of_values(int(tz['D17'].value))
    insert_into_collumn(5, 'T', pos, request_of_delete_users)

    # "Количество посещений библиотечных мероприятий всего"
    all_visits_in_events = list_of_values(int(tz['D22'].value))
    insert_into_collumn(5, 'U', pos, all_visits_in_events)

    # Число посещений с мероприятий детям до 14 лет
    events_human_14_visits = list_of_values(int(tz['D24'].value))
    insert_into_collumn(5, 'X', pos, events_human_14_visits)

    # 15-17
    for_other_ages = (int(tz['D22'].value) - int(tz['D24'].value)) // 2
    events_young_humans_visits = list_of_values(for_other_ages)
    insert_into_collumn(5, 'Y', pos, events_young_humans_visits)

    # 18-35
    events_mid_humans_visits = list_of_values(for_other_ages)
    insert_into_collumn(5, 'Z', pos, events_mid_humans_visits)

    # ОМиЭР
    # Всего
    all_consultations = list_of_values(int(tz1['B6'].value))
    insert_into_collumn(6, 'U', knig, all_consultations)

    # В.т.ч. удаленные
    deleted_consultations = list_of_values(int(tz1['C6'].value))
    insert_into_collumn(6, 'X', knig, deleted_consultations)

    # "В.т.ч. детям До 14"
    humans_14_consultations = list_of_values(int(tz1['D6'].value))
    insert_into_collumn(6, 'V', knig, deleted_consultations)

    # "В.т.ч. детям До 14"
    young_humans_consultations = list_of_values(int(tz1['E6'].value))
    insert_into_collumn(6, 'W', knig, young_humans_consultations)

    wb.save(path_to_doc)


def insert_into_collumn(first_row_index, column, sheet, list_of_values):
    index = 0
    for i in range(91):
        cell_adress = f"{column}{i + first_row_index}"
        if sheet[cell_adress].value != 0:
            sheet[cell_adress] = (sheet[cell_adress].value or 0) + list_of_values[index]
            index += 1

def clear_data():
    path = "data/samples/scan.xlsx"
    wb = load_workbook(path)

    path_test = "data/output/test.xlsx"
    wbb = load_workbook(path_test)

    wbb = wb
    wbb.save(path_test)
# clear_data()