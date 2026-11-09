import pandas as pd
from datetime import date, datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from docx import Document
import random, shutil, numpy, json
from pathlib import Path
from platformdirs import user_desktop_dir

black_fill = PatternFill(fill_type="solid", fgColor="000000")
holidays_all = 0
JSON_PATH = Path("data/samples/events_counter.json")


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

def return_all_days():
    path = "data/output/test.xlsx"
    wb = load_workbook(path)
    pos = wb["Посещения"]

    dates = [pos[f"A{i}"].value for i in range(5, 97)]
    return [d for d in dates if d is not None]

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
            for c in sheet.iter_rows(min_row=row, max_row=row, min_col=2, max_col=max_col):
                for x in c:
                    x.fill = black_fill
                    x.value = 0       


def list_of_values(numb, max_iterations=91):

    list_before_delta = []
    global holidays_all
    iterations = max_iterations - holidays_all
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
    print(len(list_before_delta))
    return list_before_delta

def make_sorted_random_array(list_len : int, pos):
    valid_indexes = []

    for i in range(0, 92):
        cell_address = f"B{i + 5}"
        cell_value = pos[cell_address].value
        if cell_value != 0 and cell_value is not None:
            valid_indexes.append(i)

    indexes_array = numpy.random.choice(valid_indexes, size=list_len, replace=False)
    indexes_array.sort()
    return indexes_array


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
    
    global holidays_all
    list_len = int(tz['D19'].value) + holidays_all
    len_min_events = int(tz["D21"].value) + holidays_all

    indexes_array = make_sorted_random_array(list_len, pos)
    min_indexes_array = numpy.random.choice(indexes_array, size=len_min_events, replace=False)

    if JSON_PATH.exists():
        JSON_PATH.unlink() 

    all_visits_in_events = list_of_values(int(tz['D22'].value), list_len)
    random_insert_into_collumn(5, 'U', pos, all_visits_in_events, indexes_array, True)

    # Число посещений с мероприятий детям до 14 лет
    events_human_14_visits = list_of_values(int(tz['D24'].value), len_min_events)
    random_insert_into_collumn(5, 'X', pos, events_human_14_visits, min_indexes_array, True, True)

    # 15-17
    for_other_ages = (int(tz['D22'].value) - int(tz['D24'].value)) // 2

    events_young_humans_visits = list_of_values(for_other_ages, list_len)
    random_insert_into_collumn(5, 'Y', pos, events_young_humans_visits, indexes_array)

    # 18-35
    events_mid_humans_visits = list_of_values(for_other_ages, list_len)
    random_insert_into_collumn(5, 'Z', pos, events_mid_humans_visits, indexes_array)

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
    copy_to_desktop(path_to_doc)


def insert_into_collumn(first_row_index, column, sheet, list_of_values, list_len=91):
    index = 0
    for i in range(list_len):
        cell_adress = f"{column}{i + first_row_index}"
        if sheet[cell_adress].value != 0:
            sheet[cell_adress] = (sheet[cell_adress].value or 0) + list_of_values[index]
            index += 1


def random_insert_into_collumn(first_row_index, column, sheet, list_of_values, indexes_array, is_log=False, is_min=False):
    index = 0
    try:
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            events = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        events = {}

    for i in indexes_array:
        cell_adress = f"{column}{i + first_row_index}"
        cell_value = sheet[cell_adress].value
        if index >= len(list_of_values):
            break
        new_value = list_of_values[index]
        sheet[cell_adress] = (cell_value or 0) + new_value
        if is_log:
            date_cell = sheet[f"A{i + first_row_index}"].value
            if hasattr(date_cell, "strftime"):
                date_key = date_cell.strftime("%Y-%m-%d")
            else:
                date_key = str(date_cell)

            events.setdefault(date_key, []).append({
                "is_min": is_min,   
                "count": new_value,
            })

            

        index += 1

    JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(events, f, ensure_ascii=False, indent=2)

def copy_from_column(first_row_index, last_row_index, column, sheet):

    count = 0
    for i in range(first_row_index, last_row_index + 1):
        cell_adress = f"{column}{i}"
        count += sheet[cell_adress].value

    return count

def get_from_json(start, end, is_min : bool):
    with open(JSON_PATH, "r", encoding="utf-8") as file:
        events = json.load(file)
    start_str = start.strftime("%Y-%m-%d") if hasattr(start, "strftime") else str(start)
    end_str = end.strftime("%Y-%m-%d") if hasattr(end, "strftime") else str(end)

    result = []
    for date_key, items in events.items():
        if start_str <= date_key <= end_str:
            result.extend(items)
    count = 0
    len = 0
    if not is_min:
        count = sum(e["count"] for e in result if not e["is_min"])
        len = sum(1 for e in result if not e["is_min"])
    else:
        count = sum(e["count"] for e in result if e["is_min"])
        len = sum(1 for e in result if e["is_min"])

    return count, len


def find_row_by_value(sheet, target, first_row=5, last_row=97):
    for i in range(first_row, last_row):
        cell = sheet.cell(row=i, column=1)
        if cell.value == target:
            return cell.row
    return None

def make_month_report(selected : list):
    print("Work!! ")
    path_to_doc = "data/output/test.xlsx"
    path_to_report = "data/samples/month_report.docx"
    path_to_tz = "data/samples/tz0.xlsx"
    
    wb = load_workbook(path_to_doc)
    tzwb = load_workbook(path_to_tz)

    pos = wb["Посещения"]
    knig = wb["Книгов"]
    tz = tzwb.active

    first_day = selected[0]
    last_day = selected[1]

    first_day_row_index = find_row_by_value(pos, first_day)
    last_day_row_index = find_row_by_value(pos, last_day)

    doc = Document(path_to_report)
    table = doc.tables[0]

    # Читатели – Посещения, столбец B, вновь поступившие. J, K, L
    readers = copy_from_column(first_day_row_index, last_day_row_index, "B", pos) 
    kids_readers = copy_from_column(first_day_row_index, last_day_row_index, "J", pos)
    young_readers = copy_from_column(first_day_row_index, last_day_row_index, "K", pos)
    mid_readers = copy_from_column(first_day_row_index, last_day_row_index, "L", pos)

    table.cell(1, 1).text = str(readers)
    table.cell(3, 1).text = str(kids_readers)
    table.cell(4, 1).text = str(young_readers)
    table.cell(5, 1).text = str(mid_readers)

    # Посещения – Посещения, столбец U, Посещения. X.
    visits = copy_from_column(first_day_row_index, last_day_row_index, "U", pos)
    kids_visits = copy_from_column(first_day_row_index, last_day_row_index, "X", pos)

    table.cell(7, 1).text = str(visits)
    table.cell(9, 1).text = str(kids_visits)

    # Книговыдача – Книгов, столбец B, кнв всего. L, M, N
    book_lending = copy_from_column(first_day_row_index + 1, last_day_row_index + 1, "B", knig)
    book_lending_kids = copy_from_column(first_day_row_index + 1, last_day_row_index + 1, "L", knig)
    book_lending_youngs = copy_from_column(first_day_row_index + 1, last_day_row_index + 1, "M", knig)
    book_lending_mid = copy_from_column(first_day_row_index + 1, last_day_row_index + 1, "N", knig)

    table.cell(11, 1).text = str(book_lending)
    table.cell(13, 1).text = str(book_lending_kids)
    table.cell(14, 1).text = str(book_lending_youngs)
    table.cell(15, 1).text = str(book_lending_mid)

    # Массовые мероприятия – кол.во самих мероприятий, 34
    # детям до 14 – кол.во детских мероприятий, 19

    events_count, events_count_len = get_from_json(first_day, last_day, False)
    child_events_count, child_events_count_len = get_from_json(first_day, last_day, True)

    table.cell(17, 1).text = str(events_count_len)
    table.cell(19, 1).text = str(child_events_count_len)
    
    table.cell(21, 1).text = str(events_count)
    table.cell(23, 1).text = str(child_events_count)

    # Обращения удаленных пользователей – Посещения, столбец T, Обращ. удал. польз.
    requests_od_deleted_users = copy_from_column(first_day_row_index, last_day_row_index, "T", pos)
    table.cell(25, 1).text = str(requests_od_deleted_users)


    doc.save(path_to_report)

#make_month_report([datetime(2026, 11, 13, 0, 0), datetime(2027, 2, 11, 0, 0)])

def copy_to_desktop(source_path: str):
    source = Path(source_path)
    if not source.exists():
        print(f"Файл не найден: {source}")
        return

    desktop = Path(user_desktop_dir())

    destination = desktop / source.name
    shutil.copy2(source, destination)

def clear_data():
    path_test = "data/output/test.xlsx"
    wb = load_workbook(path_test)
    pos = wb["Посещения"]
    knig = wb["Книгов"]

    if pos.max_row >= 5:
        for row in pos.iter_rows(min_row=5, max_row=pos.max_row, min_col=2, max_col=pos.max_column):
            for cell in row:
                cell.value = None
    if knig.max_row >= 6:
        for row in knig.iter_rows(min_row=6, max_row=knig.max_row, min_col=2, max_col=knig.max_column):
            for cell in row:
                cell.value = None

    wb.save(path_test)
#clear_data()