import pandas as pd
from datetime import date, datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from docx import Document
import random, shutil, numpy, json, sys
from pathlib import Path
from platformdirs import user_desktop_dir, user_documents_path

black_fill = PatternFill(fill_type="solid", fgColor="000000")
holidays_all = 0

base = user_documents_path() / "ReportAutomation"
base.mkdir(parents=True, exist_ok=True)

JSON_PATH = base / "data" / "samples" / "events_counter.json"
JSON_PATH.parent.mkdir(parents=True, exist_ok=True)

def word_to_excel(number, path):
    doc = Document(path)
    table = doc.tables[0]
    data = []
    for rows in table.rows:
        data.append([cell.text for cell in rows.cells])
    df = pd.DataFrame(data[1:], columns=data[0])

    out_dir = base / "data" / "samples"
    out_dir.mkdir(parents=True, exist_ok=True)      
    out_path = out_dir / f"tz{number}.xlsx"

    df.to_excel(out_path, index=False)

def resource_path(relative):
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return Path(__file__).resolve().parent / relative

def draw_dates(first_day_str):
    first_day = datetime.strptime(first_day_str, "%d/%m/%Y").date()
    path = resource_path("data/samples/scan.xlsx")
    wb = load_workbook(path)
    pos = wb["Посещения"]
    knig = wb["Книгов"]  
    pos.column_dimensions['A'].width = 22
    knig.column_dimensions['A'].width = 22

    dates = {}          
    for i in range(91):
        dat = first_day + timedelta(days=i)
        dates[dat] = 1 if (dat.weekday() == 5 or dat.day == 15) else 0

        c = pos.cell(row=5 + i, column=1, value=dat)
        c.number_format = "DD MMMM YYYY"
        b = knig.cell(row=6 + i, column=1, value=dat)
        b.number_format = "DD MMMM YYYY"

    out_dir = base / "data" / "output"
    out_dir.mkdir(parents=True, exist_ok=True)      
    out_path = out_dir / "test.xlsx"

    wb.save(out_path)
    return dates

def return_all_days():
    path = base / "data" / "output" / "test.xlsx"
    wb = load_workbook(path)
    pos = wb["Посещения"]

    dates = {}
    for i in range(5, 97):
        value = pos[f"A{i}"].value
        if value is None:
            continue
        dates[value] = 0
        
    return dates

def draw_holidays(selected):
    global holidays_all
    holidays_all = len(selected)

    path = base / "data" / "output" / "test.xlsx"
    wb = load_workbook(path)

    fill_rows(wb["Посещения"], selected, black_fill)
    fill_rows(wb["Книгов"], selected, black_fill, start_row=6, max_col=25) 

    wb.save(path)

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


def list_of_values(numb, max_iterations=91, is_max=False, min_list : list | None = None):
    global holidays_all
    iterations = max_iterations - holidays_all
    mu = numb / iterations 
    sigma = mu * 0.3
    list_before_delta = []

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

def normalization(main_data: list, lists: list[list]):
    # Работаем с копией главного списка, чтобы не мутировать его случайно
    # (если нужно менять исходный, уберите [:])
    res_main = main_data[:]
    
    n = len(res_main)
    
    # Считаем, сколько всего данных в зависимых списках (они неизменны)
    # и сколько в главном
    total_lists_sum = sum(sum(lst) for lst in lists)
    total_main_sum = sum(res_main)
    
    if total_lists_sum > total_main_sum:
        raise ValueError("Невозможно распределить данные: общая сумма зависимых списков больше, чем сумма главного списка.")

    def col_sum(idx):
        return sum(lst[idx] for lst in lists)

    # Шаг 1: Находим дефицитные и профицитные индексы
    deficits = []  # (индекс, сколько не хватает)
    excesses = []  # (индекс, сколько можно забрать)
    
    for i in range(n):
        current_sum = col_sum(i)
        if current_sum > res_main[i]:
            deficits.append((i, current_sum - res_main[i]))
        elif res_main[i] > current_sum:
            excesses.append((i, res_main[i] - current_sum))

    # Шаг 2: Переливаем значения из профицитных индексов в дефицитные
    # Работаем с копиями списков кортежей, чтобы не запутаться с индексами
    d_idx = 0
    e_idx = 0
    
    while d_idx < len(deficits) and e_idx < len(excesses):
        d_i, d_val = deficits[d_idx]
        e_i, e_val = excesses[e_idx]
        
        # Сколько можем перекинуть за один шаг
        transfer = min(d_val, e_val)
        
        # Увеличиваем главный список в дефицитном индексе
        res_main[d_i] += transfer
        # Уменьшаем главный список в профицитном индексе
        res_main[e_i] -= transfer
        
        # Обновляем остатки
        deficits[d_idx] = (d_i, d_val - transfer)
        excesses[e_idx] = (e_i, e_val - transfer)
        
        # Если дефицит закрыт — переходим к следующему
        if deficits[d_idx][1] == 0:
            d_idx += 1
        # Если профицит исчерпан — переходим к следующему
        if excesses[e_idx][1] == 0:
            e_idx += 1

    # Возвращаем в том же порядке: сначала (измененный) main_data, затем исходные lists
    return res_main, lists

def reverse_normalization(main_data: list, lists: list[list]):
    # Работаем с копиями, чтобы не мутировать исходные данные
    res_lists = [lst[:] for lst in lists]
    n = len(main_data)

    def col_sum(idx):
        return sum(lst[idx] for lst in res_lists)

    # 1. Находим все дефицитные индексы
    deficits = []
    for i in range(n):
        current_sum = col_sum(i)
        if current_sum < main_data[i]:
            deficits.append([i, main_data[i] - current_sum]) # [индекс, сколько нужно добрать]

    # 2. Проходим по каждому дефициту и ищем, откуда взять числа
    for d_i, need in deficits:
        # Ищем профицитные индексы
        for e_i in range(n):
            if need == 0:
                break
            if e_i == d_i:
                continue
            
            current_e_sum = col_sum(e_i)
            
            # Если в этом индексе у зависимых списков есть числа,
            # и мы можем их забрать (например, там в main_data 0, или сумма с избытком)
            if current_e_sum > 0 and (main_data[e_i] == 0 or current_e_sum > main_data[e_i]):
                
                # Проходим по каждому зависимому списку
                for lst in res_lists:
                    if need == 0:
                        break
                    
                    # Сколько можем забрать из этого конкретного списка в этом индексе
                    # (забираем минимум из того, что есть в списке, и того, сколько нужно)
                    take = min(lst[e_i], need)
                    
                    if take > 0:
                        lst[e_i] -= take  # Убираем из профицитного индекса
                        lst[d_i] += take  # Добавляем в дефицитный
                        need -= take

    return main_data, res_lists


def parse_events_json(pos):
    # 1. Загружаем JSON
    try:
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            events = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        events = {}

    # 2. Сортируем даты (хотя для поиска по ключу это не обязательно, 
    # но полезно, если выводить их на печать для проверки)
    sorted_dates = sorted(events.keys())

    # 3. Готовим списки
    adults_list = []
    children_list = []

    # 4. Идем по датам в Excel с шагом 91
    for i in range(91):
        cell_value = pos[f"A{5 + i}"].value
        b_value = pos[f"B{5 + i}"].value
        
        if b_value is None:
            # Если ячейка пустая, пишем нули
            adults_list.append(0)
            children_list.append(0)
            continue
        elif b_value == 0:
                continue
        # Приводим дату из Excel к формату ключа JSON ("YYYY-MM-DD")
        if hasattr(cell_value, "strftime"):
            date_key = cell_value.strftime("%Y-%m-%d")
        else:
            # Если вдруг в ячейке строка, пытаемся её распарсить
            try:
                dt = datetime.strptime(str(cell_value).split()[0], "%d.%m.%Y") # или другой формат
                date_key = dt.strftime("%Y-%m-%d")
            except:
                date_key = str(cell_value)

        # 5. Проверяем, есть ли эта дата в JSON
        if date_key in events:
            day_events = events[date_key]
            
            # Инициализируем нулями на случай, если в этот день только одно мероприятие
            adult_count = 0
            child_count = 0

            # Проходим по всем мероприятиям этого дня
            for item in day_events:
                if item.get("is_min") is False:
                    adult_count = item.get("count", 0)
                elif item.get("is_min") is True:
                    child_count = item.get("count", 0)

            adults_list.append(adult_count)
            children_list.append(child_count)
        else:
            # Даты нет в JSON — пишем нули в оба списка
            adults_list.append(0)
            children_list.append(0)

    return adults_list, children_list

def center():
    path_to_doc = base / "data" / "output" / "test.xlsx"
    path_to_tz = base / "data" / "samples" / "tz0.xlsx"
    path_to_tz1 = base / "data" / "samples" / "tz1.xlsx"

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
    human_14 = list_of_values(int(tz['E3'].value))
    young_humans = list_of_values(int(tz['E4'].value))
    mid_humans = list_of_values(int(tz['E5'].value))

    fix_all_users, fix_others_guys = normalization(all_users, [human_14, young_humans, mid_humans])

    insert_into_collumn(5, 'B', pos, fix_all_users)

    # - дети до 14 лет включительно
    insert_into_collumn(5, 'J', pos, fix_others_guys[0])

    # "- молодежь 15-17 лет"
    insert_into_collumn(5, 'K', pos, fix_others_guys[1])

    # "- молодежь18-35 лет"
    insert_into_collumn(5, 'L', pos, fix_others_guys[2])

    ## Общее кол-во книговыдач за отчетный период
    # -в стационарных условиях 
    all_books = list_of_values(int(tz['E8'].value))
    for_human_14_books = list_of_values(int(tz['E9'].value))
    for_young_humans_books = list_of_values(int(tz['E10'].value))
    for_mid_humans_books = list_of_values(int(tz['E11'].value))

    fix_all_books, fix_other_books = normalization(all_books, [for_human_14_books, for_young_humans_books, for_mid_humans_books])

    # Вновь поступившие
    insert_into_collumn(6, 'B', knig, fix_all_books)

    # -детям до 14 лет включительно
    insert_into_collumn(6, 'L', knig, fix_other_books[0])

    # "-молодежи 15-17 лет"
    insert_into_collumn(6, 'M', knig, fix_other_books[1])

    # "молодежи 18-35 лет"
    insert_into_collumn(6, 'N', knig, fix_other_books[2])

    ## Общее количество посещений 
    # - в стационарных условиях 
    all_visits = list_of_values(int(tz['D14'].value))
    human_14_visits = list_of_values(int(tz['D15'].value))
    for_other_ages = (int(tz['D14'].value) - int(tz['D15'].value)) // 3.5
    events_young_humans_visits = list_of_values(for_other_ages)
    events_mid_humans_visits = list_of_values(for_other_ages)
    events_old_humans_visits = list_of_values(for_other_ages)

    # "Количество посещений библиотечных мероприятий всего"
    global holidays_all
    list_len = int(tz['D19'].value) + holidays_all
    len_min_events = int(tz["D21"].value) + holidays_all
    len_for_ev_all = list_len-len_min_events + holidays_all

    indexes_array = make_sorted_random_array(int(tz['D19'].value), pos)
    indexes_array1 = numpy.random.choice(indexes_array, size=int(tz['D19'].value) - int(tz['D21'].value), replace=False)
    min_indexes_array = numpy.random.choice(indexes_array, size=int(tz['D21'].value), replace=False)

    if JSON_PATH.exists():
        JSON_PATH.unlink() 

    events_human_14_visits = list_of_values(int(tz['D24'].value), len_min_events)
    all_visits_in_events = list_of_values(int(tz['D22'].value) - int(tz['D24'].value), len_for_ev_all, True, events_human_14_visits)

    # Число посещений с мероприятий
    random_insert_into_collumn(5, 'U', pos, all_visits_in_events, indexes_array1, True)

    # Число посещений с мероприятий детям до 14 лет
    random_insert_into_collumn(5, 'X', pos, events_human_14_visits, min_indexes_array, True, True)

    with open(JSON_PATH, "r", encoding="utf-8") as file:
        events = json.load(file)
    max_count, min_count = parse_events_json(pos)
    print(min_count)
    print(max_count)
    # fix_human_14_visits, new_min_count = normalization(human_14_visits, [min_count])
    # fix_visits, new_counts = normalization(all_visits, [fix_human_14_visits, max_count])

    # fix_all_visits, fix_other_visits = normalization(fix_visits, [fix_human_14_visits, events_young_humans_visits, events_mid_humans_visits])
    human_14_visits, new_min_count = normalization(human_14_visits, [min_count])
    max_count, new_events_visits_15_70 = reverse_normalization(max_count, [events_young_humans_visits, events_mid_humans_visits, events_old_humans_visits])

    fix_all_visits, fix_other_visits = normalization(all_visits, [human_14_visits, new_events_visits_15_70[0], new_events_visits_15_70[1], new_events_visits_15_70[2]])
    
    
    # Посещения
    insert_into_collumn(5, 'U', pos, fix_all_visits)

    # - дети до 14 лет включительно
    insert_into_collumn(5, 'X', pos, fix_other_visits[0])

    # 15-17
    insert_into_collumn(5, 'Y', pos, fix_other_visits[1])

    # 18-35
    insert_into_collumn(5, 'Z', pos, fix_other_visits[2])

    # Число обращений удаленных пользователей
    request_of_delete_users = list_of_values(int(tz['D17'].value))
    insert_into_collumn(5, 'T', pos, request_of_delete_users)


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
    make_txt_report()


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
        else:
            sheet[cell_adress] = (cell_value or 0) + new_value

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
        count = sum(e["count"] for e in result)
        len = sum(1 for e in result)
    else:
        count = sum(e["count"] for e in result if e["is_min"])
        len = sum(1 for e in result if e["is_min"])

    return count, len

def make_txt_report():
    try:
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            events = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        events = {}

    out_path = base / "data" / "output" / "events_report.txt"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with open(out_path, "w", encoding="utf-8") as f:
        for date_key, items in sorted(events.items()):
            try:
                dt = datetime.strptime(date_key, "%Y-%m-%d")
                date_str = dt.strftime("%d.%m.%Y")
            except ValueError:
                date_str = date_key  

            for item in items:
                is_min = item.get("is_min", False)
                count = item.get("count", 0)

                event_type = "Детское" if is_min else "Взрослое"

                line = f"{date_str} — {event_type} — Пришло: {count} чел.\n"
                f.write(line)

    copy_to_desktop(out_path)
    

def find_row_by_value(sheet, target, first_row=5, last_row=97):
    for i in range(first_row, last_row):
        cell = sheet.cell(row=i, column=1)
        if cell.value == target:
            return cell.row
    return None

def make_month_report(selected : list):
    #print("Work!! ")
    path_to_doc = base / "data" / "output" / "test.xlsx"
    path_to_report = resource_path("data/samples/month_report.docx")
    path_to_tz = base / "data" / "samples" / "tz0.xlsx"
    
    wb = load_workbook(path_to_doc)
    tzwb = load_workbook(path_to_tz)

    pos = wb["Посещения"]
    knig = wb["Книгов"]

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


    doc.save(base / "data" / "samples" / "month_report.docx")
    copy_to_desktop(base / "data" / "samples" / "month_report.docx")

#make_month_report([datetime(2026, 11, 13, 0, 0), datetime(2027, 2, 11, 0, 0)])

def copy_to_desktop(source_path):
    source = Path(source_path)
    if not source.exists():
        print(f"Файл не найден: {source}")
        return

    desktop = Path(user_desktop_dir())

    destination = desktop / source.name
    shutil.copy2(source, destination)

def clear_data():
    path_test = base / "data" / "output" / "test.xlsx"
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