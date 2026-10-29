import customtkinter as ctk
from ctkdateentry import CTkDateEntry, CTkStringVar
from my_tools import draw_dates, draw_holidays, word_to_excel
from datetime import datetime, date
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
from doc2docx import convert as doc2docx_convert

class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()

        self.TkdndVersion = TkinterDnD._require(self)
        self.geometry("900x800")
        self.title("Автоматизация отчётности")

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)   
        self.grid_rowconfigure(1, weight=0)   

        left = ctk.CTkFrame(self, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)  

        self.input_date = CTkStringVar(self, value="Введите первый день квартала")
        self.input_date.trace_add("write", self.check_another_value_of_date_input)

        self.calendar = CTkDateEntry(left, variable=self.input_date, width=240)
        self.calendar.grid(row=0, column=0, pady=(0, 10), padx=10, sticky="ew")

        self.scroll_frame = ctk.CTkScrollableFrame(
            left, width=280, height=250, label_text="Выберите выходные дни"
        )
        self.scroll_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10), padx=10)

        self.checkboxes = {}

        right = ctk.CTkFrame(self, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1) 

        self.drop_area = ctk.CTkLabel(
            right,
            text="Перетащите Word-файл за квартал сюда",
            fg_color=("gray85", "gray20"),   # светлая/тёмная тема
            text_color=("gray10", "white"),  # явный цвет текста
            corner_radius=10,
            height=100,
        )
        self.drop_area.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind("<<Drop>>", self.on_file_drop)

        self.file_order = FileOrderFrame(right)
        self.file_order.grid(row=1, column=0, sticky="nsew")

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
        bottom.grid_columnconfigure(0, weight=1)
        bottom.grid_columnconfigure(1, weight=1)

        self.button_fill_holidays = ctk.CTkButton(
            bottom, text="Заполнить выходные дни", command=self.scroll_button
        )
        self.button_fill_holidays.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.button_fill_table = ctk.CTkButton(
            bottom, text="Заполнить таблицу", command=self.on_submit
        )
        self.button_fill_table.grid(row=0, column=1, padx=(5, 0), sticky="ew")

    def on_file_drop(self, event):
        raw_paths = self.tk.splitlist(event.data)
        if not raw_paths:
            self.drop_area.configure(text="Нужны файлы")
            return

        current = self.file_order.get_files()
        added = 0
        errors = []

        for path in raw_paths:
            lower = path.lower()

            if lower.endswith(".docx"):
                if path not in current:
                    current.append(path)
                    added += 1

            elif lower.endswith(".doc"):
                try:
                    base, _ = os.path.splitext(path)
                    temp_docx = base + "_converted.docx"
                    doc2docx_convert(path, temp_docx)
                    if temp_docx not in current:
                        current.append(temp_docx)
                        added += 1
                except Exception as e:
                    errors.append(f"{os.path.basename(path)}: {e}")

            else:
                errors.append(f"{os.path.basename(path)}: не Word-файл")

        self.file_order.set_files(current)

        if errors:
            self.drop_area.configure(text=f"Ошибка: {errors[0]}")
        else:
            self.drop_area.configure(text=f"Загружено файлов: {len(current)} (добавлено {added})")

    def on_submit(self):
        paths = self.file_order.get_files()
        if not paths:
            return
        for i, path in enumerate(paths):
            word_to_excel(i, path)
            
       
    def scroll_button(self):
        selected = [d for d, var in self.checkboxes.items() if var.get()]
        draw_holidays(selected)
    
    def date_button(self, *args):
        chosen_date = self.input_date.get()
        days = draw_dates(chosen_date)
        self.update_scroll_bar(days)

    def update_scroll_bar(self, days):
        for i, dat in enumerate(days):
            var = ctk.BooleanVar(value=False)
            one_date = ctk.CTkCheckBox(
                self.scroll_frame,
                text=dat.strftime("%d.%m.%Y (%a)"),
                variable=var 
            )
            one_date.grid(row=i, column=0, sticky="w", pady=2)
            self.checkboxes[dat] = var


    def check_another_value_of_date_input(self, *args):
        current_value = self.input_date.get()
        data = datetime.strptime(current_value, "%d/%m/%Y").date()
        if current_value and type(data) == date:
            self.date_button(data)

class FileOrderFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.file_paths = []          # текущий порядок файлов
        self._row_widgets = []        # виджеты строк (для перерисовки)

        self.title_label = ctk.CTkLabel(
            self,
            text="Порядок обработки файлов",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.title_label.pack(anchor="w", padx=10, pady=(10, 4))

        self.instruction = ctk.CTkLabel(
            self,
            text=(
                "Расположите файлы в нужном порядке:\n"
                "1-й — Контрольные показатели за квартал,\n"
                "2-й — Справки и консультации за квартал ,\n"
                "3-й — Отчетные показатели.\n"
                "Используйте стрелки для изменения позиции."
            ),
            justify="left",
            anchor="w",
            text_color="gray70",
        )
        self.instruction.pack(anchor="w", padx=10, pady=(0, 8))

        self.rows_container = ctk.CTkFrame(self, fg_color="transparent")
        self.rows_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.empty_label = ctk.CTkLabel(
            self.rows_container,
            text="Файлы ещё не добавлены",
            text_color="gray50",
        )
        self.empty_label.pack(pady=20)

    def set_files(self, paths):
        """Установить список путей (например, после drop)."""
        self.file_paths = list(paths)
        self._redraw()

    def get_files(self):
        """Получить текущий порядок."""
        return list(self.file_paths)

    def _redraw(self):
        for w in self.rows_container.winfo_children():
            w.destroy()
        self._row_widgets.clear()

        if not self.file_paths:
            self.empty_label = ctk.CTkLabel(
                self.rows_container,
                text="Файлы ещё не добавлены",
                text_color="gray50",
            )
            self.empty_label.pack(pady=20)
            return

        for idx, path in enumerate(self.file_paths):
            TEXT_COLOR = ("gray10", "white")

            row = ctk.CTkFrame(self.rows_container, fg_color="transparent", border_color=("gray70", "gray40"), corner_radius=8)
            row.pack(fill="x", pady=4)

            # номер позиции
            num = ctk.CTkLabel(row, text=f"{idx + 1}.", text_color=TEXT_COLOR, width=24, anchor="w")
            num.pack(side="left", padx=(8, 4), pady=6)

            # имя файла (обрезанное, если длинное)
            name = path.split("/")[-1]
            name_label = ctk.CTkLabel(row, text=name, anchor="w", text_color=TEXT_COLOR, justify="left")
            name_label.pack(side="left", fill="x", expand=True, padx=4, pady=6)

            # кнопки
            up_btn = ctk.CTkButton(
                row,
                text="↑",
                width=32,
                height=28,
                state="normal" if idx > 0 else "disabled",
                command=lambda i=idx: self._move_up(i),
            )
            up_btn.pack(side="left", padx=2)

            down_btn = ctk.CTkButton(
                row,
                text="↓",
                width=32,
                height=28,
                state="normal" if idx < len(self.file_paths) - 1 else "disabled",
                command=lambda i=idx: self._move_down(i),
            )
            down_btn.pack(side="left", padx=(2, 8))

    def _move_up(self, index):
        if index > 0:
            self.file_paths[index], self.file_paths[index - 1] = (
                self.file_paths[index - 1],
                self.file_paths[index],
            )
            self._redraw()

    def _move_down(self, index):
        if index < len(self.file_paths) - 1:
            self.file_paths[index], self.file_paths[index + 1] = (
                self.file_paths[index + 1],
                self.file_paths[index],
            )
            self._redraw()

app = App()
app.mainloop()