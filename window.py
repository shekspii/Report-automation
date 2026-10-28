import customtkinter as ctk
from ctkdateentry import CTkDateEntry, CTkStringVar
from my_tools import draw_dates, draw_holidays, word_to_excel
from datetime import datetime, date
from tkinterdnd2 import TkinterDnD, DND_FILES

class App(ctk.CTk,  TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()

        self.TkdndVersion = TkinterDnD._require(self)
        self.geometry("500x750")
        self.title("Автоматизация отчётности")

        self.input_date = CTkStringVar(self, value = "Введите первый день квартала")
        self.input_date.trace_add("write", self.check_another_value_of_date_input)

        self.calendar = CTkDateEntry(self, variable = self.input_date, width=240)
        self.calendar.pack(pady=10, padx=10)

        self.scroll_frame = ctk.CTkScrollableFrame(self, width=280, height=250, label_text="Выберите выходные дни")
        self.scroll_frame.pack(fill="both", expand=True, pady=10, padx=10)

        self.checkboxes = {}

        self.button = ctk.CTkButton(self, text="Заполнить выходные дни", command=self.scroll_button)
        self.button.pack(pady=10)

        self.drop_area = ctk.CTkLabel(
            self, 
            text="Перетащите Word-файл за квартал сюда",
            fg_color="gray20",
            corner_radius=10,
            height=100
        )
        self.drop_area.pack(fill="x", padx=20, pady=20)

        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind("<<Drop>>", self.on_file_drop)

    def on_file_drop(self, event):
        file_path = self.tk.splitlist(event.data)[0]
        self.drop_area.configure(text=f"Загружен: {file_path}")
        word_to_excel(file_path)
       
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

app = App()
app.mainloop()