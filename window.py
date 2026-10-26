import customtkinter as ctk
from ctkdateentry import CTkDateEntry, CTkStringVar
from my_tools import draw_dates
from datetime import datetime, date

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.geometry("300x100")
        self.title("Автоматизация отчётности")

        self.input_date = CTkStringVar(self, value = "Введите первый день квартала")
        self.input_date.trace_add("write", self.check_another_value_of_date_input)
        self.calendar = CTkDateEntry(self, variable = self.input_date, width=240)
        self.calendar.pack(pady=10, padx=10)
        

    def date_button(self, *args):
        chosen_date = self.input_date.get()
        draw_dates(chosen_date)

    def check_another_value_of_date_input(self, *args):
        current_value = self.input_date.get()
        data = datetime.strptime(current_value, "%d/%m/%Y").date()
        if current_value and type(data) == date:
            self.date_button(data)

app = App()
app.mainloop()