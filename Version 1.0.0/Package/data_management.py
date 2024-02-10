import datetime

from openpyxl.styles import Font
from .styles import *

import customtkinter as ctk

class DataManager:
    def __init__(self, App, timer_manager, workbook, worksheet):
        self.app = App
        self.timer_manager = timer_manager
        self.workbook = workbook
        self.worksheet = worksheet
        self.initialize_variables()


    def initialize_variables(self):
        self.day_name_list = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        self.date_list = []
        self.duration_list = []
        self.total_duration = 0
        self.graph_color = "#f38064"
        self.widget_list = []
        

    def initialize_new_file_variables(self):
        self.goal_amount = 0
        self.data_amount = 0
        self.monday_duration = self.tuesday_duration = self.wednesday_duration = self.thursday_duration = self.friday_duration = self.saturday_duration = self.sunday_duration = 0
        self.color_name = "Orange"
        self.save_color()

    
    def collect_data(self):
        self.data_amount = int(self.worksheet["Z2"].value)
        self.goal_amount = int(self.worksheet["R2"].value)

        self.collect_day_data()

        print("Data collected.")


    def data_to_variable(self):
        self.clear_graph_lists()

        for data in range(2, self.data_amount + 2):
            if "/" in str(self.worksheet["B" + str(data)].value):
                self.date_list.append(datetime.datetime.strptime(str(self.worksheet["B" + str(data)].value).split(" ")[0], "%d/%m/%Y").date())
            elif "-" in str(self.worksheet["B" + str(data)].value):
                self.date_list.append(datetime.datetime.strptime(str(self.worksheet["B" + str(data)].value).split(" ")[0], "%Y-%m-%d").date())
            self.duration_list.append(round(self.worksheet["C" + str(data)].value))

        self.total_duration = sum(self.duration_list)


    def save_data(self):
        self.initialize_variables()
        self.save_weekday()
        self.save_color()
        
        self.data_amount += 1

        self.stop_time = datetime.datetime.now()

        self.duration = self.calculate_duration()

        self.workbook.save(self.app.data_file)

        self.write_to_excel()

        print("Data saved.")

        self.timer_manager.initialize_variables()
        self.app.reset_timers()


    def write_to_excel(self):
        self.worksheet["A" + str((self.data_amount + 1))].value = self.start_time.strftime("%d/%m/%Y %H:%M")
        self.worksheet["B" + str((self.data_amount + 1))].value = self.stop_time.strftime("%d/%m/%Y %H:%M")
        self.worksheet["C" + str((self.data_amount + 1))].value = self.duration
        self.worksheet["D" + str((self.data_amount + 1))].value = self.timer_manager.break_time/60

        self.worksheet["R2"].value = self.goal_amount

        self.worksheet["Z2"].value = self.data_amount
        self.workbook.save(self.app.data_file)


    def customize_excel(self):
        self.worksheet["A1"].value = "Start:"
        self.worksheet["B1"].value = "End:"
        self.worksheet["C1"].value = "Duration:"
        self.worksheet["D1"].value = "Break:"

        self.worksheet["R1"].value = "Goals reached:"
        self.worksheet["R2"].value = self.goal_amount

        self.worksheet["T2"].value = self.color_name

        self.worksheet["W2"].value = self.monday_duration
        self.worksheet["W3"].value = self.tuesday_duration
        self.worksheet["W4"].value = self.wednesday_duration
        self.worksheet["W5"].value = self.thursday_duration
        self.worksheet["W6"].value = self.friday_duration
        self.worksheet["W7"].value = self.saturday_duration
        self.worksheet["W8"].value = self.sunday_duration

        self.worksheet["Z1"].value = "Data amount: "
        self.worksheet["Z1"].font = Font(bold=True, size=14)
        self.worksheet["Z2"].value = self.data_amount

        self.style_excel()

        self.workbook.save(self.app.data_file)
        print("Excel customized.")


    def style_excel(self):
        self.worksheet["A1"].font = Font(bold=True, size=14)
        self.worksheet["B1"].font = Font(bold=True, size=14)
        self.worksheet["C1"].font = Font(bold=True, size=14)
        self.worksheet["D1"].font = Font(bold=True, size=14)
        self.worksheet["E1"].font = Font(bold=True, size=14)


    def get_start_time(self):
        self.start_time = datetime.datetime.now()


    def calculate_duration(self):
        duration = self.timer_manager.timer_time - self.timer_manager.break_time
        if duration < 0: 
            duration = 0
        else:
            duration /= 60
        return duration
    

    def increase_goal_streak(self):
        self.goal_amount += 1

    
    def reset_data(self, workbook, worksheet):
        self.workbook = workbook
        self.worksheet = worksheet

        self.initialize_new_file_variables()
        self.customize_excel()

        self.workbook.save(self.app.data_file)

        print("Data reset.")


    def clear_graph_lists(self):
        self.date_list.clear()
        self.duration_list.clear()


    def collect_day_data(self):
        self.monday_duration = int(self.worksheet["W2"].value)
        self.tuesday_duration = int(self.worksheet["W3"].value)
        self.wednesday_duration = int(self.worksheet["W4"].value)
        self.thursday_duration = int(self.worksheet["W5"].value)
        self.friday_duration = int(self.worksheet["W6"].value)
        self.saturday_duration = int(self.worksheet["W7"].value)
        self.sunday_duration = int(self.worksheet["W8"].value)

        self.day_duration_list = [self.monday_duration, self.tuesday_duration, self.wednesday_duration, self.thursday_duration, self.friday_duration, self.saturday_duration, self.sunday_duration]
    

    def save_weekday(self):
        duration = self.calculate_duration()
        weekday_today = datetime.datetime.now().weekday()

        match weekday_today:
            case 0:
                self.monday_duration += duration
                self.worksheet["W2"].value = self.monday_duration
            case 1:
                self.tuesday_duration += duration
                self.worksheet["W3"].value = self.tuesday_duration
            case 2:
                self.wednesday_duration += duration
                self.worksheet["W4"].value = self.wednesday_duration
            case 3:
                self.thursday_duration += duration
                self.worksheet["W5"].value = self.thursday_duration
            case 4:
                self.friday_duration += duration
                self.worksheet["W6"].value = self.friday_duration
            case 5:
                self.saturday_duration += duration
                self.worksheet["W7"].value = self.saturday_duration
            case 6:
                self.sunday_duration += duration
                self.worksheet["W8"].value = self.sunday_duration

        self.workbook.save(self.app.data_file)
        print("Weekday saved.")


    def set_color(self, color_dropdown):
        self.color_name = color_dropdown.get()
        print("Color set.")
        self.save_color()


    def save_color(self):
        self.worksheet["T2"].value = self.color_name
        self.load_color()


    def load_color(self):
        color_name = self.worksheet["T2"].value
        self.app.color_dropdown.configure(variable=ctk.StringVar(value=color_name))
        colors = {"Orange": [orange_button_color, orange_highlight_color, orange_pie_colors], 
                    "Green": [green_button_color, green_highlight_color, green_pie_colors], 
                    "Blue": [blue_button_color, blue_highlight_color, blue_pie_colors]}
        
        self.color = colors[color_name][0]
        self.highlight_color = colors[color_name][1]
        self.pie_colors = colors[color_name][2]
        self.graph_color = self.color
        print("Color loaded.")
        self.change_color(self.color, self.highlight_color)


    def change_color(self, color, highlight_color):
        for widget in self.widget_list:
            widget.configure(fg_color=color, hover_color=highlight_color)
        self.app.progressbar.configure(progress_color = color)

        self.app.create_graphs()

        print("Color changed.")