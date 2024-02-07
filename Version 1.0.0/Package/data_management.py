import os
import datetime

import openpyxl as op
from openpyxl.styles import Font

class DataManager:
    def __init__(self, App, timer_manager, workbook, worksheet):
        self.app = App
        self.timer_manager = timer_manager
        self.workbook = workbook
        self.worksheet = worksheet
        self.initialize_variables()


    def initialize_variables(self):
        self.date_list = []
        self.duration_list = []
        self.total_duration = 0
        

    def initialize_new_file_variables(self):
        self.goal_amount = 0
        self.data_amount = 0

    
    def collect_data(self):
        self.data_amount = int(self.worksheet["Z2"].value)
        self.goal_amount = int(self.worksheet["R2"].value)

        print("Data collected.")


    def data_to_variable(self):
        for data in range(2, self.data_amount + 2):
            if "/" in str(self.worksheet["B" + str(data)].value):
                self.date_list.append(datetime.datetime.strptime(str(self.worksheet["B" + str(data)].value).split(" ")[0], "%d/%m/%Y").date())
            elif "-" in str(self.worksheet["B" + str(data)].value):
                self.date_list.append(datetime.datetime.strptime(str(self.worksheet["B" + str(data)].value).split(" ")[0], "%Y-%m-%d").date())
            self.duration_list.append(round(self.worksheet["C" + str(data)].value))

        self.total_duration = sum(self.duration_list)


    def save_data(self):
        self.initialize_variables()
        
        self.data_amount += 1
        
        self.timer_time = self.timer_manager.timer_time
        self.break_time = self.timer_manager.break_time

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
        self.worksheet["D" + str((self.data_amount + 1))].value = self.break_time/60
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

        self.worksheet["A1"].font = Font(bold=True, size=14)
        self.worksheet["B1"].font = Font(bold=True, size=14)
        self.worksheet["C1"].font = Font(bold=True, size=14)
        self.worksheet["D1"].font = Font(bold=True, size=14)
        self.worksheet["E1"].font = Font(bold=True, size=14)

        self.worksheet["Z1"].value = "Data amount: "
        self.worksheet["Z1"].font = Font(bold=True, size=14)
        self.worksheet["Z2"].value = self.data_amount
        self.workbook.save(self.app.data_file)
        print("Excel customized.")


    def get_start_time(self):
        self.start_time = datetime.datetime.now()


    def calculate_duration(self):
        duration = self.timer_time - self.break_time
        if duration < 0: 
            duration = 0
        else:
            duration /= 60
        return duration
    

    def increase_goal_streak(self):
        print("AAA")
        self.goal_amount += 1
        print(self.goal_amount)

    
    def reset_data(self, workbook, worksheet):
        self.workbook = workbook
        self.worksheet = worksheet

        self.workbook.save(self.app.data_file)

        self.initialize_new_file_variables()

        print("Data reset.")