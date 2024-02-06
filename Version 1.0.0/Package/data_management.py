import os
import datetime

import openpyxl as op

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

    
    def collect_data(self):
        self.data_amount = int(self.worksheet["Z1"].value)
        self.goal_amount = int(self.worksheet["R1"].value)


        for data in range(2, self.data_amount + 2):
            if "/" in str(self.worksheet["B" + str(data)].value):
                self.date_list.append(datetime.datetime.strptime(str(self.worksheet["B" + str(data)].value).split(" ")[0], "%d/%m/%Y").date())
            elif "-" in str(self.worksheet["B" + str(data)].value):
                self.date_list.append(datetime.datetime.strptime(str(self.worksheet["B" + str(data)].value).split(" ")[0], "%Y-%m-%d").date())
            self.duration_list.append(round(self.worksheet["C" + str(data)].value))

        self.total_duration = sum(self.duration_list)
        print("Data collected.")
        self.app.update_streak_values()


    def save_data(self):
        if self.timer_manager.timer_time < 60:
            return print("No data to save.")
        
        self.timer_time = self.timer_manager.timer_time
        self.break_time = self.timer_manager.break_time

        self.duration = self.calculate_duration()

        self.stop_time = datetime.datetime.now()

        print("Data saved.")

        self.timer_manager.initialize_variables()
        self.app.reset_timers()



    def calculate_duration(self):
        duration = self.timer_time - self.break_time
        if duration < 0: 
            duration = 0
        else:
            duration /= 60
        return duration
    

    def increase_goal_streak(self):
        self.goal_amount += 1

 