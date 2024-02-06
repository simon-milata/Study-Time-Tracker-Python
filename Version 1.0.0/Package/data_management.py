import os
import datetime

import openpyxl as op

class DataManager:
    def __init__(self, timer_manager, workbook, worksheet):
        self.initialize_variables()

        self.timer_manager = timer_manager
        self.workbook = workbook
        self.worksheet = worksheet


    def initialize_variables(self):
        self.timer_time = 0
        self.break_time = 0
        self.data_amount = 0
        self.stop_time = ""


    def save_data(self, timer_button, break_button, time_display_label, break_display_label):
        if self.timer_manager.timer_time < 60:
            return print("No data to save.")
        
        self.timer_time = self.timer_manager.timer_time
        self.break_time = self.timer_manager.break_time

        self.duration = self.calculate_duration()

        self.stop_time = datetime.datetime.now()

        timer_button.configure(text="Start")
        break_button.configure(text="Start")
        time_display_label.configure(text="0:00:00")
        break_display_label.configure(text="0:00:00")
        print(self.timer_time, self.break_time)
        self.timer_manager.initialize_variables()

    def calculate_duration(self):
        duration = self.timer_time - self.break_time
        if duration < 0: 
            duration = 0
        else:
            duration /= 60
        return duration
 