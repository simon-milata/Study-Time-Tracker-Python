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
        self.timer_time, self.break_time = self.timer_manager.get_time_data()
        if self.timer_time < 60:
            print("No data to save.")
            return

        self.stop_time = datetime.datetime.now()

        timer_button.configure(text="Start")
        break_button.configure(text="Start")
        time_display_label.configure(text="0:00:00")
        break_display_label.configure(text="0:00:00")
        print(self.timer_time, self.break_time)
 