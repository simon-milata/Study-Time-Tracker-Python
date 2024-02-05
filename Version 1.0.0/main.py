import datetime
import os
import random

import openpyxl as op
from openpyxl.styles import Font
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk
from matplotlib.ticker import MaxNLocator, FuncFormatter
from winotify import Notification

from Package import *



class App:
    def __init__(self):
        self.file_setup()
        self.window_setup()
        self.create_gui()

        self.timer_manager = TimerManager(self.WINDOW)
        self.data_manager = DataManager(self.timer_manager, self.workbook, self.worksheet)


    def create_gui(self):
        self.main_frame_gui_setup()
        self.timer_frame_gui_setup()
        self.timer_gui()
        self.break_gui()
        self.save_data_gui()


    def file_setup(self):
        self.APPNAME = "Timer App"
        FILENAME = "Timer Data.xlsx"

        self.local_folder = os.path.expandvars(rf"%APPDATA%\{self.APPNAME}")
        self.data_file = os.path.expandvars(rf"%APPDATA%\{self.APPNAME}\{FILENAME}")

        os.makedirs(self.local_folder, exist_ok=True)

        if os.path.isfile(self.data_file):
            self.workbook = op.load_workbook(self.data_file)
            self.worksheet = self.workbook.active

            #collect_data()
            print("File loaded")

        else:
            self.workbook = op.Workbook()
            self.worksheet = self.workbook.active

            self.workbook.save(self.data_file)
            print("New file created")
            #customize_excel(worksheet)


    def window_setup(self):
        self.WINDOW = ctk.CTk()
        self.WINDOW.geometry(str(WIDTH + BORDER_WIDTH + main_frame_pad_x + tab_frame_width) + "x" + str(HEIGHT+((widget_padding_x+frame_padding)*2)))
        self.WINDOW.title(self.APPNAME)
        self.WINDOW.configure(background=window_color)
        self.WINDOW.resizable(False, False)
        self.WINDOW.grid_propagate(False)


    def main_frame_gui_setup(self):
        self.main_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.main_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.main_frame.grid_propagate(False)

        self.statistics_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.statistics_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.statistics_frame.grid_forget()

        self.settings_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.settings_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.settings_frame.grid_forget()

        self.achievements_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.achievements_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.achievements_frame.grid_forget()

        self.history_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.history_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.history_frame.grid_forget()


    def timer_frame_gui_setup(self):
        self.timer_break_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width)
        self.timer_break_frame.grid(row=0, column=1)
        self.timer_break_frame.pack_propagate(False)


    def timer_gui(self):
        timer_frame = ctk.CTkFrame(self.timer_break_frame, fg_color=frame_color, corner_radius=10, width=frame_width, height=220)
        timer_frame.pack(padx=frame_padding, pady=frame_padding)
        timer_frame.pack_propagate(False)

        timer_label = ctk.CTkLabel(timer_frame, text="Timer", font=(font_family, font_size), text_color=font_color)
        timer_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.time_display_label = ctk.CTkLabel(timer_frame, text="0:00:00", font=(font_family, int(font_size*3)), text_color=font_color)
        self.time_display_label.place(anchor="center", relx=0.5, rely=0.45)
        self.timer_button = ctk.CTkButton(timer_frame, text="Start", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=self.timer_mechanism)
        self.timer_button.place(anchor="s", relx=0.5, rely=0.9)
        
    
    def break_gui(self):
        break_frame = ctk.CTkFrame(self.timer_break_frame, fg_color=frame_color, corner_radius=10, width=frame_width, height=220)
        break_frame.pack(padx=frame_padding, pady=frame_padding)
        break_frame.pack_propagate(False)

        break_label = ctk.CTkLabel(break_frame, text="Break", font=(font_family, font_size), text_color=font_color)
        break_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.break_display_label = ctk.CTkLabel(break_frame, text="0:00:00", font=(font_family, int(font_size*3)), text_color=font_color)
        self.break_display_label.place(anchor="center", relx=0.5, rely=0.45)
        self.break_button = ctk.CTkButton(break_frame, text="Start", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=self.break_mechanism)
        self.break_button.place(anchor="s", relx=0.5, rely=0.9)

    def save_data_gui(self):
        self.data_frame = ctk.CTkFrame(self.main_frame, fg_color=frame_color, corner_radius=10, width=WIDTH-10, height=button_height*2)
        self.data_frame.place(anchor="s", relx=0.5, rely=0.985)
        self.data_frame.grid_propagate(False)
        save_data_btn = ctk.CTkButton(self.data_frame, text="Save Data", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                    border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, width=450, command=self.save_data)
        save_data_btn.place(relx=0.5, anchor="center", rely=0.5)


    def timer_mechanism(self):
        self.timer_manager.timer_mechanism(self.timer_button, self.break_button, self.time_display_label)

    
    def break_mechanism(self):
        self.timer_manager.break_mechanism(self.break_button, self.timer_button, self.break_display_label)


    def save_data(self):
        self.data_manager.save_data(self.timer_button, self.break_button, self.time_display_label, self.break_display_label)


    def run(self):
        self.WINDOW.mainloop()


if __name__ == "__main__":
    App().run()
