import os
import random
import time
from threading import Thread

import openpyxl as op
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
        self.APPNAME = "Timer App"
        self.FILENAME = "Timer Data.xlsx"

        self._window_setup()
        self.initialize_variables()
        self.create_gui()
        self._file_setup()

        self.t1 = Thread(target=self.eye_protection)
        self.data_manager.load_subject()
        self.data_manager.load_eye_care()

        self.WINDOW.protocol("WM_DELETE_WINDOW", self.save_on_quit)


    def create_gui(self):
        self._main_frame_gui_setup()
        self._tab_frames_gui_setup()
        self._secondary_frames_gui_setup()

        self._timer_gui_setup()
        self._break_gui_setup()
        self._save_data_gui()
        self._goal_gui_setup()
        self._progress_gui_setup()
        self._streak_gui_setup()
        self._subject_gui_setup()
        self._pomodoro_gui_setup()
        self._history_gui_setup()
        self._notes_gui_setup()
        self._settings_gui_setup()


    def _file_setup(self) -> None:
        self.local_folder = os.path.expandvars(rf"%APPDATA%\{self.APPNAME}")
        self.data_file = os.path.expandvars(rf"%APPDATA%\{self.APPNAME}\{self.FILENAME}")

        os.makedirs(self.local_folder, exist_ok=True)

        self.timer_manager = TimerManager(self, self.WINDOW)

        data_file_exists = os.path.isfile(self.data_file)

        if data_file_exists:
            print("File loaded.")
            self.workbook = op.load_workbook(self.data_file)
            self.worksheet = self.workbook.active

            self.data_manager = DataManager(self, self.timer_manager, self.workbook, self.worksheet)

            self.create_widget_list()
            self.collect_data()
            self.update_streak_values()
            self.load_history()

            self.data_manager.load_color()
            self.data_manager.load_theme()
            self.data_manager.load_notes()

        else:
            print("New file created.")
            self.workbook = op.Workbook()
            self.worksheet = self.workbook.active

            self.workbook.save(self.data_file)

            self.data_manager = DataManager(self, self.timer_manager, self.workbook, self.worksheet)

            self.create_widget_list()

            self.data_manager.initialize_new_file_variables()

            self.data_manager.save_eye_care("Off", "Off")


    def initialize_variables(self) -> None:
        self.widget_list = []
        self.default_choice = ctk.StringVar(value="1 hour")
        self.notification_limit_on = False
        self.goal = 60


    def _window_setup(self) -> None:
        self.WINDOW = ctk.CTk()
        self.WINDOW.geometry(str(WIDTH + main_frame_pad_x + tab_frame_width) + "x" + str(HEIGHT+((widget_padding_x+frame_padding)*2)))
        self.WINDOW.title(self.APPNAME)
        self.WINDOW.configure(fg_color=(light_window_color, window_color))
        self.WINDOW.resizable(False, False)
        self.WINDOW.grid_propagate(False)


    def _main_frame_gui_setup(self) -> None:
        self.tab_frame = ctk.CTkFrame(self.WINDOW, width=tab_frame_width, height=HEIGHT+((widget_padding_x+frame_padding)*2), fg_color=(light_tab_frame_color, tab_frame_color))
        self.tab_frame.grid(column=0, row=0)
        self.tab_frame.pack_propagate(False)

        self.main_frame = ctk.CTkFrame(self.WINDOW, fg_color=(light_main_frame_color, main_frame_color), height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.main_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.main_frame.grid_propagate(False)

        self.statistics_frame = ctk.CTkFrame(self.WINDOW, fg_color=(light_main_frame_color, main_frame_color), height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.statistics_frame.grid(column=2, row=0, padx=main_frame_pad_x)

        self.settings_frame = ctk.CTkFrame(self.WINDOW, fg_color=(light_main_frame_color, main_frame_color), height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.settings_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.settings_frame.grid_forget()

        self.achievements_frame = ctk.CTkFrame(self.WINDOW, fg_color=(light_main_frame_color, main_frame_color), height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.achievements_frame.grid(column=2, row=0, padx=main_frame_pad_x)

        self.history_frame = ctk.CTkFrame(self.WINDOW, fg_color=(light_main_frame_color, main_frame_color), height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.history_frame.grid(column=2, row=0, padx=main_frame_pad_x)

        self.notes_frame = ctk.CTkFrame(self.WINDOW, fg_color=(light_main_frame_color, main_frame_color), height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.notes_frame.grid(column=2, row=0, padx=main_frame_pad_x)

        self.forget_and_propagate([self.statistics_frame, self.settings_frame, self.achievements_frame, self.history_frame, self.notes_frame])


    def _tab_frames_gui_setup(self) -> None:
        self.timer_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=(light_tab_color, tab_color))
        self.timer_tab.pack(pady=tab_padding_y)
        self.timer_tab_button = ctk.CTkButton(self.timer_tab, text="Timer", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=(light_font_color, font_color),
                                              fg_color=(light_tab_selected_color, tab_selected_color), width=int(tab_frame_width*0.95), height=int(tab_height*0.8), hover_color=(light_tab_highlight_color, tab_highlight_color), 
                                              anchor="w", command=lambda: self.switch_tab("main"))
        self.timer_tab_button.place(relx=0.5, rely=0.5, anchor="center")

        self.statistics_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=(light_tab_color, tab_color))
        self.statistics_tab.pack(pady=tab_padding_y)
        self.statistics_tab_button = ctk.CTkButton(self.statistics_tab, text="Statistics", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=(light_font_color, font_color),
                                                   fg_color=(light_tab_color, tab_color), width=int(tab_frame_width*0.95), height=int(tab_height*0.8), hover_color=(light_tab_highlight_color, tab_highlight_color), 
                                                   anchor="w", command=lambda: self.switch_tab("statistics"))
        self.statistics_tab_button.place(relx=0.5, rely=0.5, anchor="center")

        self.achievements_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=(light_tab_color, tab_color))
        self.achievements_tab.pack(pady=tab_padding_y)
        self.achievements_tab_button = ctk.CTkButton(self.achievements_tab, text="Achievements", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=(light_font_color, font_color),
                                                     fg_color=(light_tab_color, tab_color), width=int(tab_frame_width*0.95), height=int(tab_height*0.8), hover_color=(light_tab_highlight_color, tab_highlight_color), 
                                                     anchor="w", command=lambda: self.switch_tab("achievements"))
        self.achievements_tab_button.place(relx=0.5, rely=0.5, anchor="center")

        self.history_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=(light_tab_color, tab_color))
        self.history_tab.pack(pady=tab_padding_y)
        self.history_tab_button = ctk.CTkButton(self.history_tab, text="History", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=(light_font_color, font_color),
                                                fg_color=(light_tab_color, tab_color), width=int(tab_frame_width*0.95), height=int(tab_height*0.8), hover_color=(light_tab_highlight_color, tab_highlight_color), 
                                                anchor="w", command=lambda: self.switch_tab("history"))
        self.history_tab_button.place(relx=0.5, rely=0.5, anchor="center")

        self.notes_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=(light_tab_color, tab_color))
        self.notes_tab.pack(pady=tab_padding_y)
        self.notes_tab_button = ctk.CTkButton(self.notes_tab, text="Notes", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=(light_font_color, font_color),
                                                fg_color=(light_tab_color, tab_color), width=int(tab_frame_width*0.95), height=int(tab_height*0.8), hover_color=(light_tab_highlight_color, tab_highlight_color), 
                                                anchor="w", command=lambda: self.switch_tab("notes"))
        self.notes_tab_button.place(relx=0.5, rely=0.5, anchor="center")

        self.settings_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=(light_tab_color, tab_color))
        self.settings_tab.place(relx=0.5, rely=1, anchor="s")
        self.settings_tab_button = ctk.CTkButton(self.settings_tab, text="Settings", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=(light_font_color, font_color),
                                 fg_color=(light_tab_color, tab_color), width=int(tab_frame_width*0.95), height=int(tab_height*0.8), hover_color=(light_tab_highlight_color, tab_highlight_color), anchor="w", command=lambda: self.switch_tab("settings"))
        self.settings_tab_button.place(relx=0.5, rely=0.5, anchor="center")


    def _secondary_frames_gui_setup(self) -> None:
        self.timer_break_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width, fg_color="transparent")
        self.timer_break_frame.grid(row=0, column=1)
        self.timer_break_frame.pack_propagate(False)

        self.goal_progress_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width, fg_color="transparent")
        self.goal_progress_frame.grid(row=0, column=0)
        self.goal_progress_frame.pack_propagate(False)

        self.subject_pomodoro_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width, fg_color="transparent")
        self.subject_pomodoro_frame.grid(row=0, column=2)
        self.subject_pomodoro_frame.pack_propagate(False)


    def _goal_gui_setup(self) -> None:
        goal_frame = ctk.CTkFrame(self.goal_progress_frame, fg_color=(light_frame_color, frame_color), height=175, width=frame_width, corner_radius=10)
        goal_frame.pack(padx=frame_padding, pady=frame_padding)
        goal_frame.pack_propagate(False)

        goal_label = ctk.CTkLabel(goal_frame, text="Goal", font=(font_family, font_size), text_color=(light_font_color, font_color))
        goal_label.place(anchor="nw", relx=0.05, rely=0.05)

        self.goal_dropdown = ctk.CTkComboBox(goal_frame, values=["1 minutes", "30 minutes", "1 hour", "1 hour, 30 minutes", "2 hours", "2 hours, 30 minutes", "3 hours", "3 hours, 30 minutes",
                                                            "4 hours", "4 hours, 30 minutes", "5 hours", "5 hours, 30 minutes", "6 hours"], variable=self.default_choice, 
                                                            state="readonly", width=200, height=30, dropdown_font=(font_family, int(font_size*0.75)), border_color=(light_border_frame_color, border_frame_color),
                                                            font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color))
        self.goal_dropdown.place(anchor="center", relx=0.5, rely=0.45)

        self.goal_button = ctk.CTkButton(goal_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                height=button_height, command=self.set_goal)
        self.goal_button.place(anchor="s", relx=0.5, rely=0.9)


    def _progress_gui_setup(self) -> None:
        progress_frame = ctk.CTkFrame(self.goal_progress_frame, fg_color=(light_frame_color, frame_color), width=frame_width, corner_radius=10, height=100)
        progress_frame.pack(padx=frame_padding, pady=frame_padding)
        progress_frame.pack_propagate(False)

        progress_label = ctk.CTkLabel(progress_frame, text="Progress", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        progress_label.place(anchor="nw", relx=0.05, rely=0.05)

        self.progressbar = ctk.CTkProgressBar(progress_frame, height=20, width=220, progress_color=button_color, fg_color=(light_border_frame_color, border_frame_color), corner_radius=10)
        self.progressbar.place(anchor="center", relx=0.5, rely=0.65)
        self.progressbar.set(0)


    def _streak_gui_setup(self) -> None:
        streak_frame = ctk.CTkFrame(self.goal_progress_frame, fg_color=(light_frame_color, frame_color), width=frame_width, corner_radius=10, height=220)
        streak_frame.pack(padx=frame_padding, pady=frame_padding)

        streak_label = ctk.CTkLabel(streak_frame, text="Streak", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        streak_label.place(anchor="nw", relx=0.05, rely=0.05)
        
        times_studied_text = ctk.CTkLabel(streak_frame, text="Goal\nreached", font=(font_family, int(font_size/1.25)), text_color=(light_font_color, font_color))
        times_studied_text.place(anchor="center", relx=0.3, rely=0.4)
        self.times_goal_reached = ctk.CTkLabel(streak_frame, text=0, font=(font_family, int(font_size*2.7)), text_color=(light_font_color, font_color))
        self.times_goal_reached.place(anchor="center", relx=0.3, rely=0.6)
        times_reached_label = ctk.CTkLabel(streak_frame, text="times", font=(font_family, int(font_size/1.25)), text_color=(light_font_color, font_color))
        times_reached_label.place(anchor="center", relx=0.3, rely=0.8)

        duration_studied_text = ctk.CTkLabel(streak_frame, text="Time\nstudied", font=(font_family, int(font_size/1.25)), text_color=(light_font_color, font_color))
        duration_studied_text.place(anchor="center", relx=0.7, rely=0.4)
        self.streak_duration = ctk.CTkLabel(streak_frame, text=0, font=(font_family, int(font_size*2.7)), text_color=(light_font_color, font_color))
        self.streak_duration.place(anchor="center", relx=0.7, rely=0.6)
        duration_minute_label = ctk.CTkLabel(streak_frame, text="minutes", font=(font_family, int(font_size/1.25)), text_color=(light_font_color, font_color))
        duration_minute_label.place(anchor="center", relx=0.7, rely=0.8)


    def _timer_gui_setup(self) -> None:
        timer_frame = ctk.CTkFrame(self.timer_break_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=frame_width, height=220)
        timer_frame.pack(padx=frame_padding, pady=frame_padding)
        timer_frame.pack_propagate(False)

        timer_label = ctk.CTkLabel(timer_frame, text="Timer", font=(font_family, font_size), text_color=(light_font_color, font_color))
        timer_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.time_display_label = ctk.CTkLabel(timer_frame, text="0:00:00", font=(font_family, int(font_size*3)), text_color=(light_font_color, font_color))
        self.time_display_label.place(anchor="center", relx=0.5, rely=0.45)
        self.timer_button = ctk.CTkButton(timer_frame, text="Start", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=self.timer_mechanism)
        self.timer_button.place(anchor="s", relx=0.5, rely=0.9)
        
    
    def _break_gui_setup(self) -> None:
        break_frame = ctk.CTkFrame(self.timer_break_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=frame_width, height=220)
        break_frame.pack(padx=frame_padding, pady=frame_padding)
        break_frame.pack_propagate(False)

        break_label = ctk.CTkLabel(break_frame, text="Break", font=(font_family, font_size), text_color=(light_font_color, font_color))
        break_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.break_display_label = ctk.CTkLabel(break_frame, text="0:00:00", font=(font_family, int(font_size*3)), text_color=(light_font_color, font_color))
        self.break_display_label.place(anchor="center", relx=0.5, rely=0.45)
        self.break_button = ctk.CTkButton(break_frame, text="Start", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=self.break_mechanism)
        self.break_button.place(anchor="s", relx=0.5, rely=0.9)

    def _save_data_gui(self) -> None:
        self.data_frame = ctk.CTkFrame(self.main_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=WIDTH-10, height=button_height*2)
        self.data_frame.place(anchor="s", relx=0.5, rely=0.985)
        self.data_frame.grid_propagate(False)
        self.save_data_button = ctk.CTkButton(self.data_frame, text="Save Data", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                    border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, width=450, command=self.save_data)
        self.save_data_button.place(relx=0.5, anchor="center", rely=0.5)


    def _history_gui_setup(self) -> None:
        history_frame_frame = ctk.CTkFrame(self.history_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, height=(HEIGHT+((widget_padding_x)*2)), width=WIDTH-frame_padding*2)
        history_frame_frame.grid(row=0, column=0, padx=frame_padding, pady=(frame_padding, 0))
        history_frame_frame.pack_propagate(False)

        history_label_frame = ctk.CTkFrame(history_frame_frame, fg_color=(light_frame_color, frame_color), width=WIDTH-(frame_padding*4), height=35)
        history_label_frame.pack(pady=(frame_padding, 0))
        history_label_frame.grid_propagate(False)

        history_data_frame = ctk.CTkScrollableFrame(history_frame_frame, fg_color="transparent", width=WIDTH-(frame_padding*4), height=520+frame_padding*2)
        history_data_frame.pack(padx=frame_padding)

        start_label = ctk.CTkLabel(history_label_frame, text="Start", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        start_label.grid(row=0, column=0)
        end_label = ctk.CTkLabel(history_label_frame, text="End", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        end_label.grid(row=0, column=1)
        duration_label = ctk.CTkLabel(history_label_frame, text="Duration", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        duration_label.grid(row=0, column=2)
        break_label = ctk.CTkLabel(history_label_frame, text="Break", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        break_label.grid(row=0, column=3)
        subject_label = ctk.CTkLabel(history_label_frame, text="Subject", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        subject_label.grid(row=0, column=4)

        start_frame = ctk.CTkFrame(history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        start_frame.grid(row=1, column=0)
        start_frame.pack_propagate(False)
        self.start_text = ctk.CTkLabel(start_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.start_text.pack()

        end_frame = ctk.CTkFrame(history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        end_frame.grid(row=1, column=1)
        end_frame.pack_propagate(False)
        self.end_text = ctk.CTkLabel(end_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.end_text.pack()

        duration_frame = ctk.CTkFrame(history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        duration_frame.grid(row=1, column=2)
        duration_frame.pack_propagate(False)
        self.duration_text = ctk.CTkLabel(duration_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.duration_text.pack()

        break_frame = ctk.CTkFrame(history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        break_frame.grid(row=1, column=3)
        break_frame.pack_propagate(False)
        self.break_text = ctk.CTkLabel(break_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.break_text.pack()

        subject_frame = ctk.CTkFrame(history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        subject_frame.grid(row=1, column=4)
        subject_frame.pack_propagate(False)
        self.subject_text = ctk.CTkLabel(subject_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.subject_text.pack()


    def _settings_gui_setup(self) -> None:
        self.color_frame = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        self.color_frame.grid(column=0)

        self._color_gui_setup()
        self._eye_care_gui_setup()

        reset_frame = ctk.CTkFrame(self.settings_frame, fg_color=(light_tab_color, tab_color))
        reset_frame.place(anchor="s", relx=0.5, rely=0.985)
        self.reset_data_button = ctk.CTkButton(reset_frame, text="Reset Data", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=self.reset_data, width=450)
        self.reset_data_button.pack()

    
    def _color_gui_setup(self) -> None:
        color_select_frame = ctk.CTkFrame(self.color_frame, fg_color=(light_frame_color, frame_color), height=200, width=int(frame_width/1.25), corner_radius=10)
        color_select_frame.grid(column=0, row=0, padx=frame_padding, pady=frame_padding)
        color_label = ctk.CTkLabel(color_select_frame, text="Color", font=(font_family, font_size), text_color=(light_font_color, font_color))
        color_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.color_dropdown = ctk.CTkComboBox(color_select_frame, values=["Orange", "Green", "Blue", "Pink"], state="readonly", width=150, height=30, border_color=(light_border_frame_color, border_frame_color),
                                         dropdown_font=(font_family, int(font_size*0.75)), font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color))
        self.color_dropdown.place(anchor="center", relx=0.5, rely=0.45)
        self.color_button = ctk.CTkButton(color_select_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                     height=button_height, command=lambda: self.data_manager.set_color(self.color_dropdown))
        self.color_button.place(anchor="s", relx=0.5, rely=0.9)

        theme_select_frame = ctk.CTkFrame(self.color_frame, fg_color=(light_frame_color, frame_color), height=200, width=int(frame_width/1.25), corner_radius=10)
        theme_select_frame.grid(row=1, column=0, padx=frame_padding, pady=frame_padding)
        theme_label = ctk.CTkLabel(theme_select_frame, text="Theme", font=(font_family, font_size), text_color=(light_font_color, font_color))
        theme_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.theme_dropdown = ctk.CTkComboBox(theme_select_frame, values=["Dark", "Light"], state="readonly", width=150, height=30, border_color=(light_border_frame_color, border_frame_color),
                                         dropdown_font=(font_family, int(font_size*0.75)), font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color))
        self.theme_dropdown.place(anchor="center", relx=0.5, rely=0.45)
        self.theme_button = ctk.CTkButton(theme_select_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                     height=button_height, command=lambda: self.data_manager.set_theme(self.theme_dropdown))
        self.theme_button.place(anchor="s", relx=0.5, rely=0.9)

    def _eye_care_gui_setup(self):
        eye_care_frame = ctk.CTkFrame(self.settings_frame, fg_color=(light_frame_color, frame_color), height=250, width=int(frame_width/1.25), corner_radius=10)
        eye_care_frame.grid(column=1, row=0, padx=frame_padding, pady=frame_padding)
        eye_care_label = ctk.CTkLabel(eye_care_frame, text="Eye care", font=(font_family, font_size), text_color=(light_font_color, font_color))
        eye_care_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.eye_care_selection = ctk.CTkComboBox(eye_care_frame, values=["On", "Off"], state="readonly", width=100, height=30, dropdown_font=(font_family, int(font_size*0.75)), 
                                             font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color), border_color=(light_border_frame_color, border_frame_color))
        self.eye_care_selection.place(anchor="center", relx=0.5, rely=0.35)
        self.eye_care_checkbox = ctk.CTkCheckBox(eye_care_frame, text="Only on when timer running", fg_color=button_color, hover=False, offvalue="Off",
                                                 font=(font_family, font_size*0.8), text_color=(light_font_color, font_color), checkmark_color=button_font_color, onvalue="On", border_color=(light_frame_border_color, border_frame_color))
        self.eye_care_checkbox.place(anchor="center", relx=0.5, rely=0.55)
        eye_care_button = ctk.CTkButton(eye_care_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, 
                                        hover_color=button_highlight_color, height=button_height, command=self.select_eye_care)
        eye_care_button.place(anchor="s", relx=0.5, rely=0.9)


    def _subject_gui_setup(self) -> None:
        subject_frame = ctk.CTkFrame(self.subject_pomodoro_frame, fg_color=(light_frame_color, frame_color), height=175, width=frame_width, corner_radius=10)
        subject_frame.pack(padx=frame_padding, pady=frame_padding)
        subject_label = ctk.CTkLabel(subject_frame, text="Subject", font=(font_family, font_size), text_color=(light_font_color, font_color))
        subject_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.subject_selection = ctk.CTkComboBox(subject_frame, values=["Mathematics", "Science", "Literature", "History", "Geography", "Language Arts", "Foreign Languages", "Social Studies",
                                                                "Economics", "Computer Science", "Psychology", "Philosophy", "Art", "Music", "Physical Education", "Other"], 
                                                            state="readonly", width=200, height=30, dropdown_font=(font_family, int(font_size*0.75)), border_color=(light_border_frame_color, border_frame_color),
                                                            font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color))
        self.subject_selection.place(anchor="center", relx=0.5, rely=0.45)
        subject_button = ctk.CTkButton(subject_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                height=button_height, command=self.select_subject)
        subject_button.place(anchor="s", relx=0.5, rely=0.9)


    def _pomodoro_gui_setup(self) -> None:
        pomodoro_frame = ctk.CTkFrame(self.subject_pomodoro_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=frame_width, height=220)
        pomodoro_frame.pack(padx=frame_padding, pady=frame_padding)
        pomodoro_label = ctk.CTkLabel(pomodoro_frame, text="Pomodoro", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        pomodoro_label.place(anchor="nw", relx=0.05, rely=0.05)
        pomodoro_button = ctk.CTkButton(pomodoro_frame, text="Start", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height)
        pomodoro_button.place(anchor="s", relx=0.5, rely=0.9)


    def _notes_gui_setup(self) -> None:
        self.notes_frame_frame = ctk.CTkFrame(self.notes_frame, fg_color="transparent", corner_radius=10, height=(HEIGHT+((widget_padding_x)*2)), width=WIDTH-frame_padding*2)
        self.notes_frame_frame.grid(row=0, column=0, padx=frame_padding, pady=(frame_padding, 0))
        self.notes_frame_frame.pack_propagate(False)

        new_note_frame = ctk.CTkFrame(self.notes_frame_frame, fg_color=(light_frame_color, frame_color), width=WIDTH-(frame_padding*4), height=35)
        new_note_frame.pack(pady=(frame_padding, 0))
        new_note_frame.grid_propagate(False)

        self.notes_data_frame = ctk.CTkScrollableFrame(self.notes_frame_frame, fg_color="transparent", width=WIDTH-(frame_padding*4), height=520+frame_padding*2)
        self.notes_data_frame.pack(padx=frame_padding)
        
        new_note_button = ctk.CTkButton(new_note_frame, text="New note", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                height=button_height, command=self._create_new_note_gui)
        new_note_button.grid()


    def _new_note_gui_setup(self):
        self.note_creation_frame = ctk.CTkFrame(self.notes_frame, fg_color="transparent", corner_radius=10, height=HEIGHT + frame_padding * 2, width=WIDTH - frame_padding * 2)
        self.note_creation_frame.grid(padx=frame_padding, pady=frame_padding)
        self.note_creation_frame.grid_propagate(False)

        self.note_title_frame = ctk.CTkFrame(self.note_creation_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=WIDTH - frame_padding * 2, height=60)
        self.note_title_frame.grid(row=0, column=0, pady=(0, frame_padding))
        self.note_title_frame.grid_propagate(False)

        self.notes_title_entry = ctk.CTkEntry(self.note_title_frame, placeholder_text="Title", font=(font_family, font_size), text_color=(light_font_color, font_color),
                                              border_color=frame_border_color, height=40, width=WIDTH - 280 - frame_padding * 6, fg_color=(light_frame_color, frame_color))
        self.notes_title_entry.grid(row=0, column=0, padx=widget_padding_x, pady=widget_padding_y)
        self.create_note_button = ctk.CTkButton(self.note_title_frame, height=button_height, text="Create note", fg_color=self.data_manager.color, 
                                                hover_color=self.data_manager.highlight_color, font=(font_family, font_size), text_color=button_font_color, command=self.create_new_note)
        self.create_note_button.grid(row=0, column=1)
        exit_create_note_button = ctk.CTkButton(self.note_title_frame, height=button_height, text="Cancel", fg_color=self.data_manager.color, 
                                                hover_color=self.data_manager.highlight_color, font=(font_family, font_size), text_color=button_font_color, command=self.exit_note_creation)
        exit_create_note_button.grid(row=0, column=2, padx=widget_padding_x, pady=widget_padding_y)
        


        notes_text_frame = ctk.CTkFrame(self.note_creation_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=WIDTH - frame_padding * 2, height=HEIGHT - 40 - frame_padding * 2)
        notes_text_frame.grid(row=1, column=0, pady=(frame_padding, 0))
        self.notes_textbox = ctk.CTkTextbox(notes_text_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), border_color=frame_border_color, 
                                            width=WIDTH - frame_padding * 4, height=HEIGHT - frame_padding * 8, fg_color=(light_frame_color, frame_color), border_width=2, bg_color="black")
        self.notes_textbox.grid(row=1, column=0, padx=widget_padding_x, pady=widget_padding_y)


    def create_time_spent_graph(self) -> None:
        data = {"Date": self.data_manager.date_list, "Duration": self.data_manager.duration_list}
        df = pd.DataFrame(data)
        grouped_data = df.groupby("Date")["Duration"].sum().reset_index()
        fig1, ax = plt.subplots()
        ax.bar(grouped_data["Date"], grouped_data["Duration"], color=self.data_manager.graph_color)
        ax.set_title("Duration of Study Sessions by Date", color=font_color)
        ax.tick_params(colors="white")
        ax.set_facecolor(graph_fg_color)
        fig1.set_facecolor(graph_bg_color)
        ax.spines["top"].set_color(spine_color)
        ax.spines["bottom"].set_color(spine_color)
        ax.spines["left"].set_color(spine_color)
        ax.spines["right"].set_color(spine_color)
        fig1.set_size_inches(graph_width/100, graph_height/100, forward=True)
        ax.tick_params(axis='x', labelrotation = 45)

        def _format_func(value, tick_number):
            return f"{int(value)} m"
        
        plt.gca().yaxis.set_major_formatter(FuncFormatter(_format_func))
        date_format = mdates.DateFormatter("%d/%m")
        ax.xaxis.set_major_formatter(date_format)
        ax.xaxis.set_major_locator(MaxNLocator(integer=True, prune='both'))
        time_spent_frame = FigureCanvasTkAgg(fig1, master=self.statistics_frame)
        plt.subplots_adjust(bottom=0.2)

        time_spent_graph = time_spent_frame.get_tk_widget()
        time_spent_graph.grid(row=0, column=0, padx=10, pady=10)
        time_spent_graph.config(highlightbackground=frame_border_color, highlightthickness=2, background=frame_color)


    def create_weekday_graph(self) -> None:
        self.data_manager.collect_day_data()

        if self.data_manager.day_duration_list:
            non_zero_durations = [duration for duration in self.data_manager.day_duration_list if duration != 0]
            non_zero_names = [name for name, duration in zip(self.data_manager.day_name_list, self.data_manager.day_duration_list) if duration != 0]

        else:
            non_zero_durations = [0]
            non_zero_names = []

        def _autopct_format(values):
            def _my_format(pct):
                total = sum(values)
                val = int(round(pct*total/100.0))
                return "{v:d} m".format(v=val)
            return _my_format

        fig, ax = plt.subplots()
        ax.pie(non_zero_durations, labels=non_zero_names, autopct=_autopct_format(non_zero_durations), colors=self.data_manager.pie_colors, 
               textprops={"fontsize": pie_font_size, "family": pie_font_family, "color": font_color}, counterclock=False, startangle=90)
        fig.set_size_inches(graph_width/100, graph_height/100, forward=True)
        fig.set_facecolor(graph_bg_color)
        ax.tick_params(colors="white")
        ax.set_facecolor(graph_fg_color)
        ax.set_title("Duration of Study Sessions by Day of the Week", color=font_color)
        ax.spines["top"].set_color(spine_color)
        ax.spines["bottom"].set_color(spine_color)
        ax.spines["left"].set_color(spine_color)
        ax.spines["right"].set_color(spine_color)

        weekday_frame = FigureCanvasTkAgg(fig, master=self.statistics_frame)

        weekday_graph = weekday_frame.get_tk_widget()
        weekday_graph.grid(row=0, column=1, padx=10, pady=10)
        weekday_graph.config(highlightbackground=frame_border_color, highlightthickness=2, background=frame_color)


    def create_graphs(self) -> None:
        self.statistics_frame.grid_propagate(False)
        self.create_time_spent_graph()
        self.create_weekday_graph()

    
    def forget_and_propagate(self, list: list) -> None:
        for item in list:
            item.grid_forget()
            item.grid_propagate(False)


    def switch_tab(self, tab = str) -> None:
        tabs = {
            "main": [self.main_frame, self.timer_tab_button],
            "statistics": [self.statistics_frame, self.statistics_tab_button],
            "settings": [self.settings_frame, self.settings_tab_button],
            "achievements": [self.achievements_frame, self.achievements_tab_button],
            "history": [self.history_frame, self.history_tab_button],
            "notes": [self.notes_frame, self.notes_tab_button]
        }

        tab_list = [self.main_frame, self.statistics_frame, self.settings_frame, self.achievements_frame, self.history_frame, self.notes_frame]
        tab_list.remove(tabs[tab][0])

        def _forget_tabs():
            for frame in tab_list:
                frame.grid_forget()

        _forget_tabs()

        tabs[tab][0].grid(column=2, row=0, padx=main_frame_pad_x)
        tabs[tab][0].grid_propagate(False)

        tab_button_list = [self.timer_tab_button, self.statistics_tab_button, self.settings_tab_button, self.achievements_tab_button, self.history_tab_button, self.notes_tab_button]
        tab_button_list.remove(tabs[tab][1])

        def _decolor_tabs():
            for button in tab_button_list:
                button.configure(fg_color=(light_tab_color, tab_color), hover_color=(light_tab_highlight_color, tab_highlight_color))

        _decolor_tabs()

        tabs[tab][1].configure(fg_color=(light_tab_selected_color, tab_selected_color), hover_color=(light_tab_selected_color, tab_selected_color))



    def timer_mechanism(self) -> None:
        self.timer_manager.timer_mechanism(self.timer_button, self.break_button, self.time_display_label)
        
        #Get start time only at the start of timer
        if self.timer_manager.timer_time <= 1:
            self.data_manager.get_start_time()


    def reset_gui_values(self) -> None:
        self.times_goal_reached.configure(text=0)
        self.streak_duration.configure(text=0)
        self.progressbar.set(0)
        self.progressbar.configure(progress_color = self.data_manager.color)
        self.reset_timers()


    def reset_timers(self) -> None:
        self.timer_button.configure(text="Start")
        self.break_button.configure(text="Start")
        self.time_display_label.configure(text="0:00:00")
        self.break_display_label.configure(text="0:00:00") 


    def set_goal(self) -> None:
        x = 0
        choice = self.goal_dropdown.get()

        #Make an int out of a string e.g. "1 hour, 30 minutes"
        if "hour" in choice:
            x += int(choice.split(" ")[0]) * 60
        if "minutes" in choice and "hour" in choice:
            x += int(choice.split(", ")[1].removesuffix(" minutes"))
        if "hour" not in choice:
            x += int(choice.split(" ")[0])
        self.goal = x


    def select_subject(self):
        subject = self.subject_selection.get()
        self.subject_selection.configure(variable=ctk.StringVar(value=subject))

        self.data_manager.save_subject(subject)


    def select_eye_care(self):
        eye_care = self.eye_care_selection.get()
        self.eye_care_selection.configure(variable=ctk.StringVar(value=eye_care))

        checkbox = self.eye_care_checkbox.get()
        self.eye_care_checkbox.configure(variable=ctk.StringVar(value=checkbox))

        self.data_manager.save_eye_care(eye_care, checkbox)


    def reach_goal(self, timer_time: int) -> None:
        time_in_minutes = timer_time / 60
        if time_in_minutes < self.goal:
            self.progressbar.set(time_in_minutes/self.goal)
        elif time_in_minutes >= self.goal and not self.notification_limit_on:
            self.progressbar.set(1)

            message = random.choice(["Congratulations! You've reached your study goal. Take a well-deserved break and recharge!", "Study session complete! Great job on reaching your goal. Time for a quick break!",
                                     "You did it! Study session accomplished. Treat yourself to a moment of relaxation!", "Well done! You've met your study goal. Now, take some time to unwind and reflect on your progress.",
                                     "Study session over! You've achieved your goal. Reward yourself with a brief pause before your next task.", "Goal achieved! Take a breather and pat yourself on the back for your hard work.",
                                     "Mission accomplished! You've hit your study target. Enjoy a short break before diving back in.", "Study session complete. Nicely done! Use this time to relax and rejuvenate before your next endeavor.",
                                     "You've reached your study goal! Treat yourself to a well-deserved break. You've earned it!", "Goal achieved! Take a moment to celebrate your success. Your dedication is paying off!"])
            self.send_notification("Study Goal Reached", message)


    def update_streak_values(self) -> None:
        self.times_goal_reached.configure(text=self.data_manager.goal_amount)
        self.streak_duration.configure(text=self.data_manager.total_duration)

    
    def break_mechanism(self) -> None:
        self.timer_manager.break_mechanism(self.break_button, self.timer_button, self.break_display_label)


    def collect_data(self) -> None:
        self.data_manager.collect_data()
        self.data_manager.data_to_variable()


    def save_data(self) -> None:
        time_in_minutes = self.timer_manager.timer_time / 60

        #Only be able to save if time is higher than 1m
        if time_in_minutes >= 1:
            if time_in_minutes >= self.goal:
                self.data_manager.increase_goal_streak()

            self.data_manager.save_data()

            self.collect_data()
            self.reset_gui_values()
            self.update_streak_values()
            self.create_time_spent_graph()
            self.create_weekday_graph()
            self.load_history()
            
            self.notification_limit_on = False
        else:
            print("No data to save. (time less than 1m)")


    def eye_protection(self):
        checkbox = self.eye_care_checkbox.get()
        time_between = 60 * 20
        if self.eye_care_selection.get() == "On":
            if checkbox == "On" and self.timer_manager.timer_running:
                time.sleep(time_between)
                if checkbox == "On" and self.timer_manager.timer_running:
                    self.send_notification("Eye Protection", "Look away 20ft for 20 seconds")
                    self.eye_protection()
            else:
                time.sleep(time_between)
                if self.eye_care_selection.get() == "On" and checkbox == "Off":
                    self.send_notification("Eye Protection", "Look away 20ft for 20 seconds")
                    self.eye_protection()
        else:
            time.sleep(10)
            self.eye_protection()


    def load_history(self):
        start_history = ""
        end_history = ""
        duration_history = ""
        break_history = ""
        subject_history = ""

        if self.data_manager.data_amount > 0:
            for data in range(self.data_manager.data_amount+1, 1, -1):
                start_history += str(self.worksheet["A" + str(data)].value)
                start_history += "\n"
                end_history += str(self.worksheet["B" + str(data)].value)
                end_history += "\n"
                duration_history += str(round(self.worksheet["C" + str(data)].value)) + "m"
                duration_history += "\n"
                break_history += str(round(self.worksheet["D" + str(data)].value)) + "m"
                break_history += "\n"
                subject_history += str(self.worksheet["E" + str(data)].value)
                subject_history += "\n"

            self.start_text.configure(text=start_history)
            self.end_text.configure(text=end_history)
            self.duration_text.configure(text=duration_history)
            self.break_text.configure(text=break_history)
            self.subject_text.configure(text=subject_history)


    def _create_new_note_gui(self):
        self.notes_frame_frame.grid_forget()

        self._new_note_gui_setup()
        self.widget_list.append(self.create_note_button)


    def exit_note_creation(self):
        self.notes_frame_frame.grid(row=0, column=0, padx=frame_padding, pady=(frame_padding, 0))
        self.note_creation_frame.grid_forget()


    def create_new_note(self):
        if len(self.notes_title_entry.get()) > 0 or len(self.notes_textbox.get("0.0", "end")) > 1:
            self.data_manager.create_new_note(self.notes_title_entry.get(), self.notes_textbox.get("0.0", "end"))
            self.exit_note_creation()
            return
        
        print("Error. Note title or text can't be empty.")

    
    def send_notification(self, title, message) -> None:
        toast = Notification(app_id=self.APPNAME, title=title, msg=message)
        toast.show()
        self.notification_limit_on = True
        print("Notification " + title + " sent.")


    def reset_data(self) -> None:
        os.remove(self.data_file)
        self.timer_manager.timer_running = False
        self.timer_manager.break_running = False
        self.reset_gui_values()

        self._file_setup()


    #Get all buttons other than tab buttons
    def create_widget_list(self) -> None:
        frame_list = []
        widgets = []

        for frame in self.WINDOW.winfo_children():
            if isinstance(frame, ctk.CTkFrame):
                frame_list.append(frame)
        frame_list.pop(0)

        def _get_widgets(frame):
            for child in frame.winfo_children():
                widgets.append(child)
                if isinstance(child, ctk.CTkFrame):
                    _get_widgets(child)

        for frame in frame_list:
            _get_widgets(frame)

        for widget in widgets:
            if isinstance(widget, ctk.CTkButton):
                self.widget_list.append(widget)


    def save_on_quit(self) -> None:
        self.save_data()
        print("Data saved on exit.")

        self.workbook.save(self.data_file)
        self.WINDOW.destroy()


    def run(self) -> None:
        self.WINDOW.mainloop()


if __name__ == "__main__":
    App().run()
