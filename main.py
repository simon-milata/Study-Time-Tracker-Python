import os
import random
from PIL import Image
import sys

import openpyxl as op
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator, FuncFormatter
import customtkinter as ctk
from winotify import Notification
from CTkMessagebox import CTkMessagebox

from Package import *



class App:
    def __init__(self):
        self.APPNAME = "Timer App"
        self.FILENAME = "Timer Data.xlsx"

        self._window_setup()
        self.initialize_variables()
        self._icon_setup()
        self.create_gui()
        self._file_setup()

        self.data_manager.load_subject()
        self.data_manager.load_eye_care()
        self.data_manager.load_autobreak()

        self.WINDOW.bind_all("<Button>", self.change_focus)
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
        self._autobreak_gui_setup()
        self._statistics_gui_setup()
        self._history_gui_setup()
        self._notes_gui_setup()
        self._achievement_gui_setup()
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

        else:
            print("New file created.")
            self.workbook = op.Workbook()
            self.worksheet = self.workbook.active

            self.workbook.save(self.data_file)

            self.data_manager = DataManager(self, self.timer_manager, self.workbook, self.worksheet)

            self.create_widget_list()

            self.data_manager.initialize_new_file_variables()
            self.data_manager.load_autobreak()
            self.data_manager.save_eye_care("Off", "Off")


    def initialize_variables(self) -> None:
        self.statistics_scroll_position = "left"
        self.achievements_scroll_position = "left"
        self.scrolling = False
        self.widget_list = []
        self.default_choice = ctk.StringVar(value="1 hour")

        self.notification_limit_on = False
        self.goal = 60

    
    def _icon_setup(self):
        self.icon_attribution = "UIcons by https://www.flaticon.com/uicons Flaticon"

        self.clock_icon = ctk.CTkImage(light_image=Image.open("icons/clock_icon_dark.png"), dark_image=Image.open("icons/clock_icon_light.png"), size=(tab_icon_size, tab_icon_size))
        self.statistics_icon = ctk.CTkImage(light_image=Image.open("icons/statistics_icon_dark.png"), dark_image=Image.open("icons/statistics_icon_light.png"), size=(tab_icon_size, tab_icon_size))
        self.achievements_icon = ctk.CTkImage(light_image=Image.open("icons/achievements_icon_dark.png"), dark_image=Image.open("icons/achievements_icon_light.png"), size=(tab_icon_size, tab_icon_size))
        self.history_icon = ctk.CTkImage(light_image=Image.open("icons/history_icon_dark.png"), dark_image=Image.open("icons/history_icon_light.png"), size=(tab_icon_size, tab_icon_size))
        self.notes_icon = ctk.CTkImage(light_image=Image.open("icons/notes_icon_dark.png"), dark_image=Image.open("icons/notes_icon_light.png"), size=(tab_icon_size, tab_icon_size))
        self.settings_icon = ctk.CTkImage(light_image=Image.open("icons/settings_icon_dark.png"), dark_image=Image.open("icons/settings_icon_light.png"), size=(tab_icon_size, tab_icon_size))

        self.save_icon = ctk.CTkImage(Image.open("icons/save_icon_dark.png"), size=(icon_size, icon_size))
        self.edit_icon = ctk.CTkImage(Image.open("icons/edit_icon_dark.png"), size=(icon_size, icon_size))
        self.delete_icon = ctk.CTkImage(Image.open("icons/delete_icon_dark.png"), size=(icon_size, icon_size))
        self.open_icon = ctk.CTkImage(Image.open("icons/open_icon_dark.png"), size=(icon_size, icon_size))
        self.back_icon = ctk.CTkImage(Image.open("icons/back_icon_dark.png"), size=(icon_size, icon_size))
        self.play_icon = ctk.CTkImage(Image.open("icons/play_icon_dark.png"), size=(icon_size, icon_size))
        self.create_note_icon = ctk.CTkImage(Image.open("icons/create_note_icon_dark.png"), size=(icon_size, icon_size))
        self.checkmark_icon = ctk.CTkImage(Image.open("icons/checkmark_icon_dark.png"), size=(icon_size, icon_size))
        self.left_arrow_icon = ctk.CTkImage(dark_image=Image.open("icons/arrow_left_icon_light.png"), light_image=Image.open("icons/arrow_left_icon_dark.png"), size=(icon_size, icon_size))
        self.right_arrow_icon = ctk.CTkImage(dark_image=Image.open("icons/arrow_right_icon_light.png"), light_image=Image.open("icons/arrow_right_icon_dark.png"), size=(icon_size, icon_size))


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
        self.statistics_frame.grid(column=2, row=0)

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

        tab_buttons = []

        def _initialize_tab(tab_name: str, icon: ctk.CTkImage) -> None:
            #Create a tab frame and a tab button for each tab in tab_list
            tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=(light_tab_color, tab_color))
            tab_button = ctk.CTkButton(tab, image=icon, text=" " + tab_name, font=(tab_font_family, tab_font_size, tab_font_weight), text_color=(light_font_color, font_color),
                                                   fg_color=(light_tab_color, tab_color), width=int(tab_frame_width*0.95), height=int(tab_height*0.9), hover_color=(light_tab_highlight_color, tab_highlight_color), 
                                                   anchor="w", command=lambda: self.switch_tab(tab_button, tab_buttons))
            
            #Place settings tab frame on the bottom of tab frame
            if tab_name == "Settings":
                tab.place(relx=0.5, rely=0.995, anchor="s")
            else:
                tab.pack(pady=tab_padding_y)
            tab_button.place(relx=0.5, rely=0.5, anchor="center")

            #Append each button to tab_buttons list for future manipulation
            tab_buttons.append(tab_button)

        tab_list = ["Timer", "Statistics", "History", "Notes", "Achievements", "Settings"]
        tab_icons = [self.clock_icon, self.statistics_icon, self.history_icon, self.notes_icon, self.achievements_icon, self.settings_icon]

        for icon, tab in enumerate(tab_list):
            _initialize_tab(tab, tab_icons[icon])

        #Change the color of the timer tab button by default as it's selected
        tab_buttons[0].configure(fg_color=(light_tab_selected_color, tab_selected_color), hover=False)


    def _secondary_frames_gui_setup(self) -> None:
        self.timer_break_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width, fg_color="transparent")
        self.timer_break_frame.grid(row=0, column=1)
        self.timer_break_frame.pack_propagate(False)

        self.goal_progress_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width, fg_color="transparent")
        self.goal_progress_frame.grid(row=0, column=0)
        self.goal_progress_frame.pack_propagate(False)

        self.subject_autobreak_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width, fg_color="transparent")
        self.subject_autobreak_frame.grid(row=0, column=2)
        self.subject_autobreak_frame.pack_propagate(False)


    def _goal_gui_setup(self) -> None:
        goal_frame = ctk.CTkFrame(self.goal_progress_frame, fg_color=(light_frame_color, frame_color), height=175, width=frame_width, corner_radius=10)
        goal_frame.pack(padx=frame_padding, pady=frame_padding)
        goal_frame.pack_propagate(False)

        goal_label = ctk.CTkLabel(goal_frame, text="Goal", font=(font_family, font_size), text_color=(light_font_color, font_color))
        goal_label.place(anchor="nw", relx=0.05, rely=0.05)

        self.goal_dropdown = ctk.CTkComboBox(goal_frame, values=["30 minutes", "1 hour", "1 hour, 30 minutes", "2 hours", "2 hours, 30 minutes", "3 hours", "3 hours, 30 minutes",
                                                            "4 hours", "4 hours, 30 minutes", "5 hours", "5 hours, 30 minutes", "6 hours"], variable=self.default_choice, 
                                                            state="readonly", width=200, height=30, dropdown_font=(font_family, int(font_size*0.75)), border_color=(light_border_frame_color, border_frame_color),
                                                            font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color))
        self.goal_dropdown.place(anchor="center", relx=0.5, rely=0.45)

        self.goal_button = ctk.CTkButton(goal_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                height=button_height, command=self.set_goal, text_color_disabled=button_font_color)
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
        self.break_button = ctk.CTkButton(break_frame, text="Start", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color, text_color_disabled=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=self.break_mechanism)
        self.break_button.place(anchor="s", relx=0.5, rely=0.9)

    def _save_data_gui(self) -> None:
        self.data_frame = ctk.CTkFrame(self.main_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=WIDTH-10, height=button_height*2)
        self.data_frame.place(anchor="s", relx=0.5, rely=0.985)
        self.data_frame.grid_propagate(False)
        self.save_data_button = ctk.CTkButton(self.data_frame, text="Save Data", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                    border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, width=450, command=self.save_data)
        self.save_data_button.place(relx=0.5, anchor="center", rely=0.5)


    def _statistics_gui_setup(self):
        self.statistics_scroll_frame = ctk.CTkScrollableFrame(self.statistics_frame, fg_color="transparent", width=WIDTH, height=HEIGHT+((widget_padding_x+frame_padding)), 
                                                              orientation="horizontal")
        self.statistics_scroll_frame.grid(padx=0, pady=0)

        button_frame = ctk.CTkFrame(self.statistics_frame, fg_color="transparent", width=WIDTH)
        button_frame.place(anchor="s", relx=0.5, rely=1)

        self.left_button = ctk.CTkButton(button_frame, image=self.left_arrow_icon, text="Previous Graph", font=(font_family, font_size), fg_color=(light_frame_color, frame_color), text_color=(light_font_color, font_color),
                                         border_color=frame_border_color, hover_color=(light_tab_highlight_color, tab_highlight_color), height=30, width=WIDTH/2, command=lambda: self.scroll_statistics("left"))
        self.left_button.grid(row=0, column=0)
        self.right_button = ctk.CTkButton(button_frame, text="Next Graph", image=self.right_arrow_icon, compound="right", font=(font_family, font_size), fg_color=(light_frame_color, frame_color), text_color=(light_font_color, font_color),
                                         border_color=frame_border_color, hover_color=(light_tab_highlight_color, tab_highlight_color), height=30, width=WIDTH/2, command=lambda: self.scroll_statistics("right"))
        self.right_button.grid(row=0, column=1)

        self.statistics_graph_frame = ctk.CTkFrame(self.statistics_scroll_frame, fg_color="transparent", width=WIDTH)
        self.statistics_graph_frame.grid(row=0, column=0, padx=0, pady=0)

        self.statistics_facts_frame = ctk.CTkFrame(self.statistics_scroll_frame, fg_color="transparent", width=WIDTH)
        self.statistics_facts_frame.grid(row=1, column=0, padx=0, pady=0)


    def _graph_gui_frame(self, row_index: int, column_index: int) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self.statistics_graph_frame, fg_color=(light_frame_color, frame_color), corner_radius=10)
        frame.grid(row=row_index, column=column_index, padx=frame_padding, pady=frame_padding)
        return frame


    def _history_gui_setup(self) -> None:
        history_frame_frame = ctk.CTkFrame(self.history_frame, fg_color="transparent", corner_radius=10, height=(HEIGHT+((widget_padding_x)*4)), width=WIDTH)
        history_frame_frame.grid(row=0, column=0)
        history_frame_frame.pack_propagate(False)

        history_label_frame = ctk.CTkFrame(history_frame_frame, fg_color=(light_frame_color, frame_color), width=WIDTH - frame_padding * 2, corner_radius=10, height=50)
        history_label_frame.pack(pady=(frame_padding*1.2, frame_padding))
        history_label_frame.grid_propagate(False)

        self.history_data_frame = ctk.CTkScrollableFrame(history_frame_frame, fg_color=(light_frame_color, frame_color), width=WIDTH-(frame_padding*4), height=500+frame_padding*4, corner_radius=10)
        self.history_data_frame.pack(padx=frame_padding)

        start_label = ctk.CTkLabel(history_label_frame, text="Start", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5, height=30)
        start_label.grid(row=0, column=0, pady=widget_padding_y)
        end_label = ctk.CTkLabel(history_label_frame, text="End", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5, height=30)
        end_label.grid(row=0, column=1, pady=widget_padding_y)
        duration_label = ctk.CTkLabel(history_label_frame, text="Duration", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5, height=30)
        duration_label.grid(row=0, column=2, pady=widget_padding_y)
        break_label = ctk.CTkLabel(history_label_frame, text="Break", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5, height=30)
        break_label.grid(row=0, column=3, pady=widget_padding_y)
        subject_label = ctk.CTkLabel(history_label_frame, text="Subject", font=(font_family, int(font_size*1.25)), text_color=(light_font_color, font_color), fg_color="transparent", width=(WIDTH-(frame_padding*4))/5, height=30)
        subject_label.grid(row=0, column=4, pady=widget_padding_y)

        start_frame = ctk.CTkFrame(self.history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        start_frame.grid(row=1, column=0)
        start_frame.pack_propagate(False)
        self.start_text = ctk.CTkLabel(start_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.start_text.pack()

        end_frame = ctk.CTkFrame(self.history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        end_frame.grid(row=1, column=1)
        end_frame.pack_propagate(False)
        self.end_text = ctk.CTkLabel(end_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.end_text.pack()

        duration_frame = ctk.CTkFrame(self.history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        duration_frame.grid(row=1, column=2)
        duration_frame.pack_propagate(False)
        self.duration_text = ctk.CTkLabel(duration_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.duration_text.pack()

        break_frame = ctk.CTkFrame(self.history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        break_frame.grid(row=1, column=3)
        break_frame.pack_propagate(False)
        self.break_text = ctk.CTkLabel(break_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.break_text.pack()

        subject_frame = ctk.CTkFrame(self.history_data_frame, fg_color="transparent", width=(WIDTH-(frame_padding*4))/5)
        subject_frame.grid(row=1, column=4)
        subject_frame.pack_propagate(False)
        self.subject_text = ctk.CTkLabel(subject_frame, font=(font_family, font_size), text_color=(light_font_color, font_color), text="-")
        self.subject_text.pack()


    def _settings_gui_setup(self) -> None:
        self.color_frame = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        self.color_frame.grid(column=0, row=0)

        self.eye_care_export_frame = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        self.eye_care_export_frame.grid(column=1, row=0)

        self._color_gui_setup()
        self._eye_care_gui_setup()
        self._export_gui_setup()

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
        eye_care_frame = ctk.CTkFrame(self.eye_care_export_frame, fg_color=(light_frame_color, frame_color), height=250, width=int(frame_width/1.25), corner_radius=10)
        eye_care_frame.grid(column=0, row=0, padx=frame_padding, pady=frame_padding)
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


    def _export_gui_setup(self):
        export_frame = ctk.CTkFrame(self.eye_care_export_frame, fg_color=(light_frame_color, frame_color), height=200, width=int(frame_width/1.25), corner_radius=10)
        export_frame.grid(row=1, column=0, padx=frame_padding, pady=frame_padding)
        export_label = ctk.CTkLabel(export_frame, text="Export data", font=(font_family, font_size), text_color=(light_font_color, font_color))
        export_label.place(anchor="nw", relx=0.05, rely=0.05)
        export_file_selection = ctk.CTkComboBox(export_frame, values=["Excel"], state="readonly", width=100, height=30, dropdown_font=(font_family, int(font_size*0.75)), font=(font_family, int(font_size)), variable=ctk.StringVar(value="Excel"),
                                                fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color), border_color=(light_border_frame_color, border_frame_color))
        export_file_selection.place(anchor="center", relx=0.5, rely=0.45)
        export_button = ctk.CTkButton(export_frame, text="Export", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, 
                                        hover_color=button_highlight_color, height=button_height, command=self.export_data)
        export_button.place(anchor="s", relx=0.5, rely=0.9)

    
    def export_data(self):
        self.data_manager.export_data()


    def _subject_gui_setup(self) -> None:
        subject_frame = ctk.CTkFrame(self.subject_autobreak_frame, fg_color=(light_frame_color, frame_color), height=175, width=frame_width, corner_radius=10)
        subject_frame.pack(padx=frame_padding, pady=frame_padding)
        subject_label = ctk.CTkLabel(subject_frame, text="Subject", font=(font_family, font_size), text_color=(light_font_color, font_color))
        subject_label.place(anchor="nw", relx=0.05, rely=0.05)
        self.subject_selection = ctk.CTkComboBox(subject_frame, values=["Mathematics", "Science", "Literature", "History", "Geography", "Language Arts", "Foreign Languages", "Social Studies",
                                                                "Economics", "Computer Science", "Psychology", "Philosophy", "Art", "Music", "Physical Education", "Other"], 
                                                            state="readonly", width=200, height=30, dropdown_font=(font_family, int(font_size*0.75)), border_color=(light_border_frame_color, border_frame_color),
                                                            font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color))
        self.subject_selection.place(anchor="center", relx=0.5, rely=0.45)
        self.subject_button = ctk.CTkButton(subject_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                height=button_height, command=self.select_subject, text_color_disabled=button_font_color)
        self.subject_button.place(anchor="s", relx=0.5, rely=0.9)


    def _autobreak_gui_setup(self) -> None:
        autobreak_frame = ctk.CTkFrame(self.subject_autobreak_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=frame_width, height=400)
        autobreak_frame.pack(padx=frame_padding, pady=frame_padding)
        autobreak_frame.pack_propagate(False)
        autobreak_label = ctk.CTkLabel(autobreak_frame, text="Auto-start break", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        autobreak_label.place(anchor="nw", relx=0.05, rely=0.05)

        frequency_duration_frame = ctk.CTkFrame(autobreak_frame, fg_color="transparent")
        frequency_duration_frame.pack(pady=(frame_padding*5, frame_padding*2))

        frequency_frame = ctk.CTkFrame(frequency_duration_frame, fg_color="transparent")
        frequency_frame.grid(row=0, column=0, padx=frame_padding*2)
        frequency_label = ctk.CTkLabel(frequency_frame, text="Every", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        frequency_label.grid(row=0, column=0, pady=5)
        self.frequency_input = ctk.CTkEntry(frequency_frame, placeholder_text="30", font=(font_family, font_size * 2.7), text_color=(light_font_color, font_color),
                                              border_color=frame_border_color, height=75, width=75, fg_color=(light_frame_color, frame_color), justify="center")
        self.frequency_input.grid(row=1, column=0)
        minutes_label = ctk.CTkLabel(frequency_frame, text="minutes", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        minutes_label.grid(row=2, column=0)

        duration_frame = ctk.CTkFrame(frequency_duration_frame, fg_color="transparent")
        duration_frame.grid(row=0, column=1, padx=frame_padding*2)
        duration_label = ctk.CTkLabel(duration_frame, text="For", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        duration_label.grid(row=0, column=0, pady=5)
        self.duration_input = ctk.CTkEntry(duration_frame, placeholder_text="5", font=(font_family, font_size * 2.7), text_color=(light_font_color, font_color),
                                              border_color=frame_border_color, height=75, width=75, fg_color=(light_frame_color, frame_color), justify="center")
        self.duration_input.grid(row=1, column=0)
        duration_minutes_label = ctk.CTkLabel(duration_frame, text="minutes", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        duration_minutes_label.grid(row=2, column=0)

        self.autobreak_switch = ctk.CTkComboBox(autobreak_frame,values=["On", "Off"], state="readonly", width=100, height=30, dropdown_font=(font_family, int(font_size*0.75)), 
                                                font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color), border_color=(light_border_frame_color, border_frame_color))
        self.autobreak_switch.pack()

        self.autobreak_button = ctk.CTkButton(autobreak_frame, text="Save", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=self.save_autobreak, text_color_disabled=button_font_color)
        self.autobreak_button.place(anchor="s", relx=0.5, rely=0.925)


    def _notes_gui_setup(self) -> None:
        self.notes_frame_frame = ctk.CTkFrame(self.notes_frame, fg_color="transparent", corner_radius=10, height=(HEIGHT+((widget_padding_x)*2)), width=WIDTH-frame_padding*2)
        self.notes_frame_frame.grid(row=0, column=0, padx=frame_padding, pady=(frame_padding, 0))
        self.notes_frame_frame.pack_propagate(False)

        new_note_frame = ctk.CTkFrame(self.notes_frame_frame, fg_color=(light_frame_color, frame_color), width=WIDTH - frame_padding * 2, height=60)
        new_note_frame.pack(pady=(frame_padding, 0))
        new_note_frame.grid_propagate(False)

        self.notes_data_frame = ctk.CTkScrollableFrame(self.notes_frame_frame, fg_color="transparent", width=WIDTH + frame_padding*4, height=520+frame_padding*2, label_anchor="w")
        self.notes_data_frame.pack()
        
        new_note_button = ctk.CTkButton(new_note_frame, text="New note", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                        height=button_height, command=self._create_new_note_gui, width=450)
        new_note_button.place(anchor="center", relx=0.5, rely=0.5)


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
        self.create_note_button = ctk.CTkButton(self.note_title_frame, height=button_height, text="Done", fg_color=self.data_manager.color, 
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

    
    def _achievement_gui_setup(self):
        self.achievements_scroll_frame = ctk.CTkScrollableFrame(self.achievements_frame, fg_color="transparent", height=HEIGHT+((widget_padding_x+frame_padding)), width=WIDTH, corner_radius=0, orientation="horizontal")
        self.achievements_scroll_frame.grid(padx=0, pady=0)

        button_frame = ctk.CTkFrame(self.achievements_frame, fg_color="transparent", width=WIDTH)
        button_frame.place(anchor="s", relx=0.5, rely=1)

        self.left_button = ctk.CTkButton(button_frame, image=self.left_arrow_icon, text="Previous Page", font=(font_family, font_size), fg_color=(light_frame_color, frame_color), text_color=(light_font_color, font_color),
                                         border_color=frame_border_color, hover_color=(light_tab_highlight_color, tab_highlight_color), height=30, width=WIDTH/2, command=lambda: self.scroll_achievements("left"))
        self.left_button.grid(row=0, column=0)
        self.right_button = ctk.CTkButton(button_frame, text="Next Page", image=self.right_arrow_icon, compound="right", font=(font_family, font_size), fg_color=(light_frame_color, frame_color), text_color=(light_font_color, font_color),
                                         border_color=frame_border_color, hover_color=(light_tab_highlight_color, tab_highlight_color), height=30, width=WIDTH/2, command=lambda: self.scroll_achievements("right"))
        self.right_button.grid(row=0, column=1)


    def create_achievements(self):
        self.clear_frame(self.achievements_scroll_frame)

        row = 0
        column = -1
        index = 0
        for achievement in self.data_manager.achievements:
            column += 1
            if column % 3 == 0 and column != 0:
                column = 0
                row += 1
                if row % 2 == 0 and column % 2 == 0:
                    index += 3
                    row = 0
                    column = index
            self.create_achievement(row, column, achievement.name, achievement.title, achievement.value, achievement.max_value)


    def create_achievement(self, row, column, name, title, value, max_value):
        def split_sentence(sentence):
            midpoint = len(sentence) // 2
            space_index = sentence.rfind(' ', 0, midpoint)
            
            first_half = sentence[:space_index]
            second_half = sentence[space_index + 1:]
            
            return first_half, second_half
        
        
        if len(title) > 20:
            first_half, second_half = split_sentence(title)
            title = f"{first_half}\n{second_half}"

        if len(name) > 20:
            name = name.replace(" ", "\n")

        frame = ctk.CTkFrame(self.achievements_scroll_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=WIDTH / 3 - frame_padding * 2, height=(HEIGHT+((widget_padding_x+frame_padding)*2))/2 - frame_padding * 4)
        frame.grid(row=row, column=column, padx=frame_padding, pady=frame_padding)
        name = ctk.CTkLabel(frame, text=name, font=(font_family, int(font_size*1.9)), text_color=(light_font_color, font_color))
        name.place(anchor="n", relx=0.5, rely=0.1)
        title = ctk.CTkLabel(frame, text=title, font=(font_family, int(font_size*0.9)), text_color=(light_font_color, font_color))
        title.place(anchor="center", relx=0.5, rely=0.45)
        value_text = ctk.CTkLabel(frame, text=f"{value}/{max_value}", font=(font_family, int(font_size)), text_color=(light_font_color, font_color))
        value_text.place(anchor="center", relx=0.5, rely=0.7)
        progress_bar = ctk.CTkProgressBar(frame, height=25, width=WIDTH/5, progress_color=self.data_manager.color, fg_color=(light_border_frame_color, border_frame_color), corner_radius=25)
        progress_bar.place(anchor="center", relx=0.5, rely=0.8)
        progress_bar.set(value / max_value)
        if value == 0:
            progress_bar.configure(progress_color=(light_border_frame_color, border_frame_color))


    def create_time_spent_graph(self, frame) -> None:
        data = {"Date": self.data_manager.date_list, "Duration": self.data_manager.duration_list}
        df = pd.DataFrame(data)
        grouped_data = df.groupby("Date")["Duration"].sum().reset_index()
        fig, ax = plt.subplots()
        ax.bar(grouped_data["Date"], grouped_data["Duration"], color=self.data_manager.graph_color)
        ax.set_title("Duration of Study Sessions by Date", color=self.data_manager.font_color)
        ax.tick_params(colors=self.data_manager.font_color)
        ax.set_facecolor(self.data_manager.graph_fg_color)
        fig.set_facecolor(self.data_manager.graph_bg_color)
        ax.spines["top"].set_color(self.data_manager.spine_color)
        ax.spines["bottom"].set_color(self.data_manager.spine_color)
        ax.spines["left"].set_color(self.data_manager.spine_color)
        ax.spines["right"].set_color(self.data_manager.spine_color)
        fig.set_size_inches(graph_width/100, graph_height/100, forward=True)
        ax.tick_params(axis="x", labelrotation = 45)

        def _format_func(value, tick_number):
            return f"{int(value)} m"
        
        plt.gca().yaxis.set_major_formatter(FuncFormatter(_format_func))
        date_format = mdates.DateFormatter("%d/%m")
        ax.xaxis.set_major_formatter(date_format)
        ax.xaxis.set_major_locator(MaxNLocator(integer=True, prune="both"))
        time_spent_frame = FigureCanvasTkAgg(fig, master=frame)
        plt.subplots_adjust(bottom=0.15)

        time_spent_graph = time_spent_frame.get_tk_widget()
        time_spent_graph.pack(padx=3, pady=3)


    def create_weekday_graph(self, frame) -> None:
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
        ax.pie(non_zero_durations, labels=non_zero_names, autopct=_autopct_format(non_zero_durations), colors=self.data_manager.pie_colors, wedgeprops=dict(linewidth=1),
               textprops={"fontsize": pie_font_size, "family": pie_font_family, "color": self.data_manager.font_color}, counterclock=False, startangle=90)
        fig.set_size_inches(graph_width/100, graph_height/100, forward=True)
        ax.set_facecolor(self.data_manager.graph_fg_color)
        fig.set_facecolor(self.data_manager.graph_bg_color)
        ax.tick_params(colors=self.data_manager.font_color)
        ax.set_title("Duration of Study Sessions by Day of the Week", color=self.data_manager.font_color)
        ax.spines["top"].set_color(self.data_manager.spine_color)
        ax.spines["bottom"].set_color(self.data_manager.spine_color)
        ax.spines["left"].set_color(self.data_manager.spine_color)
        ax.spines["right"].set_color(self.data_manager.spine_color)
        plt.subplots_adjust(bottom=0.0)

        for wedge in ax.patches:
            wedge.set_edgecolor(self.data_manager.graph_fg_color)

        weekday_frame = FigureCanvasTkAgg(fig, master=frame)

        weekday_graph = weekday_frame.get_tk_widget()
        weekday_graph.pack(padx=3, pady=3)


    def create_total_time_graph(self, frame):
        dates = self.data_manager.date_list
        times = self.data_manager.duration_list

        # Calculate cumulative time
        cumulative_times = [sum(times[:i+1]) for i in range(len(times))]

        # Create subplot
        fig, ax = plt.subplots(figsize=(5, 5))

        # Plot
        ax.plot(dates, cumulative_times, color=self.data_manager.color)

        fig, ax = plt.subplots()
        ax.plot(dates, cumulative_times, color=self.data_manager.graph_color)
        ax.fill_between(dates, cumulative_times, color=self.data_manager.graph_color)
        ax.set_title("Cumulative Time By Date", color=self.data_manager.font_color)
        ax.tick_params(colors=self.data_manager.font_color)
        ax.set_facecolor(self.data_manager.graph_fg_color)
        fig.set_facecolor(self.data_manager.graph_bg_color)
        ax.spines["top"].set_color(self.data_manager.spine_color)
        ax.spines["bottom"].set_color(self.data_manager.spine_color)
        ax.spines["left"].set_color(self.data_manager.spine_color)
        ax.spines["right"].set_color(self.data_manager.spine_color)
        fig.set_size_inches(graph_width/100, graph_height/100, forward=True)
        ax.tick_params(axis="x", labelrotation = 45)

        def _format_func(value, tick_number):
            return f"{int(value)} m"
        
        plt.gca().yaxis.set_major_formatter(FuncFormatter(_format_func))
        date_format = mdates.DateFormatter("%d/%m")
        ax.xaxis.set_major_formatter(date_format)
        ax.xaxis.set_major_locator(MaxNLocator(integer=True, prune="both"))
        plt.subplots_adjust(bottom=0.15)

        total_time_frame = FigureCanvasTkAgg(fig, master=frame)

        total_time_graph = total_time_frame.get_tk_widget()
        total_time_graph.pack(padx=3, pady=3)

    
    def clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

    def create_graphs(self) -> None:
        self.statistics_frame.grid_propagate(False)
        plt.close("all")
        self.clear_frame(self.statistics_facts_frame)
        self.clear_frame(self.statistics_graph_frame)

        self.create_time_spent_graph(self._graph_gui_frame(0, 0))
        self.create_weekday_graph(self._graph_gui_frame(0, 1))
        self.create_total_time_graph(self._graph_gui_frame(0, 2))

        print("Graphs created.")
        if self.data_manager.duration_list:
            self.create_funfact(0, 6, "Longest Session", round(max(self.data_manager.duration_list), 1), "Minutes")
        else:
            self.create_funfact(0, 6, "Longest Session", "0", "Minutes")

        if self.data_manager.data_amount == 0:
            self.create_funfact(0, 0, "Average Study Duration", "0", "Minutes")
            self.create_funfact(0, 1, "Average Break Duration", "0", "Minutes")
            self.create_funfact(0, 2, "Average study Start Time", "00:00")
            self.create_funfact(0, 3, "Goal Met in", "0%", "of Sessions")
            self.create_funfact(0, 4, "Favorite subject", "")
            self.create_funfact(0, 5, "Most Productive Day", "")

        else:
            self.create_funfact(0, 0, "Average Study Duration", round(self.data_manager.total_duration/self.data_manager.data_amount, 1), "Minutes")
            self.create_funfact(0, 1, "Average Break Duration", round(self.data_manager.total_break_duration/self.data_manager.data_amount, 1), "Minutes")
            self.create_funfact(0, 2, "Average Study Start Time", self.data_manager.average_time)
            self.create_funfact(0, 4, "Favorite Subject", self.data_manager.most_common_subject, None, 3)
            self.create_funfact(0, 5, "Most Productive Day", self.data_manager.best_weekday, None, 2.7)
            if self.data_manager.goal_amount == 0:
                self.create_funfact(0, 3, "Goal Met in", "0%", "of Sessions")
            else:
                self.create_funfact(0, 3, "Goal Met in", str(int(self.data_manager.goal_amount / self.data_manager.data_amount * 100)) + "%", "of Sessions")

        print("Fun facts created.")


    def create_funfact(self, row_index: int, column_index: int, title: str, text: str, under_text: str = None, text_size: int = 4) -> None:
        frame = ctk.CTkFrame(self.statistics_facts_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=240, height=180)
        frame.grid(row=row_index, column=column_index, padx=frame_padding, pady=frame_padding)
        title = ctk.CTkLabel(frame, text=title, font=(font_family, int(font_size*1.2)), text_color=(light_font_color, font_color))
        title.place(anchor="n", relx=0.5, rely=0.05)
        if " " in str(text):
            text = text.replace(" ", "\n")
            text_size = 2.7
        text = ctk.CTkLabel(frame, text=text, font=(font_family, int(font_size*text_size)), text_color=(light_font_color, font_color))
        text.place(anchor="center", relx=0.5, rely=0.5)
        if under_text != None:
            label = ctk.CTkLabel(frame, text=under_text, font=(font_family, int(font_size*1.2)), text_color=(light_font_color, font_color))
            label.place(anchor="center", relx=0.5, rely=0.8)

    
    def forget_and_propagate(self, list: list) -> None:
        for item in list:
            item.grid_forget()
            item.grid_propagate(False)


    def switch_tab(self, selected_tab_button: ctk.CTkButton, tab_buttons: list[ctk.CTkButton]) -> None:
        """
        Switches the active tab and updates tab button colors accordingly.
        """

        #Change the color of all buttons to unselected color

        self.forget_and_propagate([self.main_frame, self.statistics_frame, self.settings_frame, self.achievements_frame, self.history_frame, self.notes_frame])


        def _show_selected_tab():
            frame = None
            match selected_tab_button.cget("text").replace(" ", ""):
                case "Timer":
                    frame = self.main_frame
                case "Statistics":
                    frame = self.statistics_frame
                case "Settings":
                    frame = self.settings_frame
                case "Achievements":
                    frame = self.achievements_frame
                case "History":
                    frame = self.history_frame
                case "Notes":
                    frame = self.notes_frame

            if frame is not None:
                if frame == self.statistics_frame:
                    frame.grid(column=2, row=0)
                    frame.grid_propagate(False)
                else:
                    frame.grid(column=2, row=0, padx=main_frame_pad_x)
                    frame.grid_propagate(False)

        _show_selected_tab()


        def _decolor_tabs() -> None:
            for button in tab_buttons:
                button.configure(fg_color=(light_tab_color, tab_color), hover_color=(light_tab_highlight_color, tab_highlight_color))

        _decolor_tabs()

        #Change color of pressed button to selected color
        selected_tab_button.configure(fg_color=(light_tab_selected_color, tab_selected_color), hover=False)



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
        self.data_manager.create_total_data()


    def save_data(self) -> None:
        time_in_minutes = self.timer_manager.timer_time / 60

        #Only be able to save if time is higher than 1m
        if time_in_minutes >= 1:
            self.unlock_widgets()

            if time_in_minutes >= self.goal:
                self.data_manager.increase_goal_streak()

            self.data_manager.save_data()

            self.collect_data()
            self.reset_gui_values()
            self.update_streak_values()
            self.load_history()

            self.data_manager.load_theme()
            
            self.notification_limit_on = False

            self.create_achievements()
        else:
            print("No data to save. (time less than 1m)")


    def save_autobreak(self) -> None:
        frequency_input = self.frequency_input.get()[:2]
        if "." in frequency_input:
            frequency_input = (frequency_input.split(".")[0]).replace(".", "")

        if len(frequency_input) == 0 or int(frequency_input) == 0:
            frequency_input = self.data_manager.autobreak_frequency
        else:
            if self.data_manager.autobreak_frequency != self.frequency_input.cget("placeholder_text"):
                if len(frequency_input) == 0 or int(frequency_input) < 1:
                    frequency_input = "25"
                elif not frequency_input.isdigit():
                    return
    
        duration_input = self.duration_input.get()[:2]
        if "." in duration_input:
            duration_input = (duration_input.split(".")[0]).replace(".", "")

        if len(duration_input) == 0 or int(duration_input) == 0:
            duration_input = self.data_manager.autobreak_duration
        else:
            if self.data_manager.autobreak_duration != self.duration_input.cget("placeholder_text"):
                if len(duration_input) == 0 or int(duration_input) < 1:
                    duration_input = "5"
                elif not duration_input.isdigit():
                    return
            
        
        self.frequency_input.delete("0", "end")
        self.duration_input.delete("0", "end")
        self.frequency_input.insert("0", frequency_input)
        self.duration_input.insert("0", duration_input)

        if self.autobreak_switch.get() == "On":
            switch = "On"
        else:
            switch = "Off"
        
        self.data_manager.save_autobreak(frequency_input, duration_input, switch)


    def eye_protection(self):
        checkbox = self.eye_care_checkbox.get()
        time_between = 60 * 20
        
        if self.eye_care_selection.get() == "On":
            if checkbox == "On" and self.timer_manager.timer_running:
                self.send_notification("Eye Protection", "It's time for a 20/20/20 break! Look away for 20 seconds at something 20 feet away.")
            else:
                self.send_notification("Eye Protection", "It's time for a 20/20/20 break! Look away for 20 seconds at something 20 feet away.")
        
        # Schedule the next iteration
        if self.eye_care_selection.get() == "On":
            self.WINDOW.after(time_between * 1000, self.eye_protection)
        else:
            self.WINDOW.after(10 * 1000, self.eye_protection)


    def auto_break(self):
        time_between = self.data_manager.autobreak_frequency * 60

        def try_timer():
            if self.timer_manager.break_time % self.data_manager.autobreak_duration == 0: 
                self.timer_manager.timer_mechanism(self.timer_button, self.break_button, self.time_display_label)
            else:
                self.WINDOW.after(1000, try_timer)
        
        if self.autobreak_switch.get() == "On" and self.timer_manager.timer_running and self.timer_manager.timer_time % time_between == 0:
            self.timer_manager.break_mechanism(self.break_button, self.timer_button, self.break_display_label)
            self.send_notification("Auto-break", f"Time for a {self.data_manager.autobreak_duration}-minute")
            self.WINDOW.after(self.data_manager.autobreak_duration * 60 * 1000, try_timer)
            round(self.timer_manager.break_time, -1)
        
        self.WINDOW.after(100, self.auto_break)  # Schedule next iteration


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
        reset_messagebox = CTkMessagebox(title="Reset data", message="Are you sure you want to reset data?", icon="warning", option_1="No", option_2="Yes", button_color=self.data_manager.color,
                                         button_hover_color=self.data_manager.highlight_color, font=(font_family, font_size), text_color=self.data_manager.font_color, button_text_color="black")
        if reset_messagebox.get() == "Yes":
            os.remove(self.data_file)
            self.restart_program()

            print("Data reset.")


    def scroll_statistics(self, direction: str) -> None:
        if self.scrolling:
            return
        
        if direction == "right":
            if self.statistics_scroll_position == "left":
                self.scroll_smoothly(0.465/2, 0-0.05, self.statistics_scroll_frame)
                self.statistics_scroll_position = "middle"

            elif self.statistics_scroll_position == "middle":
                self.scroll_smoothly(0.465+0.05, 0.465/2, self.statistics_scroll_frame)
                self.statistics_scroll_position = "right"

        if direction == "left":
            if self.statistics_scroll_position == "right":
                self.scroll_smoothly(0.465/2, 0.465+0.05, self.statistics_scroll_frame)
                self.statistics_scroll_position = "middle"

            elif self.statistics_scroll_position == "middle":
                self.scroll_smoothly(0-0.05, 0.465/2, self.statistics_scroll_frame)
                self.statistics_scroll_position = "left"

    
    def scroll_achievements(self, direction: str) -> None:
        if self.scrolling:
            return
        
        if direction == "right":
            if self.achievements_scroll_position  == "left":
                self.scroll_smoothly(0.465, 0, self.achievements_scroll_frame)
                self.achievements_scroll_position = "right"

        elif direction == "left":
            if self.achievements_scroll_position  == "right":
                self.scroll_smoothly(-0.005, 0.465, self.achievements_scroll_frame)
                self.achievements_scroll_position = "left"


    def scroll_smoothly(self, destination: float, position: float, frame: ctk.CTkScrollableFrame):        
        if position < destination:
            position += 0.005
            if position < destination:
                self.scrolling = True
                self.WINDOW.after(3, self.scroll_smoothly, destination, position, frame)
            else:
                self.WINDOW.after(50, self.enable_scrolling)
                return
        elif destination < position:
            position -= 0.005
            if destination < position:
                self.scrolling = True
                self.WINDOW.after(3, self.scroll_smoothly, destination, position, frame)
            else:
                self.WINDOW.after(50, self.enable_scrolling)
                return
        frame._parent_canvas.xview_moveto(position)



    def enable_scrolling(self):
        self.scrolling = False


    #Get all buttons other than tab buttons
    def create_widget_list(self) -> None:
        frame_list = []
        widgets = []

        for frame in self.WINDOW.winfo_children():
            if isinstance(frame, ctk.CTkFrame):
                frame_list.append(frame)
                
        frame_list.pop(0)
        frame_list.pop(1)

        def _get_widgets(frame):
            for child in frame.winfo_children():
                widgets.append(child)
                if isinstance(child, ctk.CTkFrame):
                    _get_widgets(child)

        for frame in frame_list:
            _get_widgets(frame)

        for widget in widgets:
            if isinstance(widget, ctk.CTkButton) or isinstance(widget, ctk.CTkProgressBar):
                if ".!ctkframe5" in str(widget) and isinstance(widget, ctk.CTkButton):
                    pass
                else:
                    self.widget_list.append(widget)


    def lock_widgets(self):
        self.progressbar.configure(progress_color=self.data_manager.color)
        self.frequency_input.configure(state="disabled")
        self.frequency_input.insert("end", self.data_manager.autobreak_frequency)
        self.duration_input.configure(state="disabled")
        self.duration_input.insert("end", self.data_manager.autobreak_duration)

        self.goal_dropdown.configure(state="disabled")
        self.subject_selection.configure(state="disabled")
        self.autobreak_switch.configure(state="disabled")

        self.autobreak_button.configure(state="disabled", fg_color="grey")
        self.subject_button.configure(state="disabled", fg_color="grey")
        self.goal_button.configure(state="disabled", fg_color="grey")
        if self.autobreak_switch.get() == "On":
            self.break_button.configure(state="disabled", fg_color = "grey", command=None, hover=False)


    def unlock_widgets(self):
        self.progressbar.configure(progress_color=(light_border_frame_color, border_frame_color))
        self.frequency_input.configure(state="normal")
        self.frequency_input.delete("end")
        self.duration_input.configure(state="normal")
        self.duration_input.delete("end")
        self.autobreak_button.configure(state="normal", fg_color=button_color)
        self.break_button.configure(state="normal", fg_color=button_color, command=lambda: self.timer_manager.break_mechanism(self.break_button, self.timer_button, self.break_display_label), hover=True)
        self.subject_button.configure(state="normal", fg_color=button_color)
        self.goal_button.configure(state="normal", fg_color=button_color)
        self.goal_dropdown.configure(state="normal")
        self.subject_selection.configure(state="normal")
        self.autobreak_switch.configure(state="normal")


    def save_on_quit(self) -> None:
        self.save_data()
        print("Data saved on exit.")

        self.workbook.save(self.data_file)

        self.WINDOW.destroy()


    def change_focus(self, event) -> None:
        event.widget.focus_set()


    def run(self) -> None:
        self.WINDOW.mainloop()

    
    def restart_program(self) -> None:
        python = sys.executable
        os.execl(python, python, *sys.argv)


if __name__ == "__main__":
    App().run()
