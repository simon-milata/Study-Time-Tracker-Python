import datetime
import os
import random

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

        self.window_setup()
        self.initialize_variables()
        self.create_gui()
        self.file_setup()


    def create_gui(self):
        self.main_frame_gui_setup()
        self.tab_frames_gui_setup()
        self.secondary_frames_gui_setup()

        self.timer_gui()
        self.break_gui()
        self.save_data_gui()
        self.goal_gui_setup()
        self.progress_gui_setup()
        self.streak_gui_setup()
        self.settings_gui_setup()


    def file_setup(self):
        self.local_folder = os.path.expandvars(rf"%APPDATA%\{self.APPNAME}")
        self.data_file = os.path.expandvars(rf"%APPDATA%\{self.APPNAME}\{self.FILENAME}")

        os.makedirs(self.local_folder, exist_ok=True)

        self.timer_manager = TimerManager(self, self.WINDOW)

        if os.path.isfile(self.data_file):
            self.workbook = op.load_workbook(self.data_file)
            self.worksheet = self.workbook.active

            self.data_manager = DataManager(self, self.timer_manager, self.workbook, self.worksheet)

            self.collect_data()
            self.update_streak_values()
            print("File loaded")

        else:
            self.workbook = op.Workbook()
            self.worksheet = self.workbook.active

            self.workbook.save(self.data_file)

            self.data_manager = DataManager(self, self.timer_manager, self.workbook, self.worksheet)

            self.data_manager.initialize_new_file_variables()
            self.data_manager.customize_excel()

            print("New file created")


    def initialize_variables(self):
        self.default_choice = ctk.StringVar(value="1 hour")
        self.notification_limit = False
        self.goal = 60


    def window_setup(self):
        self.WINDOW = ctk.CTk()
        self.WINDOW.geometry(str(WIDTH + BORDER_WIDTH + main_frame_pad_x + tab_frame_width) + "x" + str(HEIGHT+((widget_padding_x+frame_padding)*2)))
        self.WINDOW.title(self.APPNAME)
        self.WINDOW.configure(background=window_color)
        self.WINDOW.resizable(False, False)
        self.WINDOW.grid_propagate(False)


    def main_frame_gui_setup(self):
        self.tab_frame = ctk.CTkFrame(self.WINDOW, width=tab_frame_width, height=HEIGHT+((widget_padding_x+frame_padding)*2), fg_color=tab_frame_color)
        self.tab_frame.grid(column=0, row=0)
        self.tab_frame.pack_propagate(False)

        self.main_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.main_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.main_frame.grid_propagate(False)

        self.statistics_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.statistics_frame.grid(column=2, row=0, padx=main_frame_pad_x)

        self.settings_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.settings_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.settings_frame.grid_forget()

        self.achievements_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.achievements_frame.grid(column=2, row=0, padx=main_frame_pad_x)

        self.history_frame = ctk.CTkFrame(self.WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.history_frame.grid(column=2, row=0, padx=main_frame_pad_x)

        self.forget_and_propagate(list = [self.statistics_frame, self.settings_frame, self.achievements_frame, self.history_frame])


    def tab_frames_gui_setup(self):
        self.timer_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=tab_color)
        self.timer_tab.pack(pady=tab_padding_y)
        self.timer_tab_button = ctk.CTkButton(self.timer_tab, text="Timer", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=font_color,
                                              fg_color=tab_color, width=int(tab_frame_width*0.95), height=int(tab_height*0.7), hover_color=tab_highlight_color, 
                                              anchor="w", command=lambda: self.switch_tab("main"))
        self.timer_tab_button.place(relx=0.5, rely=0.5, anchor="center")

        self.statistics_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=tab_color)
        self.statistics_tab.pack(pady=tab_padding_y)
        self.statistics_tab_button = ctk.CTkButton(self.statistics_tab, text="Statistics", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=font_color,
                                                   fg_color=tab_color, width=int(tab_frame_width*0.95), height=int(tab_height*0.7), hover_color=tab_highlight_color, 
                                                   anchor="w", command=lambda: self.switch_tab("statistics"))
        self.statistics_tab_button.place(relx=0.5, rely=0.5, anchor="center")

        self.achievements_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=tab_color)
        self.achievements_tab.pack(pady=tab_padding_y)
        self.achievements_tab_button = ctk.CTkButton(self.achievements_tab, text="Achievements", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=font_color,
                                                     fg_color=tab_color, width=int(tab_frame_width*0.95), height=int(tab_height*0.7), hover_color=tab_highlight_color, 
                                                     anchor="w", command=lambda: self.switch_tab("achievements"))
        self.achievements_tab_button.place(relx=0.5, rely=0.5, anchor="center")

        self.history_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=tab_color)
        self.history_tab.pack(pady=tab_padding_y)
        self.history_tab_button = ctk.CTkButton(self.history_tab, text="History", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=font_color,
                                                fg_color=tab_color, width=int(tab_frame_width*0.95), height=int(tab_height*0.7), hover_color=tab_highlight_color, 
                                                anchor="w", command=lambda: self.switch_tab("history"))
        self.history_tab_button.place(relx=0.5, rely=0.5, anchor="center")
        self.settings_tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=tab_color)
        self.settings_tab.place(relx=0.5, rely=1, anchor="s")
        settings_tab_button = ctk.CTkButton(self.settings_tab, text="Settings", font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=font_color,
                                 fg_color=tab_color, width=int(tab_frame_width*0.95), height=int(tab_height*0.7), hover_color=tab_highlight_color, anchor="w", command=lambda: self.switch_tab("settings"))
        settings_tab_button.place(relx=0.5, rely=0.5, anchor="center")


    def secondary_frames_gui_setup(self):
        self.timer_break_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width)
        self.timer_break_frame.grid(row=0, column=1)
        self.timer_break_frame.pack_propagate(False)

        self.goal_progress_frame = ctk.CTkFrame(self.main_frame, height=(HEIGHT-button_height*1.5), width=frame_width)
        self.goal_progress_frame.grid(row=0, column=0)
        self.goal_progress_frame.pack_propagate(False)


    def goal_gui_setup(self):
        goal_frame = ctk.CTkFrame(self.goal_progress_frame, fg_color=frame_color, height=175, width=frame_width, corner_radius=10)
        goal_frame.pack(padx=frame_padding, pady=frame_padding)
        goal_frame.pack_propagate(False)

        goal_label = ctk.CTkLabel(goal_frame, text="Goal", font=(font_family, font_size), text_color=font_color)
        goal_label.place(anchor="nw", relx=0.05, rely=0.05)

        self.goal_dropdown = ctk.CTkComboBox(goal_frame, values=["1 minutes", "30 minutes", "1 hour", "1 hour, 30 minutes", "2 hours", "2 hours, 30 minutes", "3 hours", "3 hours, 30 minutes",
                                                            "4 hours", "4 hours, 30 minutes", "5 hours", "5 hours, 30 minutes", "6 hours"], variable=self.default_choice, 
                                                            state="readonly", width=200, height=30, dropdown_font=(font_family, int(font_size*0.75)),
                                                            font=(font_family, int(font_size)), fg_color=border_frame_color, button_color=border_frame_color)
        self.goal_dropdown.place(anchor="center", relx=0.5, rely=0.45)

        goal_button = ctk.CTkButton(goal_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                height=button_height, command=self.set_goal)
        goal_button.place(anchor="s", relx=0.5, rely=0.9)


    def progress_gui_setup(self):
        progress_frame = ctk.CTkFrame(self.goal_progress_frame, fg_color=frame_color, width=frame_width, corner_radius=10, height=100)
        progress_frame.pack(padx=frame_padding, pady=frame_padding)
        progress_frame.pack_propagate(False)

        progress_label = ctk.CTkLabel(progress_frame, text="Progress", font=(font_family, int(font_size)), text_color=font_color)
        progress_label.place(anchor="nw", relx=0.05, rely=0.05)

        self.progressbar = ctk.CTkProgressBar(progress_frame, height=20, width=220, progress_color=button_color, fg_color=border_frame_color, corner_radius=10)
        self.progressbar.place(anchor="center", relx=0.5, rely=0.65)
        self.progressbar.set(0)


    def streak_gui_setup(self):
        streak_frame = ctk.CTkFrame(self.goal_progress_frame, fg_color=frame_color, width=frame_width, corner_radius=10, height=220)
        streak_frame.pack(padx=frame_padding, pady=frame_padding)

        streak_label = ctk.CTkLabel(streak_frame, text="Streak", font=(font_family, int(font_size)), text_color=font_color)
        streak_label.place(anchor="nw", relx=0.05, rely=0.05)
        
        times_studied_text = ctk.CTkLabel(streak_frame, text="Goal\nreached", font=(font_family, int(font_size/1.25)), text_color=font_color)
        times_studied_text.place(anchor="center", relx=0.3, rely=0.4)
        self.times_goal_reached = ctk.CTkLabel(streak_frame, text=0, font=(font_family, int(font_size*2.7)), text_color=font_color)
        self.times_goal_reached.place(anchor="center", relx=0.3, rely=0.6)
        times_reached_label = ctk.CTkLabel(streak_frame, text="times", font=(font_family, int(font_size/1.25)), text_color=font_color)
        times_reached_label.place(anchor="center", relx=0.3, rely=0.8)

        duration_studied_text = ctk.CTkLabel(streak_frame, text="Time\nstudied", font=(font_family, int(font_size/1.25)), text_color=font_color)
        duration_studied_text.place(anchor="center", relx=0.7, rely=0.4)
        self.streak_duration = ctk.CTkLabel(streak_frame, text=0, font=(font_family, int(font_size*2.7)), text_color=font_color)
        self.streak_duration.place(anchor="center", relx=0.7, rely=0.6)
        duration_minute_label = ctk.CTkLabel(streak_frame, text="minutes", font=(font_family, int(font_size/1.25)), text_color=font_color)
        duration_minute_label.place(anchor="center", relx=0.7, rely=0.8)


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


    def settings_gui_setup(self):
        color_select_frame = ctk.CTkFrame(self.settings_frame, fg_color=frame_color, height=200, width=int(frame_width/1.25), corner_radius=10)
        color_select_frame.grid(column=0, row=0, padx=frame_padding, pady=frame_padding)
        color_label = ctk.CTkLabel(color_select_frame, text="Color", font=(font_family, font_size), text_color=font_color)
        color_label.place(anchor="nw", relx=0.05, rely=0.05)
        color_dropdown = ctk.CTkComboBox(color_select_frame, values=["Orange", "Green", "Blue"], variable=default_color, state="readonly", width=150, height=30, 
                                         dropdown_font=(font_family, int(font_size*0.75)), font=(font_family, int(font_size)), fg_color=border_frame_color, button_color=border_frame_color)
        color_dropdown.place(anchor="center", relx=0.5, rely=0.45)

        reset_frame = ctk.CTkFrame(self.settings_frame, fg_color=tab_color)
        reset_frame.place(anchor="s", relx=0.5, rely=0.985)
        reset_data_btn = ctk.CTkButton(reset_frame, text="Reset Data", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=self.reset_data, width=450)
        reset_data_btn.pack()

    
    def forget_and_propagate(self, list):
        for item in list:
            item.grid_forget()
            item.grid_propagate(False)


    def switch_tab(self, tab = str):
        tabs = {
            "main": self.main_frame,
            "statistics": self.statistics_frame,
            "settings": self.settings_frame,
            "achievements": self.achievements_frame,
            "history": self.history_frame
        }

        tab_list = [self.main_frame, self.statistics_frame, self.settings_frame, self.achievements_frame, self.history_frame]
        tab_list.remove(tabs[tab])

        def forget_tabs():
            for frame in tab_list:
                frame.grid_forget()
        forget_tabs()

        tabs[tab].grid(column=2, row=0, padx=main_frame_pad_x)
        tabs[tab].grid_propagate(False)


    def timer_mechanism(self):
        self.timer_manager.timer_mechanism(self.timer_button, self.break_button, self.time_display_label)
        if self.timer_manager.timer_time <= 1:
            self.data_manager.get_start_time()


    def reset_gui_values(self):
        self.times_goal_reached.configure(text=0)
        self.streak_duration.configure(text=0)
        self.progressbar.set(0)
        self.reset_timers()


    def reset_timers(self):
        self.timer_button.configure(text="Start")
        self.break_button.configure(text="Start")
        self.time_display_label.configure(text="0:00:00")
        self.break_display_label.configure(text="0:00:00") 


    def set_goal(self):
        x = 0
        choice = self.goal_dropdown.get()
        if "hour" in choice:
            x += int(choice.split(" ")[0]) * 60
        if "minutes" in choice and "hour" in choice:
            x += int(choice.split(", ")[1].removesuffix(" minutes"))
        if "hour" not in choice:
            x += int(choice.split(" ")[0])
        self.goal = x


    def reach_goal(self, timer_time):
        if (timer_time/60) < self.goal:
            self.progressbar.set((timer_time/60)/self.goal)
        elif timer_time/60 >= self.goal and not self.notification_limit:
            self.progressbar.set(1)

            message = random.choice(["Congratulations! You've reached your study goal. Take a well-deserved break and recharge!", "Study session complete! Great job on reaching your goal. Time for a quick break!",
                                     "You did it! Study session accomplished. Treat yourself to a moment of relaxation!", "Well done! You've met your study goal. Now, take some time to unwind and reflect on your progress.",
                                     "Study session over! You've achieved your goal. Reward yourself with a brief pause before your next task.", "Goal achieved! Take a breather and pat yourself on the back for your hard work.",
                                     "Mission accomplished! You've hit your study target. Enjoy a short break before diving back in.", "Study session complete. Nicely done! Use this time to relax and rejuvenate before your next endeavor.",
                                     "You've reached your study goal! Treat yourself to a well-deserved break. You've earned it!", "Goal achieved! Take a moment to celebrate your success. Your dedication is paying off!"])
            self.send_notification("Study Goal Reached", message)
            print(self.notification_limit)

    def update_streak_values(self):
        self.times_goal_reached.configure(text=self.data_manager.goal_amount)
        self.streak_duration.configure(text=self.data_manager.total_duration)

    
    def break_mechanism(self):
        self.timer_manager.break_mechanism(self.break_button, self.timer_button, self.break_display_label)


    def collect_data(self):
        self.data_manager.collect_data()
        self.data_manager.data_to_variable()


    def save_data(self):
        if self.timer_manager.timer_time > 60:
            if self.timer_manager.timer_time/60 >= self.goal:
                self.data_manager.increase_goal_streak()

            self.data_manager.save_data()
            self.collect_data()
            self.reset_gui_values()
            self.update_streak_values()

            self.notification_limit = False
        else:
            print("No data to save. (time less than 1m)")

    
    def send_notification(self, title, message):
        toast = Notification(app_id=self.APPNAME, title=title, msg=message)
        toast.show()
        self.notification_limit = True
        print("Notification " + title + " sent.")


    def reset_data(self):
        del self.workbook[self.workbook.active.title]
        self.workbook.create_sheet()
        self.worksheet = self.workbook.active

        self.data_manager.reset_data(self.workbook, self.worksheet)

        self.reset_gui_values()


    def run(self):
        self.WINDOW.mainloop()


if __name__ == "__main__":
    App().run()
