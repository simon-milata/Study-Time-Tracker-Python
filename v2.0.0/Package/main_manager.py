import random

import customtkinter as ctk

from .styles import *
from .timer_management import TimerManager

class MainManager:
    def __init__(self, App, WINDOW: object, TimerManager: object) -> None:
        self.app = App
        self.WINDOW = WINDOW

        self.initialize_variables()

        self.timer_manager = TimerManager

        self.main_gui = MainGUI(self, self.WINDOW, self.timer_manager)

        self.main_gui._initialize_gui()

    
    def show_main_gui(self) -> None:
        self.main_gui.main_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.main_gui.main_frame.grid_propagate(False)


    def hide_main_gui(self) -> None:
        self.main_gui.main_frame.grid_forget()

    
    def initialize_variables(self) -> None:
        self.widget_list = []
        self.default_choice = ctk.StringVar(value="1 hour")
        self.notification_limit_on = False
        self.goal = 60


    def set_goal(self) -> None:
        x = 0
        choice = self.main_gui.goal_dropdown.get()

        #Make an int out of a string e.g. "1 hour, 30 minutes"
        if "hour" in choice:
            x += int(choice.split(" ")[0]) * 60
        if "minutes" in choice and "hour" in choice:
            x += int(choice.split(", ")[1].removesuffix(" minutes"))
        if "hour" not in choice:
            x += int(choice.split(" ")[0])
        self.goal = x

        self.main_gui.progressbar.set((self.timer_manager.timer_time / 60) / self.goal)

    
    def reach_goal(self, timer_time: int) -> None:
        time_in_minutes = timer_time / 60
        if time_in_minutes < self.goal:
            self.main_gui.progressbar.set(time_in_minutes/self.goal)
        elif time_in_minutes >= self.goal and not self.app.notification_limit_on:
            self.main_gui.progressbar.set(1)

            message = random.choice(["Congratulations! You've reached your study goal. Take a well-deserved break and recharge!", "Study session complete! Great job on reaching your goal. Time for a quick break!",
                                     "You did it! Study session accomplished. Treat yourself to a moment of relaxation!", "Well done! You've met your study goal. Now, take some time to unwind and reflect on your progress.",
                                     "Study session over! You've achieved your goal. Reward yourself with a brief pause before your next task.", "Goal achieved! Take a breather and pat yourself on the back for your hard work.",
                                     "Mission accomplished! You've hit your study target. Enjoy a short break before diving back in.", "Study session complete. Nicely done! Use this time to relax and rejuvenate before your next endeavor.",
                                     "You've reached your study goal! Treat yourself to a well-deserved break. You've earned it!", "Goal achieved! Take a moment to celebrate your success. Your dedication is paying off!"])
            self.app.send_notification("Study Goal Reached", message)


class MainGUI:
    def __init__(self, MainManager, WINDOW, TimerManager):
        self.main_manager = MainManager
        self.WINDOW = WINDOW

        self.timer_manager = TimerManager

    def _initialize_gui(self) -> None:
        self._main_frame_gui_setup()
        self._secondary_frames_gui_setup()

        self._goal_gui_setup()
        self._progress_gui_setup()
        self._streak_gui_setup()
        self._timer_gui_setup()
        self._break_gui_setup()
        self._save_data_gui()


    def _main_frame_gui_setup(self) -> None:
        self.main_frame = ctk.CTkFrame(self.WINDOW, fg_color=(light_main_frame_color, main_frame_color), height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.main_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.main_frame.grid_propagate(False)


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
                                                            "4 hours", "4 hours, 30 minutes", "5 hours", "5 hours, 30 minutes", "6 hours"], variable=self.main_manager.default_choice, 
                                                            state="readonly", width=200, height=30, dropdown_font=(font_family, int(font_size*0.75)), border_color=(light_border_frame_color, border_frame_color),
                                                            font=(font_family, int(font_size)), fg_color=(light_border_frame_color, border_frame_color), button_color=(light_border_frame_color, border_frame_color))
        self.goal_dropdown.place(anchor="center", relx=0.5, rely=0.45)

        self.goal_button = ctk.CTkButton(goal_frame, text="Save", font=(font_family, font_size), text_color=button_font_color, fg_color=button_color, hover_color=button_highlight_color,
                                height=button_height, command=self.main_manager.set_goal)
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
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=lambda: self.timer_manager.timer_mechanism(self.timer_button, self.break_button, self.time_display_label))
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
                                        border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=lambda: self.timer_manager.break_mechanism(self.break_button, self.timer_button, self.break_display_label))
        self.break_button.place(anchor="s", relx=0.5, rely=0.9)

    def _save_data_gui(self) -> None:
        self.data_frame = ctk.CTkFrame(self.main_frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=WIDTH-10, height=button_height*2)
        self.data_frame.place(anchor="s", relx=0.5, rely=0.985)
        self.data_frame.grid_propagate(False)
        self.save_data_button = ctk.CTkButton(self.data_frame, text="Save Data", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                    border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, width=450)
        self.save_data_button.place(relx=0.5, anchor="center", rely=0.5)