import datetime
import customtkinter as ctk

class TimerManager:
    def __init__(self, MainTabManager: object, window: object) -> None:
        self.main_tab_manager = MainTabManager
        self.window = window
        self.initialize_variables()


    def initialize_variables(self) -> None:
        self.timer_running = False
        self.break_running = False
        self.timer_time = 0
        self.break_time = 0


    def timer_mechanism(self, timer_button: ctk.CTkButton, break_button: ctk.CTkButton, time_display_label: ctk.CTkLabel) -> None:
        self.time_display_label = time_display_label
        if not self.timer_running:
            self.timer_running = True
            self.break_running = False
            timer_button.configure(text="Stop")
            break_button.configure(text="Start")
            self._update_time()

        elif self.timer_running:
            self.timer_running = False
            timer_button.configure(text="Start")
    

    def _update_time(self) -> None:
        if self.timer_running:
            self.timer_time += 1
            self.time_display_label.configure(text=str(datetime.timedelta(seconds=self.timer_time)))
            self.main_tab_manager.reach_goal(self.timer_time)
            self.window.after(1000, self._update_time)


    def break_mechanism(self, break_button: ctk.CTkButton, timer_button: ctk.CTkButton, break_display_label: ctk.CTkLabel) -> None:
        self.break_display_label = break_display_label
        if not self.break_running:
            self.break_running = True
            self.timer_running = False
            break_button.configure(text="Stop")
            timer_button.configure(text="Start")
            self._update_break_time()
        elif self.break_running:
            self.break_running = False
            self.break_button.configure(text="Start")


    def _update_break_time(self) -> None:
        if self.break_running:
            self.break_time += 1
            self.break_display_label.configure(text=str(datetime.timedelta(seconds=self.break_time)))
            self.window.after(1000, self._update_break_time)