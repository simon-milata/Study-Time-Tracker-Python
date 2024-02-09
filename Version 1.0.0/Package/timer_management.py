import datetime

class TimerManager:
    def __init__(self, App, window):
        self.app = App
        self.window = window
        self.initialize_variables()


    def initialize_variables(self):
        self.timer_running = False
        self.break_running = False
        self.timer_time = 0
        self.break_time = 0


    def timer_mechanism(self, timer_button, break_button, time_display_label):
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
    

    def _update_time(self):
        if self.timer_running:
            self.timer_time += 1
            self.time_display_label.configure(text=str(datetime.timedelta(seconds=self.timer_time)))
            self.app.reach_goal(self.timer_time)
            self.window.after(1000, self._update_time)


    def break_mechanism(self, break_button, timer_button, break_display_label):
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


    def _update_break_time(self):
        if self.break_running:
            self.break_time += 1
            self.break_display_label.configure(text=str(datetime.timedelta(seconds=self.break_time)))
            self.window.after(1000, self._update_break_time)