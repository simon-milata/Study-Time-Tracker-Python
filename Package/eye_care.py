import time

class EyeCare:
    def __init__(self, app):
        self.app = app
    def eye_protection(self):
        time_between = 1000 * 60 * 20
        if self.app.eye_care_selection.get() == "On":
            time.sleep(time_between)
            self.eye_protection()