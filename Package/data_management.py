from collections import Counter

import datetime
from openpyxl.styles import Font

import customtkinter as ctk
import darkdetect
from .styles import *
from .note_management import NotesManager

class DataManager:
    def __init__(self, App, timer_manager, workbook, worksheet):
        self.app = App
        self.timer_manager = timer_manager
        self.workbook = workbook
        self.worksheet = worksheet
        self.initialize_variables()

        self.notes_manager = NotesManager(self.app, self)


    def initialize_variables(self) -> None:
        self.day_name_list = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        self.date_list = []
        self.duration_list = []
        self.total_duration = 0
        self.break_list = []
        self.total_break_duration = 0
        self.hours_list = []
        self.subject_list = []
        self.average_time = "00:00"
        self.graph_color = "#f38064"
        self.graph_bg_color = graph_bg_color
        self.graph_fg_color = graph_fg_color
        self.font_color = font_color
        self.spine_color = border_frame_color
        

    def initialize_new_file_variables(self) -> None:
        self.goal_amount = 0
        self.data_amount = 0
        self.notes_amount = 0
        self.monday_duration = self.tuesday_duration = self.wednesday_duration = self.thursday_duration = self.friday_duration = self.saturday_duration = self.sunday_duration = 0
        self.color_name = "Orange"
        if darkdetect.isDark():
            self.theme_name = "Dark"
        else:
            self.theme_name = "Light"

        self.customize_excel()
        self.save_color()
        self.save_theme()

    
    def collect_data(self) -> None:
        self.data_amount = int(self.worksheet["Z2"].value)
        self.goal_amount = int(self.worksheet["R2"].value)
        self.notes_amount = int(self.worksheet["N9"].value)

        self.collect_day_data()

        print("Data collected.")


    def data_to_variable(self) -> None:
        self.clear_graph_lists()

        for data in range(2, self.data_amount + 2):
            if "/" in str(self.worksheet["B" + str(data)].value):
                self.date_list.append(datetime.datetime.strptime(str(self.worksheet["B" + str(data)].value).split(" ")[0], "%d/%m/%Y").date())
            elif "-" in str(self.worksheet["B" + str(data)].value):
                self.date_list.append(datetime.datetime.strptime(str(self.worksheet["B" + str(data)].value).split(" ")[0], "%Y-%m-%d").date())
            self.duration_list.append(round(self.worksheet["C" + str(data)].value))

        self.total_duration = sum(self.duration_list)
        
        self.create_total_data()


    def create_total_data(self):
        def get_sec(time: str) -> int:
            """Get seconds from time."""
            h, m = time.split(':')
            return int(h) * 3600 + int(m) * 60
        
        for data in range(2, self.data_amount + 2):
            self.break_list.append(float(self.worksheet["D" + str(data)].value))
            self.hours_list.append(get_sec((self.worksheet["A" + str(data)].value.split(" ")[1])))
            self.subject_list.append(self.worksheet["E" + str(data)].value)
                                   
        self.total_break_duration = sum(self.break_list)
        self.average_time = str(datetime.timedelta(seconds=(round(sum(self.hours_list) / len(self.hours_list)))))[:5]
        self.most_common_subject = Counter(self.subject_list).most_common(1)[0][0]


    def save_data(self) -> None:
        self.initialize_variables()
        self.set_weekday()
        self.save_color()
        
        self.data_amount += 1

        self.stop_time = datetime.datetime.now()

        self.duration = self.calculate_duration()

        self.workbook.save(self.app.data_file)

        self.write_to_excel()

        print("Data saved.")

        self.timer_manager.initialize_variables()
        self.app.reset_timers()


    def write_to_excel(self) -> None:
        self.worksheet["A" + str((self.data_amount + 1))].value = self.start_time.strftime("%d/%m/%Y %H:%M")
        self.worksheet["B" + str((self.data_amount + 1))].value = self.stop_time.strftime("%d/%m/%Y %H:%M")
        self.worksheet["C" + str((self.data_amount + 1))].value = self.duration
        self.worksheet["D" + str((self.data_amount + 1))].value = self.timer_manager.break_time/60
        self.worksheet["E" + str((self.data_amount + 1))].value = self.app.subject_selection.get()

        self.worksheet["N9"].value = self.notes_amount

        self.worksheet["R2"].value = self.goal_amount

        self.worksheet["Z2"].value = self.data_amount
        self.workbook.save(self.app.data_file)


    def customize_excel(self) -> None:
        self.worksheet["A1"].value = "Start:"
        self.worksheet["B1"].value = "End:"
        self.worksheet["C1"].value = "Duration:"
        self.worksheet["D1"].value = "Break:"
        self.worksheet["E1"].value = "Subject:"

        self.worksheet["N8"].value = "Notes amount:"
        self.worksheet["N9"].value = self.notes_amount
        self.worksheet["N11"].value = "Notes:"
        self.worksheet["M12"].value = "Deleted:"
        self.worksheet["N12"].value = "Date:"
        self.worksheet["O12"].value = "Title:"
        self.worksheet["P12"].value = "Text:"

        self.worksheet["P1"].value = "Autobreak:"

        self.worksheet["Q1"].value = "Eye care:"
        self.worksheet["Q4"].value = "Only when timer running:"

        self.worksheet["R1"].value = "Goals reached:"
        self.worksheet["R2"].value = self.goal_amount

        self.worksheet["S1"].value = "Subject:"

        self.worksheet["T1"].value = "Color:"
        self.worksheet["T2"].value = self.color_name

        self.worksheet["U1"].value = "Theme:"
        self.worksheet["U2"].value = self.theme_name

        self.worksheet["W1"].value = "Weekday duration:"

        self.worksheet["Z1"].value = "Data amount: "
        self.worksheet["Z1"].font = Font(bold=True, size=14)
        self.worksheet["Z2"].value = self.data_amount

        self.save_weekday_data()
        self.style_excel()

        self.workbook.save(self.app.data_file)
        print("Excel customized.")


    def style_excel(self) -> None:
        self.worksheet["A1"].font = Font(bold=True, size=14)
        self.worksheet["B1"].font = Font(bold=True, size=14)
        self.worksheet["C1"].font = Font(bold=True, size=14)
        self.worksheet["D1"].font = Font(bold=True, size=14)
        self.worksheet["E1"].font = Font(bold=True, size=14)


    def get_start_time(self) -> None:
        self.start_time = datetime.datetime.now()


    def calculate_duration(self) -> float:
        duration = self.timer_manager.timer_time
        if duration < 0: 
            duration = 0
        else:
            duration /= 60
        return duration
    

    def increase_goal_streak(self) -> None:
        self.goal_amount += 1

    
    def reset_data(self, workbook, worksheet) -> None:
        self.workbook = workbook
        self.worksheet = worksheet

        self.initialize_new_file_variables()
        self.customize_excel()

        self.workbook.save(self.app.data_file)

        print("Data reset.")


    def clear_graph_lists(self) -> None:
        self.date_list.clear()
        self.duration_list.clear()


    def collect_day_data(self) -> None:
        self.monday_duration = int(self.worksheet["W2"].value)
        self.tuesday_duration = int(self.worksheet["W3"].value)
        self.wednesday_duration = int(self.worksheet["W4"].value)
        self.thursday_duration = int(self.worksheet["W5"].value)
        self.friday_duration = int(self.worksheet["W6"].value)
        self.saturday_duration = int(self.worksheet["W7"].value)
        self.sunday_duration = int(self.worksheet["W8"].value)

        self.day_duration_list = [self.monday_duration, self.tuesday_duration, self.wednesday_duration, self.thursday_duration, self.friday_duration, self.saturday_duration, self.sunday_duration]
    

    def set_weekday(self) -> None:
        duration = self.calculate_duration()
        weekday_today = datetime.datetime.now().weekday()

        match weekday_today:
            case 0:
                self.monday_duration += duration
            case 1:
                self.tuesday_duration += duration
            case 2:
                self.wednesday_duration += duration
            case 3:
                self.thursday_duration += duration
            case 4:
                self.friday_duration += duration
            case 5:
                self.saturday_duration += duration
            case 6:
                self.sunday_duration += duration

        self.save_weekday_data()

        self.workbook.save(self.app.data_file)
        print("Weekday saved.")


    def save_weekday_data(self) -> None:
        self.worksheet["W2"].value = self.monday_duration
        self.worksheet["W3"].value = self.tuesday_duration
        self.worksheet["W4"].value = self.wednesday_duration
        self.worksheet["W5"].value = self.thursday_duration
        self.worksheet["W6"].value = self.friday_duration
        self.worksheet["W7"].value = self.saturday_duration
        self.worksheet["W8"].value = self.sunday_duration


    def save_autobreak(self, frequency: str, duration: str, switch: str) -> None:
        self.worksheet["P2"].value = switch
        self.worksheet["P3"].value = int(frequency)
        self.worksheet["P4"].value = int(duration)

        self.workbook.save(self.app.data_file)

        print("Autobreak saved.")

        self.load_autobreak()
        

    def load_autobreak(self):
        self.autobreak_on = self.worksheet["P2"].value
        self.autobreak_frequency = self.worksheet["P3"].value
        self.autobreak_duration = self.worksheet["P4"].value

        if self.worksheet["P2"].value == None:
            self.autobreak_on = "Off"

        if self.autobreak_frequency == None:
            self.autobreak_frequency = 25

        if self.autobreak_duration == None:
            self.autobreak_duration = 5

        self.app.frequency_input.configure(placeholder_text=self.autobreak_frequency)
        self.app.duration_input.configure(placeholder_text=self.autobreak_duration)
        self.app.autobreak_switch.configure(variable=ctk.StringVar(value=self.autobreak_on))

        self.app.WINDOW.after(0, self.app.auto_break)

        print("Autobreak loaded.")


    def set_color(self, color_dropdown) -> None:
        self.color_name = color_dropdown.get()
        print("Color set.")
        self.save_color()


    def save_color(self) -> None:
        self.worksheet["T2"].value = self.color_name
        self.load_color()


    def load_color(self) -> None:
        self.color_name = self.worksheet["T2"].value
        self.app.color_dropdown.configure(variable=ctk.StringVar(value=self.color_name))
        colors = {"Orange": [orange_button_color, orange_highlight_color, orange_pie_colors], 
                    "Green": [green_button_color, green_highlight_color, green_pie_colors], 
                    "Blue": [blue_button_color, blue_highlight_color, blue_pie_colors],
                    "Pink": [pink_button_color, pink_highlight_color, pink_pie_colors]}
        
        self.color = colors[self.color_name][0]
        self.highlight_color = colors[self.color_name][1]
        self.pie_colors = colors[self.color_name][2]
        self.graph_color = self.color
        print("Color loaded.")
        self.change_color()


    def change_color(self) -> None:
        for widget in self.app.widget_list:
            widget.configure(fg_color=self.color, hover_color=self.highlight_color)
        self.app.progressbar.configure(progress_color = self.color)
        self.app.eye_care_checkbox.configure(fg_color=self.color)
        self.app.create_graphs()

        self.load_notes()
        print("Color changed.")


    def set_theme(self, theme_dropdown) -> None:
        self.theme_name = theme_dropdown.get()
        print("Theme set.")
        self.save_theme()


    def save_theme(self) -> None:
        self.worksheet["U2"].value = self.theme_name
        self.load_theme()

    
    def load_theme(self) -> None:
        self.theme_name = self.worksheet["U2"].value
        self.app.theme_dropdown.configure(variable=ctk.StringVar(value=self.theme_name))

        if self.theme_name == "Dark":
            ctk.set_appearance_mode("dark")
            self.graph_bg_color = graph_bg_color
            self.graph_fg_color = graph_fg_color
            self.spine_color = border_frame_color
            self.font_color = "white"
        else:
            ctk.set_appearance_mode("light")
            self.graph_bg_color = light_graph_bg_color
            self.graph_fg_color = light_graph_fg_color
            self.spine_color = light_border_frame_color
            self.font_color = "black"

        self.app.create_graphs()

        print("Theme loaded.")


    def save_subject(self, subject: str) -> None:
        self.worksheet["S2"].value = subject
        print("Subject saved.")


    def load_subject(self) -> None:
        if self.worksheet["S2"].value != None:
            subject = self.worksheet["S2"].value
        else:
            subject = "Other"

        self.app.subject_selection.configure(variable=ctk.StringVar(value=subject))

        return subject
    

    def create_new_note(self, title, text):
        self.notes_amount += 1
        
        self.worksheet["N9"].value = self.notes_amount

        self.worksheet["N" + str(self.notes_amount + 12)].value = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        self.worksheet["O" + str(self.notes_amount + 12)].value = title
        self.worksheet["P" + str(self.notes_amount + 12)].value = text

        self.workbook.save(self.app.data_file)

        print("New note created.")

        self.load_notes()


    def load_notes(self) -> None:
        self.notes_manager.clear_notes()

        if self.notes_amount == 0:
            return None
        
        for i in range(self.notes_amount+12, 12, -1):
            if self.worksheet["M" + str(i)].value != "Yes":
                self.notes_manager.create_task(i)


    def save_eye_care(self, eye_care: str, checkbox: str) -> None:
        self.worksheet["Q2"].value = eye_care
        self.worksheet["Q5"].value = checkbox

        self.workbook.save(self.app.data_file)
        print("Eye care saved.")


    def load_eye_care(self) -> None:
        if self.worksheet["Q2"].value != None:
            eye_care = self.worksheet["Q2"].value
        else:
            eye_care = "Off"

        self.app.eye_care_selection.configure(variable=ctk.StringVar(value=eye_care))

        if self.worksheet["Q5"].value != None:
            checkbox = self.worksheet["Q5"].value
        else:
            checkbox = "Off"

        self.app.eye_care_checkbox.configure(variable=ctk.StringVar(value=checkbox))

        self.app.WINDOW.after(0, self.app.eye_protection)  # Schedule initial iteration for eye_protection
