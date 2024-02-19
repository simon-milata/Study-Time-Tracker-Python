import os
import random
import time
from threading import Thread

import openpyxl as op
import customtkinter as ctk
from winotify import Notification

from Package import *
from Package.main_manager import MainManager
from Package.statistics_management import StatisticsManager


APPNAME = "Timer App"
FILENAME = "Timer Data.xlsx"


class App:
    def __init__(self) -> None:
        self.WINDOW = None
        self.workbook = None
        self.worksheet = None
        self.main_manager = None
        self.data_manager = None
        self.timer_manager = None

        self._initialize_variables()
        self._initialize_window()
        self._initialize_file()
        self._initialize_menu()
        self._initialize_managers()

        self.statistics_manager.create_graphs()


    def _initialize_variables(self):
        self.notification_limit_on = False


    def _initialize_window(self) -> None:
        self.WINDOW = ctk.CTk()
        self.WINDOW.geometry(str(WIDTH + main_frame_pad_x + tab_frame_width) + "x" + str(HEIGHT+((widget_padding_x+frame_padding)*2)))
        self.WINDOW.title(APPNAME)
        self.WINDOW.configure(fg_color=(light_window_color, window_color))
        self.WINDOW.resizable(False, False)
        self.WINDOW.grid_propagate(False)
        self.WINDOW.protocol("WM_DELETE_WINDOW", self._save_on_quit)


    def _initialize_file(self) -> None:
        self.local_folder = os.path.expandvars(rf"%APPDATA%\{APPNAME}")
        self.data_file = os.path.expandvars(rf"%APPDATA%\{APPNAME}\{FILENAME}")

        os.makedirs(self.local_folder, exist_ok=True)

        data_file_exists = os.path.isfile(self.data_file)

        if data_file_exists:
            self._load_existing_file()
            print("File loaded.")

        else:
            self._create_new_file()
            print("New file created.")


    def _load_existing_file(self) -> None:
        self.workbook = op.load_workbook(self.data_file)
        self.worksheet = self.workbook.active

        #self.data_manager = DataManager(self, self.timer_manager, self.workbook, self.worksheet)

        #self.create_widget_list()
        #self.collect_data()
        #self.update_streak_values()
        #self.load_history()

        #self.data_manager.load_color()
        #self.data_manager.load_theme()
        #self.data_manager.load_notes()


    def _create_new_file(self) -> None:
        self.workbook = op.Workbook()
        self.worksheet = self.workbook.active

        self.workbook.save(self.data_file)

        #self.data_manager = DataManager(self, self.timer_manager, self.workbook, self.worksheet)

        #self.create_widget_list()

        #self.data_manager.initialize_new_file_variables()

        #self.data_manager.save_eye_care("Off", "Off")

    
    def _initialize_managers(self) -> None:
        self.timer_manager = TimerManager(self, self.WINDOW)
        self.main_manager = MainManager(self, self.WINDOW, self.timer_manager)
        self.statistics_manager = StatisticsManager(self, self.WINDOW)
        self.data_manager = DataManager(self, self.timer_manager, self.statistics_manager, self.workbook, self.worksheet)


    def _initialize_menu(self) -> None:
        self.tab_frame = ctk.CTkFrame(self.WINDOW, width=tab_frame_width, height=HEIGHT+((widget_padding_x+frame_padding)*2), fg_color=(light_tab_frame_color, tab_frame_color))
        self.tab_frame.grid(column=0, row=0)
        self.tab_frame.pack_propagate(False)

        self._initialize_tabs()


    def _initialize_tabs(self) -> None:

        tab_buttons = []

        def _initialize_tab(tab_name: str) -> None:
            #Create a tab frame and a tab button for each tab in tab_list
            tab = ctk.CTkFrame(self.tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=(light_tab_color, tab_color))
            tab_button = ctk.CTkButton(tab, text=tab_name, font=(tab_font_family, 22*tab_height/50, tab_font_weight), text_color=(light_font_color, font_color),
                                                   fg_color=(light_tab_color, tab_color), width=int(tab_frame_width*0.95), height=int(tab_height*0.8), hover_color=(light_tab_highlight_color, tab_highlight_color), 
                                                   anchor="w", command=lambda: self.switch_tab(tab_button, tab_buttons))
            
            #Place settings tab frame on the bottom of tab frame
            if tab_name == "Settings":
                tab.place(relx=0.5, rely=1, anchor="s")
            else:
                tab.pack(pady=tab_padding_y)
            tab_button.place(relx=0.5, rely=0.5, anchor="center")

            #Append each button to tab_buttons list for future manipulation
            tab_buttons.append(tab_button)

        tab_list = ["Main", "Statistics", "Achievements", "History", "Notes", "Settings"]

        for tab in tab_list:
            _initialize_tab(tab)

        #Change the color of the timer tab button by default as it's selected
        tab_buttons[0].configure(fg_color=(light_tab_selected_color, tab_selected_color), hover=False)

    
    def switch_tab(self, selected_tab_button: ctk.CTkButton, tab_buttons: list[ctk.CTkButton]) -> None:
        """
        Switches the active tab and updates tab button colors accordingly.
        """

        #Change the color of all buttons to unselected color

        def _forget_tabs():
            self.main_manager.hide_main_gui()
            self.statistics_manager.hide_statistics_gui()

        _forget_tabs()


        def _show_selected_tab():
            match selected_tab_button.cget("text"):
                case "Main":
                    self.main_manager.show_main_gui()

                case "Statistics":
                    self.statistics_manager.show_statistics_gui()

        _show_selected_tab()    


        def _decolor_tabs() -> None:
            for button in tab_buttons:
                button.configure(fg_color=(light_tab_color, tab_color), hover_color=(light_tab_highlight_color, tab_highlight_color))

        _decolor_tabs()

        #Change color of pressed button to selected color
        selected_tab_button.configure(fg_color=(light_tab_selected_color, tab_selected_color), hover=False)


    def send_notification(self, title, message) -> None:
        toast = Notification(app_id=APPNAME, title=title, msg=message)
        toast.show()
        self.notification_limit_on = True
        print("Notification " + title + " sent.")


    def _save_on_quit(self) -> None:
        #self.save_data()
        print("Data saved on exit.")

        self.workbook.save(self.data_file)
        self.WINDOW.destroy()


    def _run(self) -> None:
        self.WINDOW.mainloop()


if __name__ == "__main__":
    App()._run()
