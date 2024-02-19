import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator, FuncFormatter
import customtkinter as ctk

from .styles import *

class StatisticsManager:
    def __init__(self, App: object, WINDOW: object):
        self.app = App
        self.WINDOW = WINDOW

        self.statistics_gui = StatisticsGUI(self.app, self.WINDOW)

        self.hide_statistics_gui()


    def show_statistics_gui(self) -> None:
        self.statistics_gui.statistics_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.statistics_gui.statistics_frame.grid_propagate(False)

    
    def hide_statistics_gui(self) -> None:
        self.statistics_gui.statistics_frame.grid_forget()


    def create_graphs(self) -> None:
        self.statistics_gui.create_graphs()


class StatisticsGUI:
    def __init__(self, App, WINDOW):
        self.app = App
        self.WINDOW = WINDOW
        self._initialize_main_frame()


    def _initialize_main_frame(self) -> None:
        self.statistics_frame = ctk.CTkFrame(self.WINDOW, fg_color=(light_main_frame_color, main_frame_color), height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
        self.statistics_frame.grid(column=2, row=0, padx=main_frame_pad_x)
        self.statistics_frame.grid_propagate(False)

    
    def create_graphs(self) -> None:
        self.create_time_spent_graph()
        self.create_weekday_graph()


    def create_time_spent_graph(self) -> None:
        data = {"Date": self.app.data_manager.date_list, "Duration": self.app.data_manager.duration_list}
        df = pd.DataFrame(data)
        grouped_data = df.groupby("Date")["Duration"].sum().reset_index()
        fig1, ax = plt.subplots()
        ax.bar(grouped_data["Date"], grouped_data["Duration"], color=self.app.data_manager.graph_color)
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
        self.app.data_manager.collect_day_data()

        if self.app.data_manager.day_duration_list:
            non_zero_durations = [duration for duration in self.app.data_manager.day_duration_list if duration != 0]
            non_zero_names = [name for name, duration in zip(self.app.data_manager.day_name_list, self.app.data_manager.day_duration_list) if duration != 0]

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
        ax.pie(non_zero_durations, labels=non_zero_names, autopct=_autopct_format(non_zero_durations), colors=self.app.data_manager.pie_colors, 
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