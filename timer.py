import datetime
import openpyxl as op
from openpyxl.styles import Font
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import customtkinter as ctk
from styles import *

APPNAME = "Timer App"
FILENAME = "Timer Data.xlsx"

local_folder = os.path.expandvars(rf"%APPDATA%\{APPNAME}")
data_file = os.path.expandvars(rf"%APPDATA%\{APPNAME}\{FILENAME}")

os.makedirs(local_folder, exist_ok=True)


WINDOW = ctk.CTk()
WINDOW.geometry(str(WIDTH + BORDER_WIDTH + main_frame_pad_x) + "x" + str(HEIGHT+((widget_padding_x+frame_padding)*2)))
WINDOW.title(APPNAME)
WINDOW.configure(background=window_color)

main_frame = ctk.CTkFrame(WINDOW, fg_color=main_frame_color)
main_frame.grid(column=2, row=0, padx=main_frame_pad_x)

statistics_frame = ctk.CTkFrame(WINDOW, fg_color=main_frame_color)
statistics_frame.grid(column=2, row=0, padx=main_frame_pad_x)
statistics_frame.grid_forget()

def customize_excel(worksheet):
    worksheet["A1"].value = "Start:"
    worksheet["B1"].value = "End:"
    worksheet["C1"].value = "Duration:"
    worksheet["D1"].value = "Break:"

    worksheet["A1"].font = Font(bold=True, size=14)
    worksheet["B1"].font = Font(bold=True, size=14)
    worksheet["C1"].font = Font(bold=True, size=14)
    worksheet["D1"].font = Font(bold=True, size=14)

    worksheet["X1"].value = "Data amount: "
    worksheet["X1"].font = Font(bold=True, size=14)
    worksheet["Z1"].value = data_amount
    workbook.save(data_file)
    print("Excel customized.")

date_list = []
duration_list = []


def create_graph(date_list, duration_list):
    data = {"Date": date_list, "Duration": duration_list}
    df = pd.DataFrame(data)
    grouped_data = df.groupby("Date")["Duration"].sum().reset_index()
    fig, ax = plt.subplots()
    ax.bar(grouped_data["Date"], grouped_data["Duration"], color=graph_color)
    ax.set_xlabel("Date", color=font_color)
    ax.set_ylabel("Duration in minutes", color=font_color)
    ax.set_title("Time Spent by Date", color=font_color)
    ax.tick_params(colors="white")
    ax.set_facecolor(graph_fg_color)
    ax.set_xticks(grouped_data["Date"])
    ax.set_yticks(grouped_data["Duration"])
    fig.set_facecolor(graph_bg_color)
    ax.spines["top"].set_color(spine_color)
    ax.spines["bottom"].set_color(spine_color)
    ax.spines["left"].set_color(spine_color)
    ax.spines["right"].set_color(spine_color)
    fig.set_size_inches(5, 4, forward=True)

    dateFormat = mdates.DateFormatter("%d/%m")
    ax.xaxis.set_major_formatter(dateFormat)
    graphFrame = FigureCanvasTkAgg(fig, master=statistics_frame)


    canvas_widget = graphFrame.get_tk_widget()
    canvas_widget.grid(row=4, column=0, padx=5, pady=10)
    canvas_widget.config(highlightbackground=frame_border_color, highlightthickness=2, background=frame_color)

    date_list.clear()
    duration_list.clear()


def collect_data():
    global data_amount, date_list, duration_list
    for data in range(2, data_amount + 2):
        if "/" in str(worksheet["B" + str(data)].value):
            date_list.append(datetime.datetime.strptime(str(worksheet["B" + str(data)].value).split(" ")[0], "%d/%m/%Y").date())
        elif "-" in str(worksheet["B" + str(data)].value):
            date_list.append(datetime.datetime.strptime(str(worksheet["B" + str(data)].value).split(" ")[0], "%Y-%m-%d").date())
        duration_list.append(round(worksheet["C" + str(data)].value))
    create_graph(date_list, duration_list)
    
if os.path.isfile(data_file):
    workbook = op.load_workbook(data_file)
    worksheet = workbook.active

    data_amount = int(worksheet["Z1"].value)

    collect_data()
    print("File loaded")
else:
    workbook = op.Workbook()
    worksheet = workbook.active

    data_amount = 0

    workbook.save(data_file)
    print("New file created")
    customize_excel(worksheet)


#------------------------------------------------------------------------------VARIABLES------------------------------------------------------------------------#
timer_running = False
break_running = False
timer_time = 0
break_time = 0
start_time = ""

#------------------------------------------------------------------------------TIMER----------------------------------------------------------------------------#
def timer_mechanism():
    global timer_running, break_running, start_time
    if not timer_running:
        timer_running = True
        break_running = False
        update_time()
    elif timer_running:
        timer_running = False

    if start_time == "":
        start_time = datetime.datetime.now()

def update_time():
    global timer_running, timer_time, time_display_label

    if timer_running:
        timer_time += 1
        time_display_label.configure(text=str(datetime.timedelta(seconds=timer_time)))
        WINDOW.after(1000, update_time)

def break_mechanism():
    global break_running, timer_running
    if not break_running:
        break_running = True
        timer_running = False
        update_break_time()
    elif break_running:
        break_running = False

def update_break_time():
    global break_running, break_time, break_display_label

    if break_running:
        break_time += 1
        break_display_label.configure(text=str(datetime.timedelta(seconds=break_time)))
        WINDOW.after(1000, update_break_time)


#------------------------------------------------------------------------------DATA-----------------------------------------------------------------------------#
def calculate_duration(timer_time, break_time):
    duration = timer_time - break_time
    if duration < 0: 
        duration = 0
    else:
        duration /= 60
    return duration

def save_data():
    global data_amount, duration_list, date_list
    global timer_running, timer_time, start_time
    global break_running, break_time

    if timer_time == 0:
        print("No data to save.")
        return
    timer_running, break_running = False, False

    duration = calculate_duration(timer_time, break_time)

    data_amount += 1
    worksheet["Z1"].value = int(data_amount)

    stop_time = datetime.datetime.now()

    worksheet["A" + str((data_amount + 1))].value = start_time.strftime("%d/%m/%Y %H:%M")
    worksheet["B" + str((data_amount + 1))].value = stop_time.strftime("%d/%m/%Y %H:%M")
    worksheet["C" + str((data_amount + 1))].value = duration
    worksheet["D" + str((data_amount + 1))].value = break_time/60

    timer_time, break_time = 0, 0
    start_time = ""
    workbook.save(data_file)
    print("Data saved.")
    collect_data()


def reset_data():
    global data_amount, duration_list, date_list
    global timer_time, timer_running, time_display_label
    global break_time, break_running, break_display_label

    data_amount = 0
    del workbook[workbook.active.title]
    workbook.create_sheet()
    worksheet = workbook.active

    timer_running, break_running = False, False
    timer_time, break_time = 0, 0
    time_display_label.configure(text="0:00:00")
    break_display_label.configure(text="0:00:00")

    worksheet["Z1"].value = int(data_amount)
    workbook.save(data_file)

    duration_list.clear()
    date_list.clear()

    print("Data reset.")
    create_graph(date_list, duration_list)
    customize_excel(worksheet)


def save_on_quit():
    global timer_time
    if timer_time > 0:
        save_data()
        print("Data saved on exit.")
    else: print("Quit.")
    workbook.save(data_file)
    WINDOW.destroy()

def to_timer():
    main_frame.grid(column=2, row=0, padx=main_frame_pad_x)
    statistics_frame.grid_forget()
def to_statistics():
    statistics_frame.grid(column=2, row=0, padx=main_frame_pad_x)
    main_frame.grid_forget()

#------------------------------------------------------------------------------GUI------------------------------------------------------------------------------#
tab_frame = ctk.CTkFrame(WINDOW, width=tab_frame_width, height=HEIGHT+((widget_padding_x+frame_padding)*2), fg_color=tab_frame_color)
tab_frame.grid(column=0, row=0)
tab_frame.pack_propagate(False)

border_frame = ctk.CTkFrame(WINDOW, width=BORDER_WIDTH, height=HEIGHT+((widget_padding_x+frame_padding)*2), fg_color=border_frame_color)
border_frame.grid(column=1, row=0)

timer_frame = ctk.CTkFrame(main_frame, border_color=frame_border_color, border_width=2, fg_color=frame_color)
timer_frame.grid(row=0, column=0, padx=frame_padding, pady=frame_padding)

timer_label = ctk.CTkLabel(timer_frame, text="Timer: ", font=(font_family, 16), text_color=font_color)
timer_label.grid(row=0, column=0, padx=(widget_padding_x*2, widget_padding_x), pady=widget_padding_y*2)
time_display_label = ctk.CTkLabel(timer_frame, text="0:00:00", font=(font_family, 16), text_color=font_color)
time_display_label.grid(row=0, column=2, padx=(widget_padding_x, widget_padding_x*2), pady=widget_padding_y)
timer_btn = ctk.CTkButton(timer_frame, text="Start/Stop", font=(font_family, 16), fg_color=button_color, text_color=button_font_color,
                                 border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=timer_mechanism)
timer_btn.grid(row=0, column=1, padx=widget_padding_x, pady=widget_padding_y)

#BREAK TIMER UI ROW
break_label = ctk.CTkLabel(timer_frame, text="Break: ", font=(font_family, 16), text_color=font_color)
break_label.grid(row=1, column=0, padx=(widget_padding_x*2, widget_padding_x), pady=widget_padding_y*2)
break_btn = ctk.CTkButton(timer_frame, text="Start/Stop", font=(font_family, 16), fg_color=button_color, text_color=button_font_color,
                                 border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=break_mechanism)
break_btn.grid(row=1, column=1, padx=widget_padding_x, pady=widget_padding_y)
break_display_label = ctk.CTkLabel(timer_frame, text="0:00:00", font=(font_family, 16), text_color=font_color)
break_display_label.grid(row=1, column=2, padx=(widget_padding_x, widget_padding_x*2), pady=widget_padding_y)

#DATA UI ROW
data_frame = ctk.CTkFrame(main_frame, border_color=frame_border_color, border_width=2, fg_color=frame_color)
data_frame.grid(row=3, column=0, padx=frame_padding, pady=frame_padding)
save_data_btn = ctk.CTkButton(data_frame, text="Save Data", font=(font_family, 16), fg_color=button_color, text_color=button_font_color,
                               border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=save_data)
save_data_btn.grid(row=0, column=0, padx=(widget_padding_x*2, widget_padding_x), pady=widget_padding_y*2)
reset_data_btn = ctk.CTkButton(data_frame, text="Reset Data", font=(font_family, 16), fg_color=button_color, text_color=button_font_color,
                                border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=reset_data)
reset_data_btn.grid(row=0, column=1, padx=(widget_padding_x, widget_padding_x*2), pady=widget_padding_y*2)

#TABS
timer_tab = ctk.CTkFrame(tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=tab_color)
timer_tab.pack(pady=tab_padding_y)
timer_tab_btn = ctk.CTkButton(timer_tab, text="Timer", font=(tab_font_family, 22*tab_height/60, tab_font_weight), text_color=font_color,
                                 fg_color=tab_color, width=int(tab_frame_width*0.95), height=int(tab_height*0.7), hover_color=tab_highlight_color, anchor="w", command=to_timer)
timer_tab_btn.place(relx=0.5, rely=0.5, anchor="center")

statistics_tab = ctk.CTkFrame(tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=tab_color)
statistics_tab.pack(pady=tab_padding_y)
statistics_btn = ctk.CTkButton(statistics_tab, text="Statistics", font=(tab_font_family, 22*tab_height/60, tab_font_weight), text_color=font_color,
                                 fg_color=tab_color, width=int(tab_frame_width*0.95), height=int(tab_height*0.7), hover_color=tab_highlight_color, anchor="w", command=to_statistics)
statistics_btn.place(relx=0.5, rely=0.5, anchor="center")

settings_tab = ctk.CTkFrame(tab_frame, width=tab_frame_width, height=tab_height*0.8, fg_color=tab_color)
settings_tab.place(relx=0.5, rely=1, anchor="s")
settings_btn = ctk.CTkButton(settings_tab, text="Settings", font=(tab_font_family, 22*tab_height/60, tab_font_weight), text_color=font_color,
                                 fg_color=tab_color, width=int(tab_frame_width*0.95), height=int(tab_height*0.7), hover_color=tab_highlight_color, anchor="w")
settings_btn.place(relx=0.5, rely=0.5, anchor="center")

WINDOW.protocol("WM_DELETE_WINDOW", save_on_quit)

WINDOW.mainloop()