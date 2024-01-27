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
WINDOW.geometry(str(WIDTH + BORDER_WIDTH + main_frame_pad_x + tab_frame_width) + "x" + str(HEIGHT+((widget_padding_x+frame_padding)*2)))
WINDOW.title(APPNAME)
WINDOW.configure(background=window_color)
WINDOW.resizable(False, False)

main_frame = ctk.CTkFrame(WINDOW, fg_color=main_frame_color, height=HEIGHT+((widget_padding_x+frame_padding)*2), width=WIDTH, corner_radius=0)
main_frame.grid(column=2, row=0, padx=main_frame_pad_x)
main_frame.grid_propagate(False)

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
goal = ""

#------------------------------------------------------------------------------TIMER----------------------------------------------------------------------------#
def timer_mechanism():
    global timer_running, break_running, start_time
    global timer_btn, break_btn
    if not timer_running:
        timer_running = True
        break_running = False
        timer_btn.configure(text="Stop")
        break_btn.configure(text="Start")
        update_time()
    elif timer_running:
        timer_running = False
        timer_btn.configure(text="Start")
    if start_time == "":
        start_time = datetime.datetime.now()

def update_time():
    global timer_running, timer_time, time_display_label

    if timer_running:
        timer_time += 1
        time_display_label.configure(text=str(datetime.timedelta(seconds=timer_time)))
        update_slider()
        WINDOW.after(1000, update_time)

def break_mechanism():
    global break_running, timer_running
    global break_btn, timer_btn
    if not break_running:
        break_running = True
        timer_running = False
        break_btn.configure(text="Stop")
        timer_btn.configure(text="Start")
        update_break_time()
    elif break_running:
        break_running = False
        break_btn.configure(text="Start")

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
    main_frame.grid_propagate(False)
    statistics_frame.grid_forget()


def to_statistics():
    statistics_frame.grid(column=2, row=0, padx=main_frame_pad_x)
    main_frame.grid_forget()


def get_goal():
    global goal
    try:
        goal = int(goal_input.get())
        update_slider()
    except ValueError:
        print("Goal must be a number.")
        pass

def update_slider():
    global timer_time, progressbar, goal
    if goal != "":
        progressbar.set((timer_time/60)/goal)

#------------------------------------------------------------------------------GUI------------------------------------------------------------------------------#
def change_focus(event):
    event.widget.focus_set()

tab_frame = ctk.CTkFrame(WINDOW, width=tab_frame_width, height=HEIGHT+((widget_padding_x+frame_padding)*2), fg_color=tab_frame_color)
tab_frame.grid(column=0, row=0)
tab_frame.pack_propagate(False)

border_frame = ctk.CTkFrame(WINDOW, width=BORDER_WIDTH, height=HEIGHT+((widget_padding_x+frame_padding)*2), fg_color=border_frame_color)
border_frame.grid(column=1, row=0)

goal_progress_frame = ctk.CTkFrame(main_frame, height=(HEIGHT-button_height*1.5), width=frame_width)
goal_progress_frame.grid(row=0, column=0)
goal_progress_frame.pack_propagate(False)

goal_frame = ctk.CTkFrame(goal_progress_frame, fg_color=frame_color, height=225, width=frame_width, corner_radius=10)
goal_frame.pack(padx=frame_padding, pady=frame_padding)
goal_frame.pack_propagate(False)

goal_label = ctk.CTkLabel(goal_frame, text="Goal", font=(font_family, font_size), text_color=font_color)
goal_label.place(anchor="nw", relx=0.05, rely=0.05)
goal_input = ctk.CTkEntry(goal_frame, placeholder_text=30, font=(font_family, int(font_size*2.5)), text_color=font_color, height=70, width=90, justify="center")
goal_input.place(anchor="center", relx=0.5, rely=0.45)
goal_btn = ctk.CTkButton(goal_frame, text="Set", font=(font_family, int(font_size)), fg_color=button_color, text_color=button_font_color, width=80, height=40,
                          hover_color=button_highlight_color, command=get_goal)
goal_btn.place(anchor="s", relx=0.5, rely=0.9)

progress_frame = ctk.CTkFrame(goal_progress_frame, fg_color=frame_color, width=frame_width, corner_radius=10, height=90)
progress_frame.pack(padx=frame_padding, pady=frame_padding)
progress_frame.pack_propagate(False)
progress_label = ctk.CTkLabel(progress_frame, text="Progress", font=(font_family, int(font_size)), text_color=font_color)
progress_label.place(anchor="nw", relx=0.05, rely=0.05)
progressbar = ctk.CTkProgressBar(progress_frame, height=20, width=220, progress_color=button_color, fg_color=border_frame_color, corner_radius=10)
progressbar.place(anchor="center", relx=0.5, rely=0.7)
progressbar.set(0)

streak_frame = ctk.CTkFrame(goal_progress_frame, fg_color=frame_color, width=frame_width, corner_radius=10, height=120)
streak_frame.pack(padx=frame_padding, pady=frame_padding)
streak_label = ctk.CTkLabel(streak_frame, text="Streak", font=(font_family, int(font_size)), text_color=font_color)
streak_label.place(anchor="nw", relx=0.05, rely=0.05)


timer_break_frame = ctk.CTkFrame(main_frame, height=(HEIGHT-button_height*1.5), width=frame_width)
timer_break_frame.grid(row=0, column=1)
timer_break_frame.propagate(False)

timer_frame = ctk.CTkFrame(timer_break_frame, fg_color=frame_color, corner_radius=10, width=frame_width, height=220)
timer_frame.pack(padx=frame_padding, pady=frame_padding)
timer_frame.pack_propagate(False)

timer_label = ctk.CTkLabel(timer_frame, text="Timer", font=(font_family, font_size), text_color=font_color)
timer_label.place(anchor="nw", relx=0.05, rely=0.05)
time_display_label = ctk.CTkLabel(timer_frame, text="0:00:00", font=(font_family, int(font_size*3)), text_color=font_color)
time_display_label.place(anchor="center", relx=0.5, rely=0.45)
timer_btn = ctk.CTkButton(timer_frame, text="Start", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                 border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=timer_mechanism)
timer_btn.place(anchor="s", relx=0.5, rely=0.9)

#BREAK TIMER UI ROW
break_frame = ctk.CTkFrame(timer_break_frame, fg_color=frame_color, corner_radius=10, width=frame_width, height=220)
break_frame.pack(padx=frame_padding, pady=frame_padding)
break_frame.pack_propagate(False)

break_label = ctk.CTkLabel(break_frame, text="Break", font=(font_family, font_size), text_color=font_color)
break_label.place(anchor="nw", relx=0.05, rely=0.05)
break_display_label = ctk.CTkLabel(break_frame, text="0:00:00", font=(font_family, int(font_size*3)), text_color=font_color)
break_display_label.place(anchor="center", relx=0.5, rely=0.45)
break_btn = ctk.CTkButton(break_frame, text="Start", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                 border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=break_mechanism)
break_btn.place(anchor="s", relx=0.5, rely=0.9)

#DATA UI ROW
data_frame = ctk.CTkFrame(main_frame, fg_color=frame_color, corner_radius=10, width=WIDTH-10, height=button_height*2)
data_frame.place(anchor="s", relx=0.5, rely=0.985)
data_frame.grid_propagate(False)
save_data_btn = ctk.CTkButton(data_frame, text="Save Data", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                               border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=save_data, width=450)
save_data_btn.place(relx=0.01, anchor="w", rely=0.5)
reset_data_btn = ctk.CTkButton(data_frame, text="Reset Data", font=(font_family, font_size), fg_color=button_color, text_color=button_font_color,
                                border_color=frame_border_color, hover_color=button_highlight_color, height=button_height, command=reset_data, width=450)
reset_data_btn.place(relx=0.99, anchor="e", rely=0.5)

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

WINDOW.bind_all('<Button>', change_focus)

WINDOW.protocol("WM_DELETE_WINDOW", save_on_quit)

WINDOW.mainloop()