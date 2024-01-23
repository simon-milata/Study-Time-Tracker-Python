import datetime
import openpyxl as op
from openpyxl.styles import Font
import tkinter as tk
import os
import atexit
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

desktop = os.path.expanduser("~/Desktop")
fileName = "Timer Data.xlsx"

if os.path.isfile(desktop + "\\" + fileName):
    wB = op.load_workbook(desktop + "\\" + fileName)
    wS = wB.active
    dataAmount = wS["Z1"].value
    print("File loaded")
else:
    dataAmount = 0
    wB = op.Workbook()
    wS = wB.active
    wS.append(["Start Time: ", "End Time: ", "Duration: ", "Break Duration: "])
    wS["X1"].value = "Data amount: "
    wS["A1"].font = Font(bold=True, size=14)
    wS["B1"].font = Font(bold=True, size=14)
    wS["C1"].font = Font(bold=True, size=14)
    wS["D1"].font = Font(bold=True, size=14)
    wS["X1"].font = Font(bold=True, size=14)
    print("New file created")
        
timerRunning = False

startTime = 0
stopTime = 0
duration = 0

def TimerStartStop():
    global timerRunning
    global startTime
    global stopTime
    global duration

    global breakRunning
    global breakStopTime
    global breakTimeTotal
    global breakStopLabel
    global breakStartLabel
    global breakTimeToDisplay

    global timerFrame
    global breakFrame

    if timerRunning == False:
        timerRunning = True
        timerFrame.config(background="OliveDrab2")
        startTime = datetime.datetime.now()
        timerStartLabel.config(text="Start: " + str(startTime.strftime("%H:%M:%S")))
        stopTime = 0
        timerStopLabel.config(text="Stop: 00:00:00")

        breakStartLabel.config(text="Start: 00:00:00")
        breakStopLabel.config(text="Stop: 00:00:00")
        print("Start:", startTime)
    else:
        timerRunning = False
        timerFrame.config(background="SystemButtonFace")
        stopTime = datetime.datetime.now()
        timerStopLabel.config(text="Stop: " + str(stopTime.strftime("%H:%M:%S")))
        print("Stop:", stopTime)
        if breakRunning == True:
            breakRunning = False
            breakFrame.config(background="SystemButtonFace")
            breakStopTime = datetime.datetime.now()
            print("Break Stop:", breakStopTime)
            breakStopLabel.config(text="Stop: " + str(breakStopTime.strftime("%H:%M:%S")))
            breakTimeTotal += ((breakStopTime.hour - breakStartTime.hour) * 60) + (breakStopTime.minute - breakStartTime.minute) + ((breakStopTime.second - breakStartTime.second) / 60)
            if breakTimeTotal > 1:
                breakTimeToDisplay = str(int(round(breakTimeTotal))) + "m"
            else:
                breakTimeToDisplay = str(int(breakTimeTotal*60)) + "s"

        duration += ((stopTime.hour - startTime.hour) * 60) + (stopTime.minute - stopTime.minute) + ((stopTime.second - startTime.second) / 60)

breakRunning = False

breakStartTime = 0
breakStopTime = 0
breakTimeTotal = 0
breakTimeToDisplay = 0

def BreakStartStop():
    global breakRunning
    global breakTimeTotal
    global breakStartTime
    global breakStopTime
    global breakTimeToDisplay

    global timerFrame
    global breakFrame

    if timerRunning == True:
        if breakRunning == False:
            breakRunning = True
            breakFrame.config(background="OliveDrab2")
            timerFrame.config(background="yellow2")
            breakStartTime = datetime.datetime.now()
            breakStartLabel.config(text="Start: " + str(breakStartTime.strftime("%H:%M:%S")))
            breakStopTime = 0
            breakStopLabel.config(text="Stop: 00:00:00")
            print("Break Start:", breakStartTime) 
        else:
            breakRunning = False
            breakFrame.config(background="SystemButtonFace")
            timerFrame.config(background="OliveDrab2")
            breakStopTime = datetime.datetime.now()
            breakStopLabel.config(text="Stop: " + str(breakStopTime.strftime("%H:%M:%S")))
            print("Break Stop:", breakStopTime)
            breakTimeTotal += ((breakStopTime.hour - breakStartTime.hour) * 60) + (breakStopTime.minute - breakStartTime.minute) + ((breakStopTime.second - breakStartTime.second) / 60)
            if breakTimeTotal > 1:
                breakTimeToDisplay = str(int(round(breakTimeTotal))) + "m"
            else:
                breakTimeToDisplay = str(int(breakTimeTotal*60)) + "s"
    else:
        print("Error: Cant break when timer not running")

def SaveData():
    global duration
    global breakTimeTotal
    global startTime
    global stopTime
    global breakTimeToDisplay
    global dataAmount
    global breakStartTime
    global breakStopTime
    
    if startTime != 0 and stopTime != 0:
        dataAmount += 1
        wS["Z1"].value = dataAmount
        totalTime = duration - breakTimeTotal
        if totalTime > 1:
            timeToDisplay = str(int(round(totalTime))) + "m"
        else:
            timeToDisplay = str(int(totalTime*60)) + "s"
        wS.append([startTime.strftime("%d/%m/%Y %H:%M"), stopTime.strftime("%d/%m/%Y %H:%M"), timeToDisplay, breakTimeToDisplay])
        wB.save(desktop + "\\" + fileName)
        print("Duration: ", timeToDisplay)
        print("Break duration: ", breakTimeToDisplay)
        print("Data saved")
        startTime = 0
        stopTime = 0
        breakStartTime = 0
        breakStopTime = 0
        timerStartLabel.config(text="Start: 00:00:00")
        timerStopLabel.config(text="Stop: 00:00:00")
        breakStartLabel.config(text="Start: 00:00:00")
        breakStopLabel.config(text="Stop: 00:00:00")
        timeToDisplay = 0
        duration = 0
        breakTimeTotal = 0
        CollectData()
    else:
        print("Error: No data to save")

def ResetData():
    global dataAmount
    global timerStartLabel
    global timerStopLabel
    global breakStartLabel
    global breakStopLabel
    timerStartLabel.config(text="Start: 00:00:00")
    timerStopLabel.config(text="Stop: 00:00:00")
    breakStartLabel.config(text="Start: 00:00:00")
    breakStopLabel.config(text="Stop: 00:00:00")
    dataAmount = 0
    del wB[wB.active.title]
    wB.create_sheet()
    wS = wB.active
    wS.append(["Start Time: ", "End Time: ", "Duration: ", "Break Duration: "])
    wS["X1"].value = "Data amount: "
    wS["A1"].font = Font(bold=True, size=14)
    wS["B1"].font = Font(bold=True, size=14)
    wS["C1"].font = Font(bold=True, size=14)
    wS["D1"].font = Font(bold=True, size=14)
    wS["X1"].font = Font(bold=True, size=14)
    wS["Z1"].value = dataAmount
    print("Data reset")
    wB.save(desktop + "\\" + fileName)

    fig, ax = plt.subplots()

    ax.bar(datetime.datetime.now().strftime("%d/%m/%Y"), 0)
    ax.set_xlabel("Date")
    ax.set_ylabel("Duration in minutes")
    ax.set_title("Time Spent by Date")

    graphFrame = FigureCanvasTkAgg(fig, master=window)
    canvas_widget = graphFrame.get_tk_widget()
    canvas_widget.grid(row=4, column=0, padx=5, pady=10)

dateList = []
durationList = []
def CollectData():
    global dataAmount
    print("Data amount: ", dataAmount)
    for data in range(2,dataAmount+2):
        dateList.append(str(wS["B"+str(data)].value).split(" ")[0])
        if "s" in wS["C"+str(data)].value:
            durationList.append(int(wS["C"+str(data)].value.replace("s", "")) / 60)
        else:
            durationList.append(int(wS["C"+str(data)].value.replace("m", "")))
    print("Date list: ", dateList)
    print("Duration list: ", durationList)
    data = {"Date": dateList, "Duration": durationList}
    df = pd.DataFrame(data)

    grouped_data = df.groupby("Date")["Duration"].sum().reset_index()

    fig, ax = plt.subplots()

    ax.bar(grouped_data["Date"], grouped_data["Duration"])
    ax.set_xlabel("Date")
    ax.set_ylabel("Duration in minutes")
    ax.set_title("Time Spent by Date")

    graphFrame = FigureCanvasTkAgg(fig, master=window)
    canvas_widget = graphFrame.get_tk_widget()
    canvas_widget.grid(row=4, column=0, padx=5, pady=10)


    dateList.clear()
    durationList.clear()

window = tk.Tk()
window.geometry("650x765")
window.title("Timer")

# TIMER

timerFrame = tk.Frame(window, highlightbackground="black", highlightthickness=2)
timerFrame.grid(row=0, column=0, padx=10, pady=10)
timerLabel = tk.Label(timerFrame, text="Timer: ", font="Calibri 16")
timerLabel.grid(row=0, column=0, padx=5, pady=10)
timerStartBtn = tk.Button(timerFrame, text="Start/Stop", command=TimerStartStop, font="Calibri 16")
timerStartBtn.grid(row=0, column=1, padx=5, pady=10)

timerStartLabel = tk.Label(timerFrame, text="Start: 00:00:00", font="Calibri 16")
timerStartLabel.grid(row=0, column=2, padx=5, pady=10)
timerStopLabel = tk.Label(timerFrame, text="Stop: 00:00:00", font="Calibri 16")
timerStopLabel.grid(row=0, column=3, padx=5, pady=10)

# BREAK

breakFrame = tk.Frame(window, highlightbackground="black", highlightthickness=2)
breakFrame.grid(row=1, column=0, padx=10, pady=10)
breakLabel = tk.Label(breakFrame, text="Break: ", font="Calibri 16")
breakLabel.grid(row=0, column=0, padx=5, pady=10)
breakStartBtn = tk.Button(breakFrame, text="Start/Stop", command=BreakStartStop, font="Calibri 16")
breakStartBtn.grid(row=0, column=1, padx=5, pady=10)

breakStartLabel = tk.Label(breakFrame, text="Start: 00:00:00", font="Calibri 16")
breakStartLabel.grid(row=0, column=2, padx=5, pady=10)
breakStopLabel = tk.Label(breakFrame, text="Stop: 00:00:00", font="Calibri 16")
breakStopLabel.grid(row=0, column=3, padx=5, pady=10)

# DATA

dataFrame = tk.Frame(window, highlightbackground="black", highlightthickness=2)
dataFrame.grid(row=3, column=0, padx=10, pady=10)
saveDataBtn = tk.Button(dataFrame, text="Save Data", command=SaveData, font="Calibri 16")
saveDataBtn.grid(row=0, column=0, padx=5, pady=10)
resetDataBtn = tk.Button(dataFrame, text="Reset Data", command=ResetData, font="Calibri 16")
resetDataBtn.grid(row=0, column=1, padx=5, pady=10)


fig, ax = plt.subplots()

ax.bar(datetime.datetime.now().strftime("%d/%m/%Y"), 0)
ax.set_xlabel("Date")
ax.set_ylabel("Duration in minutes")
ax.set_title("Time Spent by Date")

graphFrame = FigureCanvasTkAgg(fig, master=window)
canvas_widget = graphFrame.get_tk_widget()
canvas_widget.grid(row=4, column=0, padx=5, pady=10)

window.mainloop()