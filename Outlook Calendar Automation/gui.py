####################--------Imports--------####################
import tkinter as tk
from tkinter import messagebox, ttk
from logic import startTracking
from HelperFunctions import loadConfig, saveConfig, getUserId
import threading

isTracking = threading.Event()
trackingActive = False
trackingThread = None

def startTrackingButton(userId):
    global trackingActive, trackingThread
    saveConfig({"USER_ID": userId})
    if trackingActive:
        statusVar.set("Status: Tracking Already Started")
        return
    if trackingThread is None or not trackingThread.is_alive():
        isTracking.clear()
        trackingThread = threading.Thread(target=lambda: startTracking(isTracking))
        trackingThread.daemon = True
        trackingThread.start()
        trackingActive = True
        statusVar.set("Status: Tracking Started")

def stopTrackingButton():
    global trackingActive
    if not trackingActive:
        statusVar.set("Status: Tracking Already Stopped")
        return
    isTracking.set()
    trackingActive = False
    statusVar.set("Status: Tracking Stopped")

##################--------Create GUI--------##################

def createGUI():
    root = tk.Tk()
    root.title("Activity Tracker")

    user_id_label = ttk.Label(root, text="User ID:")
    user_id_label.grid(row=0, column=0, padx=5, pady=5)

    user_id_entry = ttk.Entry(root, width=30)
    user_id_entry.insert(0, getUserId())
    user_id_entry.grid(row=0, column=1, padx=5, pady=5)

    start_button = ttk.Button(root, text="Start", command=lambda: startTrackingButton(user_id_entry.get()))
    start_button.grid(row=0, column=2, padx=5, pady=5)

    global statusVar
    statusVar = tk.StringVar()
    statusVar.set("Status: Inactive")
    status_label = ttk.Label(root, textvariable=statusVar)
    status_label.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

    stop_button = ttk.Button(root, text="Stop", command=stopTrackingButton)
    stop_button.grid(row=2, column=2, padx=5, pady=5)

    root.mainloop()

if __name__ == "__main__":
    createGUI()