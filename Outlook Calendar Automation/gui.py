####################--------Imports--------####################
import tkinter as tk
from tkinter import messagebox
from logic import start_tracking
from HelperFunctions import load_config, save_config
import threading

##################--------Get USER_ID--------##################

config = load_config()
USER_ID = config.get("USER_ID", "default_user_id@example.com")

def on_save_button_click(entry):
    global USER_ID
    USER_ID = entry.get()
    config["USER_ID"] = USER_ID
    save_config(config)
    messagebox.showinfo("Info", "User ID saved")

##################--------Start Tracking--------##################

def on_start_button_click():
    tracking_thread = threading.Thread(target=start_tracking)
    tracking_thread.daemon = True  # This ensures the thread will close when the main program exits
    tracking_thread.start()
    messagebox.showinfo("Info", "Tracking started")

##################--------Create GUI--------##################

def create_gui():
    root = tk.Tk()
    root.title("Outlook Calendar Automation")

    tk.Label(root, text="User ID:").pack(pady=5)
    user_id_entry = tk.Entry(root)
    user_id_entry.pack(pady=5)
    user_id_entry.insert(0, USER_ID)

    save_button = tk.Button(root, text="Save User ID", command=lambda: on_save_button_click(user_id_entry))
    save_button.pack(pady=5)

    start_button = tk.Button(root, text="Start Tracking", command=on_start_button_click)
    start_button.pack(pady=20)

    root.mainloop()