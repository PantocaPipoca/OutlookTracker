import re
from win10toast_click import ToastNotifier
import pyperclip
import webbrowser
import pygetwindow as gw
import psutil
import ctypes
import json
from datetime import datetime, timedelta

####################--------Application Dictionary--------####################

GENERAL_WINDOW_NAMES = {
    # Browsing (could be for studying, research, etc.)
    "msedge.exe": "Browsing",
    "chrome.exe": "Browsing",
    "firefox.exe": "Browsing",

    # Coding / Development
    "code.exe": "Coding",
    "pycharm.exe": "Coding",
    "sublime_text.exe": "Coding",
    "devenv.exe": "Coding",

    # Command Line / Scripting Tools (also for coding)
    "cmd.exe": "Coding",
    "powershell.exe": "Coding",
    "bash.exe": "Coding",
    "terminal.exe": "Coding",

    # Text Editors (for note-taking or studying)
    "notepad++.exe": "Studying",
    "notepad.exe": "Studying",
    "word.exe": "Studying",
    "winword.exe": "Studying",

    # Communication / Chatting
    "teams.exe": "Chatting",
    "slack.exe": "Chatting",
    "discord.exe": "Chatting",
    "skype.exe": "Chatting",
    "whatsapp.exe": "Chatting",
    "zoom.exe": "Chatting",

    # Email
    "outlook.exe": "Email",
    "thunderbird.exe": "Email",

    # Gaming
    "steam.exe": "Gaming",
    "epicgameslauncher.exe": "Gaming",
    "origin.exe": "Gaming",
    "battle.net.exe": "Gaming",

    # Photo Editing
    "photoshop.exe": "Photo Editing",
    "illustrator.exe": "Photo Editing",
    "gimp.exe": "Photo Editing",

    # Video Editing
    "Adobe Premiere Pro.exe": "Video Editing",
    "davinciresolve.exe": "Video Editing",
    "aftereffects.exe": "Video Editing",

    # Music Creation
    "ableton.exe": "Music Creation",
    "flstudio.exe": "Music Creation",

    # Music Listening
    "spotify.exe": "Music Listening",
    "vlc.exe": "Music Listening",

    # Personalization (customizing system appearance)
    "explorer.exe": "Personalizing",
    "rainmeter.exe": "Personalizing",
}

####################--------Helper Functions--------####################

def GeneralizeWindowName(processName, windowTitle):
    processName = processName.lower()

    if processName in GENERAL_WINDOW_NAMES:
        return GENERAL_WINDOW_NAMES[processName]
    return windowTitle

def SafePrint(message):
    try:
        print(message)
    except UnicodeEncodeError:
        print("Unable to print message due to encoding issues.")

def CleanString(string):
    return re.sub(r'[\u200b-\u200f]+|[^\x20-\x7E]', '', string)

def ShowAuthNotification(flow):
    notifier = ToastNotifier()

    def openAuthLink():
        webbrowser.open(flow['verification_uri'])

    notificationText = f"Authentication needed! Code: {flow['user_code']}"
    pyperclip.copy(flow['user_code'])

    notifier.show_toast("Outlook Automation",
                        notificationText,
                        duration=120,
                        threaded=True,
                        callback_on_click=openAuthLink)

def GetActiveWindow():
    user32 = ctypes.windll.user32
    pid = ctypes.c_ulong()
    window = gw.getActiveWindow()
    
    if window:
        try:
            hwnd = window._hWnd
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            processName = psutil.Process(pid.value).name()
            title = CleanString(window.title)
            generalizedName = GeneralizeWindowName(processName, title)
            return generalizedName, processName
        except psutil.NoSuchProcess:
            return window.title, "Process no longer exists"
    return None, None

##################--------Json Handler--------##################

def load_config():
    try:
        with open("config.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def save_config(config):
    with open("config.json", "w") as f:
        json.dump(config, f, indent=4)

def saveEventToDataFile(subject, startTime, endTime): 
    try:
        with open("data.json", "r") as f:
            data = json.load(f)
    except FileNotFoundError:
        data = []

    data.append({"subject": subject, "startTime": startTime, "endTime": endTime})

    with open("data.json", "w") as f:
        json.dump(data, f, indent=4)

if __name__ == "__main__":
    saveEventToDataFile("Test Event", datetime.now().isoformat(), (datetime.now() + timedelta(hours=1)).isoformat())