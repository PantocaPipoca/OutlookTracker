####################--------Imports--------####################
import os
import time
import requests
import msal
import ctypes
from datetime import datetime, timedelta
import HelperFunctions as hf
import psutil
import pygetwindow as gw

####################--------API Data--------####################

CLIENT_ID = os.environ.get("OutlookAutom_CLIENT")
TENANT_ID = "common"
CLIENT_SECRET = os.environ.get("OutlookAutom_SECRET")
SCOPE = ["https://graph.microsoft.com/calendars.readwrite"]

####################--------Event Handler--------####################

def AddEventToCalendar(token, userId, subject, startTime, endTime):
    adjustedStartTime = (datetime.fromisoformat(startTime) - timedelta(hours=1)).isoformat()
    adjustedEndTime = (datetime.fromisoformat(endTime) - timedelta(hours=1)).isoformat()

    url = f"https://graph.microsoft.com/v1.0/users/{userId}/events"
    headers = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json"
    }
    eventData = {
        "subject": subject,
        "start": {
            "dateTime": adjustedStartTime,
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": adjustedEndTime,
            "timeZone": "UTC"
        }
    }
    response = requests.post(url, headers=headers, json=eventData)
    hf.SafePrint("Event created successfully" if response.status_code == 201 else "Failed to create event, response code: " + str(response.status_code))

def GetActiveWindow():
    user32 = ctypes.windll.user32
    pid = ctypes.c_ulong()
    window = gw.getActiveWindow()
    
    if window:
        try:
            hwnd = window._hWnd
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            processName = psutil.Process(pid.value).name()
            title = hf.CleanString(window.title)
            generalizedName = hf.GeneralizeWindowName(processName, title)
            return generalizedName, processName
        except psutil.NoSuchProcess:
            return window.title, "Process no longer exists"
    return None, None

####################--------Usage Tracker--------####################

def TrackAndLogUsage(token: str, windowCheckDelay: int, eventUpdateDelay: int, intervalDuration: int):
    lastActiveWindow = None  # Last active window for a moment
    startTime = None  # Start time of the current window

    intervalStartTime = datetime.now()
    windowsActiveInXTime = {}
    mostActiveWindowInXTime = None

    combinedStartTime = None
    combinedEndTime = None

    while True:
        time.sleep(windowCheckDelay)  # Time between window checks
        currentActiveWindow, currentActiveProcess = GetActiveWindow()

        print("///////////////////////////////////////////////////")
        hf.SafePrint(f"Current Active Window: {currentActiveWindow}")
        hf.SafePrint(f"Last Active Window: {lastActiveWindow}")
        hf.SafePrint(f"Current Active Process: {currentActiveProcess}")
        print("///////////////////////////////////////////////////")

        if lastActiveWindow and currentActiveWindow == lastActiveWindow:
            duration = now - startTime
            if duration.total_seconds() >= eventUpdateDelay:
                if lastActiveWindow in windowsActiveInXTime:
                    windowsActiveInXTime[lastActiveWindow] += duration
                else:
                    windowsActiveInXTime[lastActiveWindow] = duration

        if currentActiveWindow != lastActiveWindow:  # If changed windows, last = current window and start time = current time
            if lastActiveWindow and startTime:  # If not first change, check if more than 60 seconds
                endTime = datetime.now()  # End time of current window
                duration = endTime - startTime
                if duration.total_seconds() > eventUpdateDelay:
                    if lastActiveWindow in windowsActiveInXTime:
                        windowsActiveInXTime[lastActiveWindow] += duration
                    else:
                        windowsActiveInXTime[lastActiveWindow] = duration
                    print(windowsActiveInXTime)

            lastActiveWindow = currentActiveWindow  # Reset variables right after window change
            startTime = datetime.now()
        
        now = datetime.now()
        if now - intervalStartTime >= intervalDuration:
            window_durations = {}
            for window, duration in windowsActiveInXTime.items():
                if window in window_durations:
                    window_durations[window] += duration
                else:
                    window_durations[window] = duration
        
            if window_durations:
                mostActiveWindow = max(window_durations, key=window_durations.get)
                hf.SafePrint(f"Most Active Window in {intervalDuration} minutes: {mostActiveWindow} ({window_durations[mostActiveWindow].total_seconds() / 60} minutes)")
                if mostActiveWindow != mostActiveWindowInXTime:
                    if combinedStartTime and combinedEndTime and mostActiveWindowInXTime:
                        AddEventToCalendar(token, USER_ID, mostActiveWindowInXTime, combinedStartTime.isoformat(), combinedEndTime.isoformat()) # type: ignore

                    combinedStartTime = intervalStartTime
                    combinedEndTime = now
                    mostActiveWindowInXTime = mostActiveWindow
                else:
                    combinedEndTime = now

            windowsActiveInXTime.clear()
            intervalStartTime = datetime.now()


####################--------API Access--------####################

def GetAccessToken():
    tokenCache = LoadTokenCache()  # Load saved token file
    app = msal.PublicClientApplication(CLIENT_ID, authority="https://login.microsoftonline.com/" + TENANT_ID, token_cache=tokenCache)  # Get application

    accounts = app.get_accounts()
    if accounts:
        token = app.acquire_token_silent(SCOPE, account=accounts[0])
        if "access_token" in token:
            SaveTokenCache(tokenCache)
            return token["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPE)
    if "user_code" not in flow:
        raise Exception("Failed to get device flow")

    hf.ShowAuthNotification(flow)
    print(f"Please go to {flow['verification_uri']} and enter the code: {flow['user_code']}")

    token = app.acquire_token_by_device_flow(flow)
    if "access_token" in token:
        SaveTokenCache(tokenCache)
        return token["access_token"]

    raise Exception("Failed to get access token")

def SaveTokenCache(tokenCache):
    if tokenCache.has_state_changed:
        with open("token_cache.json", "w") as f:
            f.write(tokenCache.serialize())

def LoadTokenCache():
    tokenCache = msal.SerializableTokenCache()
    if os.path.exists("token_cache.json"):
        with open("token_cache.json", "r") as f:
            cacheData = f.read()
            tokenCache.deserialize(cacheData)
    return tokenCache

####################--------Main Function--------####################

def start_tracking():
    token = GetAccessToken()
    windowCheckDelay = 1
    eventUpdateDelay = 1
    intervalDuration = timedelta(minutes=15)
    if timedelta(seconds=eventUpdateDelay) < intervalDuration:
        TrackAndLogUsage(token, windowCheckDelay, eventUpdateDelay, intervalDuration)