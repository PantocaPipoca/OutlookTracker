####################--------Imports--------####################
import os
from dotenv import load_dotenv
from gui import create_gui

####################--------API Data--------####################
load_dotenv()

CLIENT_ID = os.environ.get("OutlookAutom_CLIENT")
TENANT_ID = "common"
CLIENT_SECRET = os.environ.get("OutlookAutom_SECRET")
SCOPE = ["https://graph.microsoft.com/calendars.readwrite"]

##################--------Main Function--------##################

def main():
    create_gui()

if __name__ == "__main__":
    main()