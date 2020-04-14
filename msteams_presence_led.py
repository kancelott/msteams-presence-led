##############################################################################
#
# MS Teams Presence LED changer
#
# Parses MS Teams logs.txt file to get current presence status and fires off
# some action based off that.
#
# Easier than getting access to the Microsoft Graph API to get presence.
#
# Valid presence statuses I've found:
#   found status: Available
#   found status: Busy
#   found status: DoNotDisturb
#   found status: InAMeeting
#   found status: Presenting
#   found status: OnThePhone
#   found status: Away
#   found status: BeRightBack
#
#   found status: NewActivity
#   found status: ConnectionError
#   found status: Unknown
#
##############################################################################

import os
import re
import requests
import signal
import sys
import time
from datetime import datetime

TEAMS_LOG = os.path.join(os.getenv("APPDATA"), "Microsoft\\Teams\\logs.txt")
SLEEP_TIME = 30

ESPHOME_IP = "10.0.80.21"
ESPHOME_ID = "living_room_1"

DEBUG = False
#TEAMS_LOG = "logs.txt"


def main():
    oldStatus = ""

    # turn off light on CTRL+C
    signal.signal(signal.SIGINT, signalHandler)

    while True:
        status, logTime = getLatestStatus(TEAMS_LOG)
        now = datetime.now().strftime("%X")
        print("[" + now + "] MS Teams status: " + status + " (" + logTime + ")")

        if status == oldStatus:
            # no change -- do nothing
            pass

        elif status in ["Busy", "DoNotDisturb", "InAMeeting", "Presenting", "OnThePhone"]:
            # turn LED to RED
            print("  Turning light RED...")
            url = "http://" + ESPHOME_IP + "/light/" + ESPHOME_ID + "/turn_on?r=255&g=0&b=0&brightness=255"
            requests.post(url)

        elif status in ["Away", "BeRightBack"]:
            # turn LED to YELLOW
            print("  Turning light YELLOW...")
            url = "http://" + ESPHOME_IP + "/light/" + ESPHOME_ID + "/turn_on?r=255&g=170&b=0&brightness=255"
            requests.post(url)

        else:
            # turn LED to GREEN
            print("  Turning light GREEN...")
            url = "http://" + ESPHOME_IP + "/light/" + ESPHOME_ID + "/turn_on?r=0&g=255&b=0&brightness=255"
            requests.post(url)

        if DEBUG:
            # don't infinite loop on DEBUG mode
            break

        oldStatus = status
        time.sleep(SLEEP_TIME)


def signalHandler(sig, frame):
    """Turn off the light on CTRL+C
    """
    print("Exiting... turning off the light")
    url = "http://" + ESPHOME_IP + "/light/" + ESPHOME_ID + "/turn_off"
    requests.post(url)
    sys.exit(0)


def getLatestStatus(logfile):
    """Read and parse MS Teams logfile to get the last presence status.

    Returns:
        (status, logTime)
    """
    status = ""

    try:
        with open(logfile, "r") as f:
            for line in f:
                if "(current state: " not in line:
                    continue
                logTime = re.search("^(.+) (.+) (.+) (.+) (.+) GMT+", line).group(5)
                logStatus = re.search(".*\(current state: (.+) -> (.+)\)", line).group(2)

                if logStatus not in ["ConnectionError", "NewActivity", "Unknown"]:
                    status = logStatus

                if DEBUG:
                    #print(line.strip())
                    #print("  found date:" logTime")
                    #print("  found status:", logStatus")
                    #print("  set status:", status")
                    pass

        return (status, logTime)

    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()
