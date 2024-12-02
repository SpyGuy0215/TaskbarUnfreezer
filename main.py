import os
import re
import sys
import win32gui
import pyautogui  # this import fixes the DPI issues with win32gui, do not remove
from PIL import ImageGrab
import screeninfo
import pytesseract
import time
import datetime

TESTING_MODE = False

# if running in background, redirect stdout and stderr to /dev/null
# to avoid errors when no cmd window is open
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w')


def main():
    hwnd, bbox = get_taskbar_window()
    if bbox is None:
        return

    while True:
        # check time from screenshot
        print('running it back')
        print(bbox)
        screenshot = ImageGrab.grab(bbox=bbox)
        ocr_str = pytesseract.image_to_string(screenshot)
        print(ocr_str)

        # get time from screenshot, including AM/PM
        time_str = re.search(r'\d+:\d+ [AP]M', ocr_str)

        # convert time_str to unix timestamp
        curr_date = datetime.datetime.now().strftime('%Y-%m-%d')
        try:
            time_str = time_str.group(0)
            time_obj = datetime.datetime.strptime(curr_date + ' ' + time_str, '%Y-%m-%d %I:%M %p')
            curr_time = datetime.datetime.now()

            # compare time_str with current time and see if they are 2 mins or more apart
            if datetime.timedelta(minutes=2) < curr_time - time_obj:

                # restart taskbar process
                os.system("taskkill /f /im explorer.exe")
                os.system("start explorer.exe")
        except Exception as e:
            # most likely, taskbar is not visible (e.g. fullscreen app is running)
            # handle exception and continue
            print(e)
            print('Time String not found')

        # 2 min sleep
        if not TESTING_MODE:
            time.sleep(120)
        else:
            # only run once if in testing mode
            break


def get_taskbar_window():
    try:
        # get window handel of taskbar and get bounding box
        hwnd = win32gui.FindWindow("Shell_traywnd", None)
        win32gui.SetForegroundWindow(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        bbox = (screeninfo.get_monitors()[0].width / 2, rect[1] + 30, rect[2], rect[3])
        return hwnd, bbox
    except Exception as e:
        return None


if __name__ == "__main__":
    main()
