import json
import logging
import os
import re
import sys
from logging.handlers import TimedRotatingFileHandler
import win32gui
import pyautogui  # this import fixes the DPI issues with win32gui, do not remove
from PIL import ImageGrab
import screeninfo
import pytesseract
import time
import datetime

TESTING_MODE = False

formatter = logging.Formatter(
    fmt='[%(asctime)s - %(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logging.basicConfig(
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S',
    format='[%(asctime)s - %(levelname)s] %(message)s'
)
logger = logging.getLogger('taskbarunfreezer')
log_handler = TimedRotatingFileHandler('taskbarunfreezer.log', when='midnight', backupCount=1)
log_handler.setFormatter(formatter)
logger.addHandler(log_handler)

# if running in background, redirect stdout and stderr to /dev/null
# to avoid errors when no cmd window is open
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w')


def main():
    check_config_exists()  # creates config if none exists (usually first run)

    hwnd, bbox = get_taskbar_window()
    if bbox is None:
        return

    while True:
        # check for config changes
        config = handle_config()
        # check time from screenshot
        logger.info(bbox) if config['logging_enabled'] else None
        try:
            screenshot = ImageGrab.grab(bbox=bbox)

            ocr_str = pytesseract.image_to_string(screenshot)
            logger.info(ocr_str) if config['logging_enabled'] else None

            # get time from screenshot, including AM/PM
            time_str = re.search(r'\d+:\d+ [AP]M', ocr_str)

            # convert time_str to unix timestamp
            curr_date = datetime.datetime.now().strftime('%Y-%m-%d')
        except Exception as e:
            logger.error(e)
            logger.error('Failed to grab/manipulate screenshot')

        try:
            time_str = time_str.group(0)
            time_obj = datetime.datetime.strptime(curr_date + ' ' + time_str, '%Y-%m-%d %I:%M %p')
            curr_time = datetime.datetime.now()

            # compare time_str with current time and see if they are 2 mins or more apart
            if datetime.timedelta(minutes=2) < curr_time - time_obj:
                # restart taskbar process
                logger.info('Restarting explorer.exe') if config['logging_enabled'] else None
                os.system("taskkill /f /im explorer.exe")
                os.system("start explorer.exe")
        except Exception as e:
            # most likely, taskbar is not visible (e.g. fullscreen app is running)
            # handle exception and continue
            logger.error(e)
            logger.error('Time not found (most likely, taskbar is not visible)')

        # 2 min sleep
        if not TESTING_MODE:
            time.sleep(config['delay'])
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


def check_config_exists():
    # check if config file exists
    if os.path.exists('taskbarunfreezer_config.json'):
        return
    else:
        # if not, create config file
        print('Creating config file')
        default_data = {
            'delay': 120,
            'logging_enabled': True
        }
        with open('taskbarunfreezer_config.json', 'w') as f:
            json.dump(default_data, f)


def handle_config():
    # handle config
    with open('taskbarunfreezer_config.json', 'r') as f:
        config = json.load(f)
        return config

if __name__ == "__main__":
    main()
