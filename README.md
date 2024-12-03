# TaskbarUnfreezer

This small-ish program is designed to fix a very small yet annoying issue 
with the Windows taskbar: sometimes, it just freezes. Rather than fix it
myself every time, I decided to write a program to do it for me.

## How it works
This program uses Google's Tesseract OCR as well as the PILlow library to 
take a screenshot of the taskbar in set intervals and read the time from it; 
if the current time is different to the taskbar's time, it restarts the 
explorer.exe process (and hence, the taskbar). 

## Requirements

- Python 3.12+
  - Not tested on other versions, but probably viable
- Tesseract OCR
  - Must be added to PATH if on Windows

## Installation and Usage
Download the latest version of the python file and set it up to run from
startup. I like to use Windows Task Scheduler for this, but as long as 
the program is running, it should work.

The program should be run like this:
```
pythonw TaskbarUnfreezer.py
```
(pythonw runs the program without a console window)

## Config
The program has a few configuration options that can be changed in the
autogenerated `config.json` file. The options are as follows:

- `interval`: The interval in seconds between each check
- `logging_enabled`: Whether to output the program's actions to a log file
  - Enabled by default and set to clean out logs every day at midnight