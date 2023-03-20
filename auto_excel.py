import os
import pynput.keyboard
from pynput.mouse import Button, Controller
from pynput.keyboard import Key, Listener, Controller
import pandas as pd
import time
import json
import setup

CURRENT_DIR = os.getcwd()
EXCEL_DIR = "Excel"
EXCEL_FILENAMES = os.listdir(f"./{EXCEL_DIR}")
WAIT_DELAY_IN_SECONDS = 0.5
sheet_name = "Time Sheet"

Mouse = pynput.mouse.Controller()
kb = pynput.keyboard.Controller()


def wait():
    time.sleep(WAIT_DELAY_IN_SECONDS)


class ExcelData:
    def __init__(self, faculty_name, course, student_name, badge_num, timesheet):
        self.faculty_name = faculty_name
        self.course = course
        self.student_name = student_name
        self.badge_num = badge_num
        self.timesheet = timesheet


def read_excel_data_to_numpy(excel_filename):
    # change directory to the Excel file directory
    os.chdir(EXCEL_DIR)
    # return a panda dataframe with data from the given Excel file
    excel_dataframe = pd.read_excel(excel_filename, sheet_name=sheet_name)
    os.chdir("..")
    return excel_dataframe.to_numpy()


def convert_datetime(datetime):
    if isinstance(datetime, str):
        return ""

    hour = datetime.hour
    if datetime.hour < 12:
        is_time_am = True
    elif datetime.hour == 12:
        is_time_am = False
    else:
        hour -= 12
        is_time_am = False
    minute = datetime.minute
    return f"{hour}:{minute}A" if is_time_am else f"{hour}:{minute}A"


def convert_numpy_to_excel_class(filename, excel_array):
    print(f"\nExtracting data from Excel file \"{filename}\":")
    student_name = excel_array[4, 2]
    reformatted_name_list = student_name.split(" ")
    reformatted_name = reformatted_name_list[1] + ', ' + reformatted_name_list[0]
    faculty_name = excel_array[5, 2]
    course = excel_array[6, 2]
    badge_num = int(excel_array[4, 9])
    timesheet = {
        "Fri": {"date": excel_array[9, 2],
                "time_in_1": convert_datetime(excel_array[9, 3]),
                "time_out_1": convert_datetime(excel_array[9, 4]),
                "time_in_2": convert_datetime(excel_array[9, 5]),
                "time_out_2": convert_datetime(excel_array[9, 6]),
                "comments": excel_array[9, 8]},
        "Sat": {"date": excel_array[10, 2],
                "time_in_1": convert_datetime(excel_array[10, 3]),
                "time_out_1": convert_datetime(excel_array[10, 4]),
                "time_in_2": convert_datetime(excel_array[10, 5]),
                "time_out_2": convert_datetime(excel_array[10, 6]),
                "comments": excel_array[10, 8]},
        "Sun": {"date": excel_array[11, 2],
                "time_in_1": convert_datetime(excel_array[11, 3]),
                "time_out_1": convert_datetime(excel_array[11, 4]),
                "time_in_2": convert_datetime(excel_array[11, 5]),
                "time_out_2": convert_datetime(excel_array[11, 6]),
                "comments": excel_array[11, 8]},
        "Mon": {"date": excel_array[12, 2],
                "time_in_1": convert_datetime(excel_array[12, 3]),
                "time_out_1": convert_datetime(excel_array[12, 4]),
                "time_in_2": convert_datetime(excel_array[12, 5]),
                "time_out_2": convert_datetime(excel_array[12, 6]),
                "comments": excel_array[12, 8]},
        "Tues": {"date": excel_array[13, 2],
                 "time_in_1": convert_datetime(excel_array[13, 3]),
                 "time_out_1": convert_datetime(excel_array[13, 4]),
                 "time_in_2": convert_datetime(excel_array[13, 5]),
                 "time_out_2": convert_datetime(excel_array[13, 6]),
                 "comments": excel_array[13, 7]},
        "Wed": {"date": excel_array[14, 2],
                "time_in_1": convert_datetime(excel_array[14, 3]),
                "time_out_1": convert_datetime(excel_array[14, 4]),
                "time_in_2": convert_datetime(excel_array[14, 5]),
                "time_out_2": convert_datetime(excel_array[14, 6]),
                "comments": excel_array[14, 8]},
        "Thurs": {"date": excel_array[15, 2],
                  "time_in_1": convert_datetime(excel_array[15, 3]),
                  "time_out_1": convert_datetime(excel_array[15, 4]),
                  "time_in_2": convert_datetime(excel_array[15, 5]),
                  "time_out_2": convert_datetime(excel_array[15, 6]),
                  "comments": excel_array[15, 8]}
    }

    excel_data = ExcelData(faculty_name, course, reformatted_name, badge_num, timesheet)

    print(f"Student Name: {excel_data.student_name}")
    print(f"Faculty Name: {excel_data.faculty_name}")
    print(f"Course: {excel_data.course}")
    print(f"Badge Number: {excel_data.badge_num}")
    print(f"Timesheet: {excel_data.timesheet}\n")

    return excel_data


def execute_timesheet(excel_class):
    def tab():
        kb.tap(Key.tab)
        wait()

    kb.type(excel_class.timesheet["Fri"]["time_in_1"])
    tab()
    kb.type(excel_class.timesheet["Fri"]["time_out_1"])
    tab()
    kb.type(excel_class.timesheet["Fri"]["time_in_2"])
    tab()
    kb.type(excel_class.timesheet["Fri"]["time_out_2"])
    tab()
    tab()
    kb.type(excel_class.timesheet["Sat"]["time_in_1"])
    tab()
    kb.type(excel_class.timesheet["Sat"]["time_out_1"])
    tab()
    kb.type(excel_class.timesheet["Sat"]["time_in_2"])
    tab()
    kb.type(excel_class.timesheet["Sat"]["time_out_2"])
    tab()
    tab()
    kb.type(excel_class.timesheet["Sun"]["time_in_1"])
    tab()
    kb.type(excel_class.timesheet["Sun"]["time_out_1"])
    tab()
    kb.type(excel_class.timesheet["Sun"]["time_in_2"])
    tab()
    kb.type(excel_class.timesheet["Sun"]["time_out_2"])
    tab()
    tab()
    kb.type(excel_class.timesheet["Mon"]["time_in_1"])
    tab()
    kb.type(excel_class.timesheet["Mon"]["time_out_1"])
    tab()
    kb.type(excel_class.timesheet["Mon"]["time_in_2"])
    tab()
    kb.type(excel_class.timesheet["Mon"]["time_out_2"])
    tab()
    tab()
    kb.type(excel_class.timesheet["Tues"]["time_in_1"])
    tab()
    kb.type(excel_class.timesheet["Tues"]["time_out_1"])
    tab()
    kb.type(excel_class.timesheet["Tues"]["time_in_2"])
    tab()
    kb.type(excel_class.timesheet["Tues"]["time_out_2"])
    tab()
    tab()
    kb.type(excel_class.timesheet["Wed"]["time_in_1"])
    tab()
    kb.type(excel_class.timesheet["Wed"]["time_out_1"])
    tab()
    kb.type(excel_class.timesheet["Wed"]["time_in_2"])
    tab()
    kb.type(excel_class.timesheet["Wed"]["time_out_2"])
    tab()
    tab()
    kb.type(excel_class.timesheet["Thurs"]["time_in_1"])
    tab()
    kb.type(excel_class.timesheet["Thurs"]["time_out_1"])
    tab()
    kb.type(excel_class.timesheet["Thurs"]["time_in_2"])
    tab()
    kb.type(excel_class.timesheet["Thurs"]["time_out_2"])
    wait()
    kb.tap(Key.enter)


def execute_config(config_json, excel_class):
    print(f"Executing waypoints for student \"{excel_class.student_name}\" with badge number {excel_class.badge_num}:")
    for i in config_json:
        # ignore the WAIT_DELAY_IN_SECONDS entry
        if i == "WAIT_DELAY_IN_SECONDS":
            continue
        # click
        if config_json[i]["type"] == "click":
            print("Clicking at the point: [%f, %f]" % (config_json[i]["pos"][0], config_json[i]["pos"][1]))
            Mouse.position = config_json[i]["pos"]
            wait()
            Mouse.click(Button.left, 1)
        # double click
        if config_json[i]["type"] == "double-click":
            print("Double clicking at the point: [%f, %f]" % (config_json[i]["pos"][0], config_json[i]["pos"][1]))
            Mouse.position = config_json[i]["pos"]
            wait()
            Mouse.click(Button.left, 2)
        # type student name
        if config_json[i]["type"] == "student-name":
            print(f"Typing student name (%s) at the point: [%f, %f]" % (excel_class.student_name, config_json[i]["pos"][0], config_json[i]["pos"][1]))
            Mouse.position = config_json[i]["pos"]
            wait()
            Mouse.click(Button.left, 1)
            wait()
            kb.type(excel_class.student_name)
            kb.tap(Key.enter)
        # enter timesheet
        if config_json[i]["type"] == "timesheet":
            print("Entering student timesheet starting at the point: [%f, %f]" % (config_json[i]["pos"][0], config_json[i]["pos"][1]))
            Mouse.position = config_json[i]["pos"]
            wait()
            Mouse.click(Button.left, 1)
            wait()
            execute_timesheet(excel_class)
        # wait
        if config_json[i]["type"] == "wait":
            print("waiting for %d seconds" % config_json[i]["seconds"])
            time.sleep(config_json[i]["seconds"])
        wait()


def execute(data):
    for file in EXCEL_FILENAMES:
        execute_config(data, convert_numpy_to_excel_class(file, read_excel_data_to_numpy(file)))


def optional_reset_config():
    filename = "config.json"
    if input("Reset program configuration? (y/n): ") == 'y':
        # remove the previous config.json file
        try:
            os.remove(filename)
            print("Replacing config.json file:")
        except FileNotFoundError:
            print("No config.json file to remove, continuing")

        # RUN SETUP SCRIPT
        # Store configuration in config.json
        with open(filename, 'a+') as f:
            json.dump(setup.setup(), f, indent=4)

        # get the new configuration from config.json
        with open(filename, 'r') as f:
            return json.load(f)
    else:
        # get the configuration from config.json
        with open(filename, 'r') as f:
            return json.load(f)


def run_config(data):
    if input("Run program using current configuration (y/n): ") == 'y':
        print("Countdown to executing waypoints:")
        for i in range(5):
            print(5-i)
            time.sleep(1)
        print("\nExtracting data from Excel files and beginning waypoint execution:\n")
        # execute the configuration
        execute(data)


if __name__ == "__main__":
    config_data = optional_reset_config()
    run_config(config_data)
