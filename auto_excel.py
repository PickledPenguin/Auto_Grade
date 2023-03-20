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
EXCEL_FILENAMES = os.listdir("./{Excel_Dir}".format(Excel_Dir=EXCEL_DIR))
WAIT_DELAY_IN_SECONDS = 0.5
sheet_name = "Time Sheet"

Mouse = pynput.mouse.Controller()
kb = pynput.keyboard.Controller()


class ExcelData:
    def __init__(self, faculty_name, course, student_name, badge_num, timesheet):
        self.faculty_name = faculty_name
        self.course = course
        self.student_name = student_name
        self.badge_num = badge_num
        self.timesheet = timesheet


def read_excel_data_to_numpy(excel_filename):
    print(excel_filename)
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


def convert_df_to_excel_class(excel_array):
    student_name = excel_array[4, 2]
    reformatted_name_list = student_name.split(" ")
    reformatted_name = reformatted_name_list[1] + ', ' + reformatted_name_list[0]
    faculty_name = excel_array[5, 2]
    course = excel_array[6, 2]
    badge_num = excel_array[4, 9]
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

    print(excel_data.student_name)
    print(excel_data.faculty_name)
    print(excel_data.course)
    print(excel_data.badge_num)
    print(excel_data.timesheet)

    return excel_data


def execute_timesheet(excel_class):
    
    def tab():
        kb.tap(Key.tab)
        time.sleep(WAIT_DELAY_IN_SECONDS)
    
    print("execute timesheet reached")
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
    time.sleep(WAIT_DELAY_IN_SECONDS)


def execute_config(config_json, excel_class):
    for i in config_json:
        if config_json[i]["type"] == "click":
            print("clicking at the following point:")
            print(config_json[i]["pos"])
            Mouse.position = config_json[i]["pos"]
            Mouse.click(Button.left, 1)
        if config_json[i]["type"] == "double-click":
            print("double clicking at the following point:")
            print(config_json[i]["pos"])
            Mouse.position = config_json[i]["pos"]
            Mouse.click(Button.left, 2)
        if config_json[i]["type"] == "student-name":
            Mouse.position = config_json[i]["pos"]
            print(f"typing name: {excel_class.student_name}")
            kb.type(excel_class.student_name)
            kb.tap(Key.enter)
        if config_json[i]["type"] == "timesheet":
            Mouse.position = config_json[i]["pos"]
            Mouse.click(Button.left, 1)
            execute_timesheet(excel_class)
        if config_json[i]["type"] == "wait":
            print("waiting for %d seconds" % config_json[i]["seconds"])
            time.sleep(config_json[i]["seconds"])

        time.sleep(WAIT_DELAY_IN_SECONDS)


def execute(file):
    for i in EXCEL_FILENAMES:
        execute_config(file, convert_df_to_excel_class(read_excel_data_to_numpy(i)))


if __name__ == "__main__":
    filename = "config.json"
    if input("Reset the configuration? (y/n)") == 'y':
        # remove the previous config.json file
        try:
            os.remove(filename)
        except FileNotFoundError:
            print("No config.json file to remove, continuing")
        with open(filename, 'a+') as f:
            json.dump(setup.setup(), f, indent=4)
        with open(filename, 'r') as f:
            data = json.load(f)
            print("data:")
            print(data)
            execute(data)
    else:
        with open(filename, 'r') as f:
            data = json.load(f)
            print("data:")
            print(data)
            execute(data)
