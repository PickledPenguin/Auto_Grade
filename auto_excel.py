import os

import pynput.keyboard
from pynput.mouse import Button, Controller
from pynput.keyboard import Key, Controller
from pynput.keyboard import Key, Listener
from pynput import keyboard
import pandas as pd
import setup
import json

CURRENT_DIR = os.getcwd()
EXCEL_DIR = "Excel"
EXCEL_FILENAMES = os.listdir("./{Excel_Dir}".format(Excel_Dir=EXCEL_DIR))
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
    # change directory to the Excel file directory
    os.chdir(EXCEL_DIR)
    # return a panda dataframe with data from the given Excel file
    excel_dataframe = pd.read_excel(excel_filename, sheet_name=sheet_name)
    os.chdir("..")
    return excel_dataframe.to_numpy()


def convert_df_to_excel_class(excel_array):
    student_name = excel_array[4, 2]
    faculty_name = excel_array[5, 2]
    course = excel_array[6, 2]
    badge_num = excel_array[4, 9]
    timesheet = {
        "Fri": {"date": excel_array[9, 2],
                "time_in_1": excel_array[9, 3],
                "time_out_1": excel_array[9, 4],
                "time_in_2": excel_array[9, 5],
                "time_out_2": excel_array[9, 6],
                "comments": excel_array[9, 8]},
        "Sat": {"date": excel_array[10, 2],
                "time_in_1": excel_array[10, 3],
                "time_out_1": excel_array[10, 4],
                "time_in_2": excel_array[10, 5],
                "time_out_2": excel_array[10, 6],
                "comments": excel_array[10, 8]},
        "Sun": {"date": excel_array[11, 2],
                "time_in_1": excel_array[11, 3],
                "time_out_1": excel_array[11, 4],
                "time_in_2": excel_array[11, 5],
                "time_out_2": excel_array[11, 6],
                "comments": excel_array[11, 8]},
        "Mon": {"date": excel_array[12, 2],
                "time_in_1": excel_array[12, 3],
                "time_out_1": excel_array[12, 4],
                "time_in_2": excel_array[12, 5],
                "time_out_2": excel_array[12, 6],
                "comments": excel_array[12, 8]},
        "Tues": {"date": excel_array[13, 2],
                 "time_in_1": excel_array[13, 3],
                 "time_out_1": excel_array[13, 4],
                 "time_in_2": excel_array[13, 5],
                 "time_out_2": excel_array[13, 6],
                 "comments": excel_array[13, 7]},
        "Wed": {"date": excel_array[14, 2],
                "time_in_1": excel_array[14, 3],
                "time_out_1": excel_array[14, 4],
                "time_in_2": excel_array[14, 5],
                "time_out_2": excel_array[14, 6],
                "comments": excel_array[14, 8]},
        "Thurs": {"date": excel_array[15, 2],
                  "time_in_1": excel_array[15, 3],
                  "time_out_1": excel_array[15, 4],
                  "time_in_2": excel_array[15, 5],
                  "time_out_2": excel_array[15, 6],
                  "comments": excel_array[15, 8]}
    }

    return ExcelData(student_name, faculty_name, course, badge_num, timesheet)


def execute_config(config_file, excel_class):
    for curr_config in config_file:
        if curr_config["type"] == "click":
            Mouse.position = curr_config["pos"]
            Mouse.click(Button.left, 1)
        if curr_config["type"] == "double-click":
            Mouse.position = curr_config["pos"]
            Mouse.click(Button.left, 2)
        if curr_config["type"] == "student-name":
            Mouse.position = curr_config["pos"]
            kb.type(excel_class.student_name)
        if curr_config["type"] == "timesheet":
            # TODO type out timesheet using tab
            continue


if __name__ == "__main__":
    if input("begin setup? (y/n)") == 'y':
        filename = "config.json"
        # remove the previous config.json te
        os.remove(filename)
        with open(filename, 'w') as f:
            json.dump(setup.setup(), f, indent=4)
            for i in range(len(EXCEL_FILENAMES)):
                execute_config(f, convert_df_to_excel_class(read_excel_data_to_numpy(EXCEL_FILENAMES[i])))
