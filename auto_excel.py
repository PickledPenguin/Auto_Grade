import os
import pynput
import pandas as pd

CURRENT_DIR = os.getcwd()
EXCEL_DIR = "Excel"
EXCEL_FILENAMES = os.listdir("./%s" % EXCEL_DIR)
skip_rows = [1, 2, 3, 4, 5, 9, 10, 18]
require_cols = "C:G,I,J"


class ExcelData:
    def __init__(self, student_name, badge_num, timesheet):
        self.student_name = student_name
        self.badge_num = badge_num
        self.timesheet = timesheet


def read_excel_data(excel_filename):
    # change directory to the Excel file directory
    os.chdir(EXCEL_DIR)
    # return a panda dataframe with data from the given Excel file
    return pd.read_excel(excel_filename, skiprows=skip_rows)


if __name__ == "__main__":
    print(EXCEL_FILENAMES)
    for filename in range(len(EXCEL_FILENAMES)):
        print(read_excel_data(EXCEL_FILENAMES[filename]))
