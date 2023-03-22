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
    """ for use after any meaningful action to give the website / other program time to catch up """
    time.sleep(WAIT_DELAY_IN_SECONDS)


def read_excel_data_to_numpy(excel_filename):
    """ Reads the given Excel file's data and returns the data in the form of a numpy array """

    # change directory to the Excel file directory
    os.chdir(EXCEL_DIR)
    # return a panda dataframe with data from the given Excel file
    excel_dataframe = pd.read_excel(excel_filename, sheet_name=sheet_name)
    os.chdir("..")
    result = excel_dataframe.to_numpy()
    print("------------------------------------------------------")
    print(f"Excel data for file \"{excel_filename}\":")
    print(result)
    print("------------------------------------------------------\n")
    return result


def execute_config(config_json, excel_data):
    """ Executes the waypoints given by the user for a specific Excel class """

    print(f"Executing waypoints")
    for i in config_json:
        # ignore the WAIT_DELAY_IN_SECONDS entry
        if i == "WAIT_DELAY_IN_SECONDS":
            continue
        # click
        elif config_json[i]["type"] == "click":
            print("Clicking at the point: [%f, %f]" % (config_json[i]["pos"][0], config_json[i]["pos"][1]))
            Mouse.position = config_json[i]["pos"]
            wait()
            Mouse.click(Button.left, 1)
        # double click
        elif config_json[i]["type"] == "double-click":
            print("Double clicking at the point: [%f, %f]" % (config_json[i]["pos"][0], config_json[i]["pos"][1]))
            Mouse.position = config_json[i]["pos"]
            wait()
            Mouse.click(Button.left, 2)
        # paste a string
        elif config_json[i]["type"] == "paste":
            print("pasting \"%s\"")
            kb.type(config_json[i]["paste"])
        # tab
        elif config_json[i]["type"] == "tab":
            print("pressing the tab key")
            kb.tap(Key.tab)
        # enter
        elif config_json[i]["type"] == "enter":
            print("pressing the enter key")
            kb.tap(Key.enter)
        # insert excel data
        elif config_json[i]["type"] == "insert-data":
            print("inserting excel data at [column %d, row %d]"
                  % (config_json[i]["excel_col"], config_json[i]["excel_row"]))
            # type the Excel data at the specified column and row
            kb.type(str(excel_data[config_json[i]["excel_col"] - 1][config_json[i]["excel_row"] - 1]))
        # wait
        elif config_json[i]["type"] == "wait":
            print("waiting for %d seconds" % config_json[i]["seconds"])
            time.sleep(config_json[i]["seconds"])

        wait()


def execute(data):
    """ Extracts the data from all Excel files in the EXCEL_FILENAMES directory and
    executes the waypoints given by the user stored in the config.json file """

    global WAIT_DELAY_IN_SECONDS

    # Set the wait delay to the user's set wait delay
    WAIT_DELAY_IN_SECONDS = data["WAIT_DELAY_IN_SECONDS"]

    for file in EXCEL_FILENAMES:
        execute_config(data, read_excel_data_to_numpy(file))
    print("\n Completed Excel data extraction and execution :)")


def optional_reset_config():
    """ Reset the wait time / waypoint configuration """

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
            json.dump(setup.config(), f, indent=4)

        # get the new configuration from config.json
        with open(filename, 'r') as f:
            return json.load(f)
    else:
        # get the configuration from config.json
        with open(filename, 'r') as f:
            return json.load(f)


def execution_countdown(data):
    """ Start a countdown until execution to allow the user to move mouse over to the desired screen """

    if input("Run program using current configuration (y/n): ") == 'y':
        print("Countdown to executing waypoints:")
        for i in range(5):
            print(5 - i)
            time.sleep(1)
        print("\nExtracting data from Excel files and beginning waypoint execution:\n")
        # execute the configuration
        execute(data)


if __name__ == "__main__":
    # Allow the user to change wait time / waypoint configuration if desired
    config_data = optional_reset_config()
    # Start countdown, then extract data from Excel files and execute the waypoints from the user's
    # desired configuration with the data from each Excel file.
    execution_countdown(config_data)
