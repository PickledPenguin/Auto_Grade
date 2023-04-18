import sys

from pynput.mouse import Controller
from pynput import keyboard
from pynput.keyboard import Key, Listener
import os
import json

configuration = {}
config_counter = 0
mouse = Controller()
wait_input = False
paste_input = False


def optional_set_wait_time():
    """ Set the wait time between actions """

    wait_configuration = {"WAIT_DELAY_IN_SECONDS": 0.5}

    if input("Set wait time between actions? (y/n): ") == 'y':

        while True:
            wait_time = input("what is your preferred wait time (between 0.0 and 100.0) in seconds between actions? ("
                              "Press enter for default value of 0.5 seconds): ")
            try:
                if wait_time == '':
                    print("Wait time set to default value of 0.5 seconds")
                    wait_time = 0.5
                    break
                elif not (0.0 <= float(wait_time) <= 100.0):
                    print(f"Wait time must be between 0.0 and 100.0 seconds (you entered: {float(wait_time)} seconds)")
                    continue
                else:
                    print(f"Wait time set to {float(wait_time)} seconds")
                    break
            except (ValueError, NameError):
                print(f"Wait time must be a number between 0.0 and 100.0 (you entered: {float(wait_time)})")
                continue

        wait_configuration["WAIT_DELAY_IN_SECONDS"] = float(wait_time)

    return wait_configuration


def optional_set_datetime_format():
    """ Set the format for datetime type data """

    format_configuration = {"DATETIME_FORMAT_STR": "default"}

    if input("Set format string for datetime type data? (y/n): ") == 'y':

        format_string = input("Enter the format string for datetime type data (press enter for default string "
                              "conversion): ")
        # if enter was pressed
        if format_string == '':
            print("Format set to default string conversion")
            format_string = 'default'
        # if there is a format string
        else:
            print(f"Format set to {format_string}")

        format_configuration["DATETIME_FORMAT_STR"] = format_string

    return format_configuration


def on_release(key):
    """ Listen for waypoints, record mouse position at those waypoints, and return the information """

    global configuration, config_counter, mouse, wait_input, paste_input

    if key == keyboard.Key.esc:
        # Stop listener
        print("Esc pressed, stopped listening for waypoints")
        print("------------------------------------------------------")
        return False

    elif hasattr(key, 'char'):

        # click waypoint
        if key.char == 'c':
            print("\"c\" pressed - \"click\" waypoint at point [%f, %f]" % (mouse.position[0], mouse.position[1]))
            configuration[config_counter] = {"type": "click", "pos": mouse.position}

        # double click waypoint
        elif key.char == 'd':
            print("\"d\" pressed - \"double click\" waypoint at point [%f, %f]" % (
                mouse.position[0], mouse.position[1]))
            configuration[config_counter] = {"type": "double-click", "pos": mouse.position}

        # press key waypoint
        elif key.char == 'p':
            print("\"p\" pressed - \"paste\" waypoint at point")
            configuration[config_counter] = {"type": "paste"}
            paste_input = True
            print("------------------------------------------------------")
            print("Paused listening for waypoints, please input what you want pasted below")
            print("------------------------------------------------------")
            return False

        # tab waypoint
        elif key.char == 't':
            print("\"t\" pressed - \"tab\" waypoint at point [%f, %f]")
            configuration[config_counter] = {"type": "tab"}

        # enter waypoint
        elif key.char == 'e':
            print("\"e\" pressed - \"enter\" waypoint")
            configuration[config_counter] = {"type": "enter"}

        # feedback waypoint
        elif key.char == 'f':
            print("\"f\" pressed - \"feedback\" waypoint")
            configuration[config_counter] = {"type": "feedback"}

        # insert data waypoint
        elif key.char == 'g':
            print("\"g\" pressed - \"grade\" waypoint")
            configuration[config_counter] = {"type": "grade"}

        # insert student id waypoint
        elif key.char == 'i':
            print("\"i\" pressed - \"Student-ID\" waypoint")
            configuration[config_counter] = {"type": "Student-ID"}



        # wait waypoint
        elif key.char == 'w':
            print("\"w\" pressed - \"wait\" waypoint.")
            configuration[config_counter] = {"type": "wait", "seconds": 1}
            wait_input = True
            # stop listening
            print("------------------------------------------------------")
            print("Paused listening for waypoints, please input desired wait time below")
            print("------------------------------------------------------")
            return False

        else:
            # decrement to offset increment
            config_counter -= 1
        # increment config counter
        config_counter += 1


def resume_listening():
    with keyboard.Listener(on_release=on_release) as listener:
        listener.join()
    input_data()


def input_data():
    """listen for waypoints and gather user input for waypoints until the user hits esc to stop listening and
    finalize the config.json file"""

    global configuration, config_counter, wait_input, paste_input

    # if listening ended to collect input for a wait waypoint
    if wait_input:
        while True:
            # get input from the user
            wait_time = input("How many seconds would you like to wait for?: ")
            try:
                # if the wait time is unreasonably large
                if int(wait_time) > sys.maxsize:
                    print(
                        f"ERROR: you have set the \"wait\" waypoint for more than {sys.maxsize} seconds, which "
                        f"is FAR too long!")
                    continue
                # if the wait time is negative
                elif int(wait_time) < 0:
                    print(f"\"wait\" waypoint cannot be negative (you entered {int(wait_time)} seconds)")
                    continue
                # if the wait time is a valid number
                else:
                    print(f"\"wait\" waypoint set for {int(wait_time)} seconds")
                    configuration[config_counter]["seconds"] = int(wait_time)
                    config_counter += 1
                    # set flag variable to False
                    wait_input = False
                    # add a warning if the wait time is larger than an hour
                    if int(wait_time) >= 3600:
                        print(f"WARNING: you set the \"wait\" waypoint for {int(wait_time)} seconds, which is "
                              f"longer than an hour!")
                    print("------------------------------------------------------")
                    print("Resumed listening for waypoints")
                    print("------------------------------------------------------")
                    break
            # if the wait time is not a number
            except ValueError:
                print(f"\"wait\" waypoint must be set as a number (you entered {str(wait_time)})")
                continue

        resume_listening()

    # if listening was ended to collect input for a paste waypoint
    if paste_input:
        while True:
            # get input from the user
            paste_str = input("What would you like to paste?: ")
            print(f"\"paste\" waypoint set to \"{paste_str}\"")
            configuration[config_counter]["paste"] = paste_str
            config_counter += 1
            # set flag variable to False
            paste_input = False
            print("------------------------------------------------------")
            print("Resumed listening for waypoints")
            print("------------------------------------------------------")
            break

        resume_listening()


def config():
    """ Run configuration of wait time and waypoints and store collected data in the config.json file """

    global config_counter

    # set the wait time between actions
    configuration["WAIT_DELAY_IN_SECONDS"] = optional_set_wait_time()["WAIT_DELAY_IN_SECONDS"]
    configuration["DATETIME_FORMAT_STR"] = optional_set_datetime_format()["DATETIME_FORMAT_STR"]

    # reconfigure waypoints
    if input("Set waypoints? (y/n): ") == 'y':
        print("------------------------------------------------------")
        print("Now listening for waypoints.")
        # Collect events until released
        with keyboard.Listener(on_release=on_release) as listener:
            listener.join()
        input_data()

        config_counter = 0

    return configuration


def reset_config():
    """ Reset the wait time / waypoint configuration """

    filename = "config.json"

    # if the user wants to reset the configuration
    if input("Reconfigure / Overwrite program configuration? (y/n): ") == 'y':
        # RUN SETUP SCRIPT
        config_data = config()
        # remove the previous config.json file
        try:
            os.remove(filename)
            print("Replacing config.json file:")
        except FileNotFoundError:
            print("No config.json file to remove, continuing")
        with open(filename, 'a+') as f:
            # create new config.json file
            json.dump(config_data, f, indent=4)

        # get the new configuration from config.json
        with open(filename, 'r') as f:
            return json.load(f)
    # if the user does not want to reset configuration
    else:
        # get the configuration from config.json
        with open(filename, 'r') as f:

            return json.load(f)


if __name__ == "__main__":
    # Reset the configuration
    config_file = reset_config()
    print("\n------------------------------------------------------")
    print("Current config.json file:")
    print(config_file)
    print("------------------------------------------------------")
