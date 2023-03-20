from pynput.mouse import Controller
from pynput import keyboard
import math

configuration = {}
config_counter = 0
mouse = Controller()
wait_input = False


def optional_set_wait_time():
    global configuration
    while True:
        wait_time = input("what is your preferred wait time (between 0.0 and 10.0) in seconds between actions? (Press "
                          "enter for default value of 0.5 seconds): ")
        try:
            if wait_time == '':
                print("Wait time set to default value of 0.5 seconds")
                wait_time = 0.5
                break
            elif not (0.0 <= float(wait_time) <= 10.0):
                print(f"Wait time must be between 0.0 and 10.0 seconds (you entered: {float(wait_time)} seconds)")
                continue
            else:
                print(f"Wait time set to {float(wait_time)} seconds")
                break
        except (ValueError, NameError):
            print(f"Wait time must be a number between 0.0 and 10.0 (you entered: {float(wait_time)})")
            continue

    configuration["WAIT_DELAY_IN_SECONDS"] = float(wait_time)
    return configuration


def on_release(key):
    global configuration, config_counter, mouse, wait_input

    if key == keyboard.Key.esc:
        # Stop listener
        print("Esc pressed, stopped listening for waypoints")
        print("------------------------------------------------------")
        return False
    if hasattr(key, 'char'):
        # if we are inputting the time for a wait waypoint
        try:
            if wait_input:
                print(f"\"wait\" waypoint set for {int(key.char)} seconds")
                configuration[config_counter - 1]["seconds"] = int(key.char)
                wait_input = False
        except ValueError:
            print(f"\"wait\" waypoint must be between 0 and 9 (you entered {key.char}). Canceled \"wait\" waypoint")
        else:
            # click waypoint
            if key.char == 'c':
                print("\"c\" pressed - \"click\" waypoint at point [%f, %f]" % (mouse.position[0], mouse.position[1]))
                configuration[config_counter] = {"type": "click", "pos": mouse.position}

            # double click waypoint
            elif key.char == 'd':
                print("\"d\" pressed - \"double click\" waypoint at point [%f, %f]" % (mouse.position[0], mouse.position[1]))
                configuration[config_counter] = {"type": "double-click", "pos": mouse.position}

            # student name waypoint
            elif key.char == 'n':
                print("\"n\" pressed - \"name\" waypoint at point [%f, %f]" % (mouse.position[0], mouse.position[1]))
                configuration[config_counter] = {"type": "student-name", "pos": mouse.position}

            # timesheet waypoint
            elif key.char == 't':
                print("\"t\" pressed - \"timesheet\" waypoint starting at point [%f, %f]" % (mouse.position[0], mouse.position[1]))
                configuration[config_counter] = {"type": "timesheet", "pos": mouse.position}

            elif key.char == 'w':
                print("\"w\" pressed - \"wait\" waypoint. How long in seconds (between 0 and 9) would you like to wait "
                      "for?: ")
                configuration[config_counter] = {"type": "wait", "seconds": 1}
                wait_input = True

            else:
                # decrement to offset increment
                config_counter -= 1
            # increment config counter
            config_counter += 1


def setup():
    global configuration, config_counter

    # set the wait time between actions
    configuration["WAIT_DELAY_IN_SECONDS"] = optional_set_wait_time()["WAIT_DELAY_IN_SECONDS"]

    # reconfigure waypoints
    if input("Reconfigure waypoints? (y/n): ") == 'y':
        print("------------------------------------------------------")
        print("Now listening for waypoints. Hit esc to end listening.")
        # Collect events until released
        with keyboard.Listener(on_release=on_release) as listener:
            listener.join()
        config_counter = 0
        return configuration
    else:
        return None
