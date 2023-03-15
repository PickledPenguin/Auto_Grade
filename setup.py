from pynput.mouse import Button, Controller
from pynput.keyboard import Key, Listener
from pynput import keyboard
import auto_excel

configuration = {}
config_counter = 0
mouse = Controller()


def on_release(key):
    global configuration, config_counter, mouse
    print('{0} released'.format(key))
    if key == keyboard.Key.esc:
        # Stop listener
        return False
    if hasattr(key, 'char'):
        # click waypoint
        if key.char == 'c':
            configuration[config_counter] = {"type": "click", "pos": mouse.position}

        # student name waypoint
        if key.char == 'n':
            configuration[config_counter] = {"type": "student-name", "pos": mouse.position}

        # timesheet waypoint
        if key.char == "t":
            configuration[config_counter] = {"type": "timesheet", "pos": mouse.position}

        # increment config counter
        config_counter += 1


def setup():
    global config_counter
    print("Now listening for waypoints. Hit esc to stop listening and end setup.")
    # Collect events until released
    with keyboard.Listener(on_release=on_release) as listener:
        listener.join()
    config_counter = 0
    return configuration
