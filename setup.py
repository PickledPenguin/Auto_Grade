from pynput.mouse import Controller
from pynput import keyboard

configuration = {}
config_counter = 0
mouse = Controller()
wait_input = False


def on_release(key):
    global configuration, config_counter, mouse, wait_input
    print('{0} released'.format(key))
    if key == keyboard.Key.esc:
        # Stop listener
        return False
    if hasattr(key, 'char'):
        # if we are inputting the time for a wait waypoint
        if wait_input:
            configuration[config_counter-1]["seconds"] = int(key.char)
            wait_input = False
        else:
            # click waypoint
            if key.char == 'c':
                configuration[config_counter] = {"type": "click", "pos": mouse.position}

            # double click waypoint
            elif key.char == 'd':
                configuration[config_counter] = {"type": "double-click", "pos": mouse.position}

            # student name waypoint
            elif key.char == 'n':
                configuration[config_counter] = {"type": "student-name", "pos": mouse.position}

            # timesheet waypoint
            elif key.char == 't':
                configuration[config_counter] = {"type": "timesheet", "pos": mouse.position}

            elif key.char == 'w':
                configuration[config_counter] = {"type": "wait", "seconds": 1}
                wait_input = True

            else:
                # decrement to offset increment
                config_counter -= 1
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
