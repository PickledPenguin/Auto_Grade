from pynput.mouse import Button, Controller
from pynput.keyboard import Key, Listener
from pynput import keyboard

configuration = {}
config_counter = 0
mouse = Controller()
input_mode_flag = False


def on_release(key):
    global configuration, config_counter, mouse, input_mode_flag
    input_data = ""
    print('{0} released'.format(key))
    if key == keyboard.Key.esc:
        # Stop listener
        return False
    # if we are confirming an input variable
    if key == keyboard.Key.enter:
        input_mode_flag = False
        configuration["input{0}".format(config_counter)]["input"] = input_data
        input_data = ""
    if 'char' in dir(key):
        # if we are currently typing an input variable
        if input_mode_flag:
            input_data += key
        # if we are adding a waypoint
        else:
            # click waypoint
            if key.char == 'c':
                configuration["click{0}".format(config_counter)] = {"pos": mouse.position}
            # input waypoint
            if key.char == 'i':
                configuration["input{0}".format(config_counter)] = {"pos": mouse.position, "input": None}
                input_mode_flag = True
        # increment config counter
        config_counter += 1


def setup():
    global config_counter, input_mode_flag
    print("Now listening for waypoints. Hit esc to stop listening and end setup.")
    # Collect events until released
    with keyboard.Listener(on_release=on_release) as listener:
        listener.join()
    config_counter = 0
    input_mode_flag = False
    return configuration
