# Description

This is a script created for the RIT SE department to help automate the process of inputting the student employees' hours from Excel spreadsheets into an external website. Previously, this had to be done manually, cell-by-cell, line-by-line, which was tedious and unnecessary.


# How it works

Once configured correctly, this script will extract data from all student Excel spreadsheets in a directory and input the gathered data into the external website through simulating moving the mouse and keypresses via the python package pynput. 


# Architecture Reasoning

This script does not determine where to input the Excel data from the external website itself, but rather relies on a manual setup process in which the user directs the program to the points of interest (waypoints) on the monitor. There are two main reasons for this approach: First, I was unable to access the external website while developing this script as it was restricted and only accessible by members of the SE department. Second, in the case that the SE department decides to use a different external website or the original external website is changed or updated, this script would still function, and it would only need to be reconfigured rather than redesigned. 
This flexibility comes with several drawbacks: The initial setup process is more complex than I would like and the script is dependent on the screen being completely static and consistent every single time the script is run. I believe these drawbacks are worth the flexibility, at least for such a small-scale project.


# Required Installations

1.) python3
2.) pip or pip3

To install all the project requirments, type the following bash command in the project directory:
"pip install -r requirements.txt" 
or 
"pip3 install -r requirements.txt"


# Setting wait time

Whenever you run the script you are given the option to set the wait time, which is the time (in seconds) between any meaningful action the script simulates. This gives the website time to catch up and helps make sure that no inputs or data is lost. If your connection is a little unstable, I would recommend increasing this wait time. The default wait time is 0.5 seconds.


# Setting up "waypoints"

Whenever you run the script you are given the option to reconfigure the waypoints, which are points of interest that you indicate on the screen for the program to either click, input data, or perform some other action. Currently, there are 5 different types of waypoints: Click, Double click, Tab, Enter, Paste, Insert, and Wait.
Each type of waypoint has a character (a key on a keyboard) associated with it. To create a waypoint, move your mouse to the point on the screen where you want the waypoint to be executed, then hit the key on your keyboard that matches the appropriate waypoint action. Some waypoints require additional data which you can input in the python window. Below is a list of all the current waypoint actions and the characters / keys associated with them:

'c' = Click (Click the left mouse button at the point on your screen where your mouse is hovering)

'd' = Double click (Double-click the left mouse button at the point on your screen where your mouse is hovering)

't' = Tab (Press the tab button)

'e' = Enter (Press the enter button)

'p' = Paste (Type out a specified text)
After hitting this key, return to the python window to input the desired text. After this is complete, the script will resume listening for other waypoints.

'i' = Insert Data (Insert / type data in a specified column and row of the current Excel file)
After hitting this key, return to the python window to input the desired column and row. After this is complete, the script will resume listening for other waypoints.

'w'= Wait (Wait for a specified number of seconds). 
After hitting this key, return to the python window to input the desired number of seconds. After this is complete, the script will resume listening for other waypoints.

Once you are finished, hit the 'esc' key to tell the script to stop listening for waypoints and create the config.json file.


# Running the Script

First, ensure that all the desired Excel files are stored in the "Excel" directory / folder.
Then, set up your preferred wait time and waypoint configuration (See "Setting wait time" and "Setting up 'waypoints'")
Next, type the following bash command in the project directory:
"python3 ./auto_excel.py"
A countdown will begin, which gives you time to complete the final step: Moving your mouse cursor to the screen / window / website where you want to execute the Excel data.
That's it! If everything is set up correctly, the script will now automatically extract data from the given Excel files and follow the waypoints you set up to input the data into the website!