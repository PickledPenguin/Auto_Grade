# Auto_Excel

This is a script created for the RIT SE department to help automate the process of inputting the student employees' hours from Excel spreadsheets into an external website. Previously, this had to be done mantually, cell-by-cell, line-by-line, which was tedious and unnessessary.
Once the setup is completed correctly, this script will read data from all student Excel spreadsheets in a directory and input them into the external website through simulating mouse movements and keypresses via pynput. 

# Architecture Reasoning

This script does not determine where to input the Excel data from the external website, but rather relies on a manual setup process in which the user directs the program to the points of interest on the monitor. There are two main reasons for this: First, I was unable to access the external website while developing this script as it was restricted and only accessable by members of the SE department. Second, in the case that the SE department decides to use a different external website or the origional external website is changed or updated, this script would still function and it would need to be reconfigured rather than redesigned. 
This flexibility comes with several drawbacks: The initial setup process is long, the script is dependant on the screen being completely static and consistant every single time the script is run, and it can be imprecise. I believe these drawbacks are worth the flexibility, at least for such a small-scale project.

# Setup
Run the following two bash commands in the project directory:
"pip install -r requirements.txt"
"chmod -x auto_excel.py"
