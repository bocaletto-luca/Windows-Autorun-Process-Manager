# Software Name: Windows-Autorun-Process-Manager
# Author: Bocaletto Luca
# Language: English
# Import necessary libraries including tkinter and other required modules

import tkinter as tk
import tkinter.ttk as ttk  # Import the ttk module from tkinter
import os
import subprocess
import ctypes  # Import ctypes for handling error messages
import winreg  # Import winreg for manipulating the Windows system registry
from tkinter import filedialog  # Import the filedialog module from tkinter for file selection dialogs

# Function to retrieve the Windows startup applications
def get_startup_apps():
    startup_apps = []
    # Open the HKEY_CURRENT_USER registry key to read the startup applications
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ) as key:
        try:
            index = 0
            while True:
                # Enumerate the values in the registry key (application name and path)
                name, value, _ = winreg.EnumValue(key, index)
                startup_apps.append((name, value))
                index += 1
        except WindowsError:
            pass
    return startup_apps

# Function to add an application to Windows startup
def add_startup_app():
    app_path = app_path_entry.get()  # Get the application path from user input
    app_name = os.path.basename(app_path)  # Extract the application name from the path
    # Open the registry key to write a new value (application name and path)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE) as key:
        winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path)
    refresh_list()  # Refresh the list of startup applications

# Function to select the application path using the "Browse" button
def select_path():
    file_path = filedialog.askopenfilename()  # Open a file selection dialog
    app_path_entry.delete(0, tk.END)  # Clear the contents of the entry widget
    app_path_entry.insert(0, file_path)  # Insert the file path into the entry widget

# Function to remove an application from Windows startup
def remove_startup_app():
    selected_item = startup_apps_list.selection()  # Get the selected item in the list
    if selected_item:
        app_name = startup_apps_list.item(selected_item, "values")[0]  # Get the application name from the list
        # Open the registry key to remove the application's startup value
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE) as key:
            winreg.DeleteValue(key, app_name)
        refresh_list()  # Refresh the list of startup applications

# Function to launch an application
def launch_application():
    selected_item = startup_apps_list.selection()  # Get the selected item in the list
    if selected_item:
        app_path = startup_apps_list.item(selected_item, "values")[1]  # Get the application's path from the list
        try:
            subprocess.Popen([app_path], shell=True)  # Launch the application using subprocess
        except Exception as e:
            # Display an error message if the application fails to launch
            ctypes.windll.user32.MessageBoxW(0, f"Unable to launch the application.\nError: {str(e)}", "Error", 0)

# Function to refresh the list of startup applications
def refresh_list():
    startup_apps_list.delete(*startup_apps_list.get_children())  # Clear all items in the current list
    for app_name, app_path in get_startup_apps():
        # Add current startup applications to the list
        startup_apps_list.insert("", "end", values=(app_name, app_path))

# Create the main window
root = tk.Tk()  # Create a new tkinter window
root.title("Windows Autorun Process Manager")  # Set the window title

# Label and entry for the application path
app_path_label = tk.Label(root, text="Application Path:")  # Create a label
app_path_label.pack()  # Display the label in the window
app_path_entry = tk.Entry(root)  # Create an entry widget for the path
app_path_entry.pack()  # Display the entry widget in the window

# "Browse" button to select the application path
browse_button = tk.Button(root, text="Browse", command=select_path)  # Create a button
browse_button.pack()  # Display the button in the window

# Buttons to add, remove, and launch applications
add_button = tk.Button(root, text="Add to Startup", command=add_startup_app)  # Create a button to add an application
remove_button = tk.Button(root, text="Remove from Startup", command=remove_startup_app)  # Create a button to remove an application
launch_button = tk.Button(root, text="Launch Application", command=launch_application)  # Create a button to launch an application
add_button.pack()  # Display the "Add" button in the window
remove_button.pack()  # Display the "Remove" button in the window
launch_button.pack()  # Display the "Launch" button in the window

# List of startup applications
startup_apps_list = ttk.Treeview(root, columns=("Name", "Path"), show="headings")  # Create a two-column list
startup_apps_list.heading("Name", text="Name")  # Set the header for the "Name" column
startup_apps_list.heading("Path", text="Path")  # Set the header for the "Path" column
startup_apps_list.pack()  # Display the list in the window
refresh_list()  # Refresh the list of startup applications

# Run the GUI main loop
root.mainloop()  # Start the main loop for the graphical interface
