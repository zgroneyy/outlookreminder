import os
import time
import subprocess
import win32com.client
import tkinter as tk
import pyautogui
import threading
import random

# Define global variables
program_active = False
email_checking_active = False
mouse_moving_active = False
idle_time_threshold = 300  # 300 seconds (5 minutes) of idle time
last_mouse_activity = 0  # Initialize the variable
mouse_moving_thread = None  # Initialize the variable

def execute_outlook():
    os.startfile("outlook")
    time.sleep(5)
    print("Outlook is starting now...")

def execute_teams():
    os.system('C:/Users/Raghav/AppData/Local/Microsoft/Teams/Update.exe --processStart "Teams.exe"')
    time.sleep(5)
    print("Teams is starting now...")

def process_exists(process_name):
    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    output = subprocess.check_output(call).decode()
    last_line = output.strip().split('\r\n')[-1]
    return last_line.lower().startswith(process_name.lower())

def new_email_check():
    ol = win32com.client.Dispatch("Outlook.Application")
    inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    for message in inbox.Items:
        if message.UnRead:
            subject = message.Subject
            sender = message.SenderEmailAddress
            open_custom_dialog(subject, sender)
            print(f"Subject: {subject}\nSender: {sender}")
            message.UnRead = False  # Mark the message as read

def open_custom_dialog(subject, sender):
    dialog = tk.Toplevel(root)
    dialog.title("New Email")
    message_label = tk.Label(dialog, text=f"Subject: {subject}\nSender: {sender}")
    message_label.pack(padx=10, pady=10)
    got_it_button = tk.Button(dialog, text="Got it", command=dialog.destroy)
    got_it_button.pack(pady=10)

def mouse_mover():
    global mouse_moving_active

    while True:
        if mouse_moving_active and mouse_moving_checkbox_var.get():
            screen_width, screen_height = pyautogui.size()
            x = random.randint(0, screen_width)
            y = random.randint(0, screen_height)
            pyautogui.moveTo(x, y, duration=5)
            time.sleep(30)  # Move the mouse cursor every 60 seconds
        else:
            time.sleep(15)
            if not mouse_moving_active or not mouse_moving_checkbox_var.get():
                while not mouse_moving_active or not mouse_moving_checkbox_var.get():
                    time.sleep(1)
                    if mouse_moving_active and mouse_moving_checkbox_var.get():
                        break
                if not mouse_moving_active or not mouse_moving_checkbox_var.get():
                    continue

def update_activation_label():
    if program_active or email_checking_active or mouse_moving_active:
        activation_label.config(text="Work is being done, be patient!", font=("Helvetica", 14))
    else:
        activation_label.config(text="Program is not active", font=("Helvetica", 14))

def toggle_email_checking():
    global email_checking_active
    if email_checking_checkbox_var.get():
        email_checking_active = True
        if program_active:
            check_for_new_mail_periodically()  # Start periodic email checking
    else:
        email_checking_active = False
    update_activation_label()

def toggle_mouse_moving():
    global mouse_moving_active, mouse_moving_thread

    if mouse_moving_checkbox_var.get():
        mouse_moving_active = True

        if not mouse_moving_thread:
            # Start the mouse_mover in a separate thread
            mouse_moving_thread = threading.Thread(target=mouse_mover)
            mouse_moving_thread.daemon = True
            mouse_moving_thread.start()
    else:
        mouse_moving_active = False
    update_activation_label()

def activate_program():
    global program_active, last_mouse_activity, mouse_moving_thread

    if not program_active:
        program_active = True
        last_mouse_activity = time.time()  # Initialize last_mouse_activity
        activate_button.config(bg="green")  # Change the button color when activated
        deactivate_button.config(bg="SystemButtonFace")  # Reset the color of the Deactivate button

        # Start the check_program_loop in a separate thread
        program_thread = threading.Thread(target=check_program_loop)
        program_thread.daemon = True
        program_thread.start()

        if email_checking_checkbox_var.get():
            check_for_new_mail_periodically()

        if mouse_moving_checkbox_var.get():
            mouse_moving_active = True

            if not mouse_moving_thread:
                # Start the mouse_mover in a separate thread
                mouse_moving_thread = threading.Thread(target=mouse_mover)
                mouse_moving_thread.daemon = True
                mouse_moving_thread.start()

    update_activation_label()

def deactivate_program():
    global program_active
    if program_active:
        program_active = False
        activate_button.config(bg="SystemButtonFace")  # Reset the color of the Activate button
        deactivate_button.config(bg="red")  # Change the button color when deactivated
    update_activation_label()

def check_program_loop():
    global last_mouse_activity
    while program_active:
        outlook_running = process_exists('OUTLOOK.exe')
        teams_running = process_exists('Teams.exe')

        if not outlook_running:
            execute_outlook()

        if not teams_running:
            execute_teams()

        # Check mouse inactivity and move the mouse slightly if idle
        current_time = time.time()
        if (current_time - last_mouse_activity) >= idle_time_threshold:
            pyautogui.move(1, 1)
            last_mouse_activity = current_time

        time.sleep(30)

def check_for_new_mail_periodically():
    if program_active and email_checking_active:
        new_email_check()
        root.after(60000, check_for_new_mail_periodically)  # Schedule the function to run again in 60 seconds

def on_closing():
    global program_active
    program_active = False
    root.destroy()

# Create the main Tkinter window
root = tk.Tk()
root.title("Outlook Auto-Opener")

# Set the window size to 200x350 and make it a fixed size
root.geometry("350x300")
root.resizable(False, False)

# Set the background color
root.configure(bg="#0066a1")

# Bind the window's close event to the on_closing function
root.protocol("WM_DELETE_WINDOW", on_closing)

# Create a label to display program status
activation_label = tk.Label(root, text="Program is not active", bg="white", font=("Helvetica", 14))
activation_label.pack(pady=10)

# Create the Activate button
activate_button = tk.Button(root, text="Activate", command=activate_program, width=10)
activate_button.pack(padx=10, pady=10)

# Create the Deactivate button
deactivate_button = tk.Button(root, text="Deactivate", command=deactivate_program, width=10)
deactivate_button.pack(padx=10, pady=10)

# Create a checkbox for email checking with a #0066a1 background
email_checking_checkbox_var = tk.BooleanVar()
email_checking_checkbox = tk.Checkbutton(
    root,
    text="Check my Mail",
    variable=email_checking_checkbox_var,
    command=toggle_email_checking,
    bg="#0066a1",
    activebackground="#0066a1",
    selectcolor="#0066a1"
)
email_checking_checkbox.pack(padx=10, pady=10)

# Create a checkbox for mouse movement with a #0066a1 background
mouse_moving_checkbox_var = tk.BooleanVar()
mouse_moving_checkbox = tk.Checkbutton(
    root,
    text="Move my mouse",
    variable=mouse_moving_checkbox_var,
    command=toggle_mouse_moving,
    bg="#0066a1",
    activebackground="#0066a1",
    selectcolor="#0066a1"
)
mouse_moving_checkbox.pack(padx=10, pady=10)

# Start the Tkinter main loop
root.mainloop()
