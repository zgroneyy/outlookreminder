import os
import time
import subprocess
import win32com.client
import tkinter as tk
from tkinter import messagebox
import pyautogui

# Define a flag to track the program's activation state and email checking
program_active = False
email_checking_active = False
mouse_moving_active = False
idle_time_threshold = 300  # 300 seconds (5 minutes) of idle time

def execute_outlook():
    os.startfile("outlook")
    time.sleep(5)
    print("Outlook is starting now...")

def execute_teams():
    os.startfile("teams")
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
    while mouse_moving_active:
        pyautogui.move(1, 1)
        time.sleep(60)  # Move the mouse cursor every 60 seconds

def activate_program():
    global program_active
    if not program_active:
        program_active = True
        activation_label.config(text="Program is now at work, just sit back")
        activate_button.config(bg="green")  # Change the button color when activated
        deactivate_button.config(bg="SystemButtonFace")  # Reset the color of the Deactivate button
        check_program_loop()
        if email_checking_checkbox_var.get():  # If the email checking checkbox is marked, start periodic email checking
            check_for_new_mail_periodically()
        if mouse_moving_checkbox_var.get():  # If the mouse movement checkbox is marked, start moving the mouse
            mouse_moving_active = True
            mouse_mover()

def deactivate_program():
    global program_active
    if program_active:
        program_active = False
        activation_label.config(text="Program voluntarily deactivated, thanks for using")
        deactivate_button.config(bg="red")  # Change the button color when deactivated
        activate_button.config(bg="SystemButtonFace")  # Reset the color of the Activate button
        mouse_moving_active = False  # Stop moving the mouse when the program is deactivated

def toggle_email_checking():
    global email_checking_active
    if email_checking_checkbox_var.get():
        email_checking_active = True
        if program_active:
            check_for_new_mail_periodically()  # Start periodic email checking
    else:
        email_checking_active = False

def toggle_mouse_moving():
    global mouse_moving_active
    if mouse_moving_checkbox_var.get():
        mouse_moving_active = True
        if program_active:
            mouse_mover()  # Start moving the mouse cursor
    else:
        mouse_moving_active = False

def check_program_loop():
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

# Create the main Tkinter window
root = tk.Tk()
root.title("Outlook Auto-Opener")

# Set the window size to 320x200 and make it a fixed size
root.geometry("320x200")
root.resizable(False, False)

# Set the background color
root.configure(bg="#0066a1")

# Create the Activate button
activate_button = tk.Button(root, text="Activate", command=activate_program)
activate_button.pack()

# Create the Deactivate button
deactivate_button = tk.Button(root, text="Deactivate", command=deactivate_program)
deactivate_button.pack()

# Create a label to display activation status
activation_label = tk.Label(root, text="Program is not active", bg="#0066a1", fg="white")
activation_label.pack()

# Create a checkbox for email checking
email_checking_checkbox_var = tk.BooleanVar()
email_checking_checkbox = tk.Checkbutton(root, text="Check my mail", variable=email_checking_checkbox_var, command=toggle_email_checking)
email_checking_checkbox.pack()

# Create a checkbox for mouse movement
mouse_moving_checkbox_var = tk.BooleanVar()
mouse_moving_checkbox = tk.Checkbutton(root, text="Move my mouse", variable=mouse_moving_checkbox_var, command=toggle_mouse_moving)
mouse_moving_checkbox.pack()

# Start the Tkinter main loop
root.mainloop()
