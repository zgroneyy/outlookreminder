'''
Teams eklenecek
Teams ya da Outlook hangisi yoksa o aya[a kaldirilacak
UI eklenecek 
Gelen emaili yansitma gelecek 
'''
import os
import time
import subprocess
import win32com.client
import ctypes

def execute_outlook():
    os.startfile("outlook")
    time.sleep(5)
    print("Outlook is starting now...")

def process_exists(process_name):
    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    # use buildin check_output right away
    output = subprocess.check_output(call).decode()
    # check in last line for process name
    last_line = output.strip().split('\r\n')[-1]
    # because Fail message could be translated
    return last_line.lower().startswith(process_name.lower())   

def new_email_check():
    ol = win32com.client.Dispatch( "Outlook.Application")
    inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    for message in inbox.Items:
        if message.UnRead == True:
            MessageBox = ctypes.windll.user32.MessageBoxW
            #print(message.SenderEmailAddress)
            MessageBox(None, message.Subject, 'new e mail' , 0)
            print(message.Subject) #or whatever command you want to do
        
while True:
    if(process_exists('OUTLOOK.exe')):
        new_email_check()
        time.sleep(30)
    else:
        execute_outlook()