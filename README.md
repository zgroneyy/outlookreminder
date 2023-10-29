# Outlook and Teams Auto-Opener

Automatically manage your Microsoft Outlook and Teams applications with this Python-based desktop automation tool. The program is designed to make it easy for users to keep their Outlook and Teams applications open, check for new emails, and simulate mouse activity to prevent the computer from going into sleep mode.

## Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [License](#license)

## Features
- Automatically opens Microsoft Outlook and Microsoft Teams applications if they are not running.
- Periodically checks for new emails in Outlook and displays notifications.
- Simulates mouse activity to prevent the computer from going into sleep mode.
- User-friendly graphical user interface (GUI) for easy configuration.
- Activate and deactivate the program with a single button click.
- Customize email checking and mouse movement options.

## Requirements
- Python 3.x
- Windows operating system (Tested on Windows 10)
- [pyautogui](https://pypi.org/project/PyAutoGUI/) library
- [pywin32](https://pypi.org/project/pywin32/) library
- tkinter (usually comes pre-installed with Python)

## Installation
1. Clone the repository to your local machine using the following command:
   ```shell
   git clone https://github.com/zgroneyy/outlookreminder
   ```
2. Navigate to the project directory:
   ```shell
   cd outlookreminder
   ```
3. Install the required Python libraries if you haven't already:
   ```shell
   pip install pyautogui pywin32
   ```

## Usage
1. Run the program by executing the `main.py` file:
   ```shell
   python main.py
   ```

2. Use the program's graphical user interface to activate and deactivate the automation.
   
3. Configure email checking and mouse movement options as desired.

4. The program will automatically open Outlook and Teams if they are not running, check for new emails, and simulate mouse activity.

5. Customize the program to fit your needs and preferences.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
```

