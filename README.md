# ColdEmailAutomation

This project is a mass email automation tool specifically designed for use with the Outlook app on a Surface laptop. The tool automates the process of sending personalized emails to a list of recipients, scheduling them at specific times, and managing email content dynamically from the clipboard.

## Features

- **Automated Email Sending:** Automatically sends emails to a list of recipients, using Outlook.
- **Personalization:** Customizes each email with the recipient's name and company, extracted from an Excel file.
- **Scheduling:** Emails are scheduled to be sent at specific times, avoiding non-working hours and weekends.
- **Clipboard Integration:** Allows the email body to be copied to the clipboard before execution, ensuring consistent content across all emails.
- **Excel Integration:** Reads recipient details (name, email, and company) from an Excel file, and logs sent emails to a separate file.
- **Image Recognition:** Uses image recognition to interact with the Outlook interface, ensuring compatibility with specific layouts and screen configurations.

## Prerequisites

- **Operating System:** Designed specifically for use on a Surface laptop running Windows.
- **Outlook Application:** Requires the Outlook desktop application to be installed and running.
- **Python 3.7+** with the following libraries:
  - `pyautogui`
  - `pandas`
  - `numpy`
  - `python_calamine`
  - `xlsxwriter`
  - `opencv-python`
  - `tkinter` (for the GUI)
  
You can install the necessary Python packages using pip:
```bash
pip install pyautogui pandas numpy python-calamine xlsxwriter opencv-python
