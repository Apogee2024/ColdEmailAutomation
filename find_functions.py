
from time import sleep
from pyautogui import moveTo, click, write, hotkey, press, size, locateCenterOnScreen, PAUSE, screenshot as screen
import pandas
import numpy as np

import python_calamine
import xlsxwriter
import cv2

from datetime import datetime, date, timedelta

import cv2
import numpy as np

COMPANY_NAME = 'Company Name'

def locate_and_click_image(target_image_path, threshold=0.8, num_clicks=1):
    # Capture the screenshot using pyautogui
    screenshot = screen()

    # Convert the screenshot to a numpy array
    screenshot_np = np.array(screenshot)

    # Convert the screenshot to grayscale
    screenshot_gray = cv2.cvtColor(screenshot_np, cv2.COLOR_BGR2GRAY)

    # Read the target image
    target_image = cv2.imread(target_image_path, cv2.IMREAD_GRAYSCALE)

    # Get the width and height of the target image
    w, h = target_image.shape[::-1]

    # Perform template matching
    result = cv2.matchTemplate(screenshot_gray, target_image, cv2.TM_CCOEFF_NORMED)

    # Get the location of the best match
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # Check if the best match exceeds the threshold
    if max_val >= threshold:
        # # Define the top-left corner of the matching area
        top_left = max_loc
        #
        # # Define the bottom-right corner of the matching area
        # bottom_right = (top_left[0] + w, top_left[1] + h)
        #
        # # Draw a rectangle around the matched region for visualization
        # cv2.rectangle(screenshot_np, top_left, bottom_right, (0, 255, 0), 2)
        #
        # # Display the result
        # cv2.imshow('Match', screenshot_np)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()

        # Click the center of the matched area the specified number of times
        center_x = top_left[0] + w // 2
        center_y = top_left[1] + h // 2

        # Move to the center of the matched area before clicking
        moveTo(center_x, center_y)
        for _ in range(num_clicks):
            click()
        return True

    else:
        return False
        print("No match found with the given threshold.")




def findandclick_image_color(img_name, ConfidLvL=.7, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x + (x1 / 100), y, MouseSpeedTime)
    if Click == True:
        click()


def findandclick(img_name, ConfidLvL=.7, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x + (x1 / 100), y, MouseSpeedTime)
    if Click == True:
        click()


def click_more_colors(img_name, ConfidLvL=.7, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x, y + (y1 / 100), MouseSpeedTime)
    if Click == True:
        sleep(.1)
        click()


def click_hex_and_enter_color(img_name, ConfidLvL=.7, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x, y + (y1 / 30), MouseSpeedTime)
    if Click == True:
        sleep(.1)
        click()
    hotkey("ctrl", 'a')
    write("203864")
    press('enter')


def change_font_size(img_name, ConfidLvL=.7, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x + (x1 / 100), y, MouseSpeedTime)
    if Click == True:
        click()

    press('tab')
    write('11')
    press('enter')


def change_font(img_name, ConfidLvL=.7, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x + (x1 / 100), y, MouseSpeedTime)
    if Click == True:
        click()
    hotkey("ctrl", 'a')
    write("Calibri")
    press('enter')


def back_to_message(img_name, ConfidLvL=.7, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x + (x1 / 100), y + (y1 / 8), MouseSpeedTime)
    if Click == True:
        click()


def text_format():
    """format the text of the body"""
    hotkey('ctrl', 'a')
    findandclick_image_color("data/text_color.png")
    click_more_colors("data/more_colors.png")
    click_hex_and_enter_color("data/hex.png")
    click_more_colors("data/color_okay.png")
    change_font_size("data/navigate_font.png")
    change_font("data/navigate_font.png")
    back_to_message("data/to.png")


def scheduled_send_click_1(img_name, ConfidLvL=.638, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x + 108, y, MouseSpeedTime)
    if Click == True:
        click()
    sleep(.4)
    moveTo(x + 108, y + 119)
    click()


def scheduled_send_entry(img_name, ConfidLvL=.49, MouseSpeedTime=.01, PriorLoadingTime=1, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x + (x1 / 120), y, MouseSpeedTime)
    if Click == True:
        click()


def final_send(img_name, ConfidLvL=.7, MouseSpeedTime=.01, PriorLoadingTime=.5, Click=True, Move=True):
    x1, y1 = size()
    sleep(PriorLoadingTime)
    x, y = locateCenterOnScreen(img_name, confidence=ConfidLvL)
    if Move == True:
        moveTo(x - (x1 / 20), y, MouseSpeedTime)
    if Click == True:
        click()


def scheduled_send(calendar, hour):
   
    scheduled_send_click_1('data/send.png')
    sleep(.2)
   
    sleep(.2)
   
    findandclick("data/custom.png")
    sleep(.2)
    moveTo(1853, 811)  #clicking date entry box
    click()
    hotkey('ctrl', 'a')
    write(calendar)
    sleep(.2)
    moveTo(1857, 881)  #clicking time entry box
    click()
    hotkey('ctrl', 'a')
    write(hour)
    sleep(.2)
    press('enter')
    sleep(.2)

    moveTo(1771, 1364)  # clicking the send button in the custom time box
    click()


def scheduled_send_outlook(calendar, hour):
    tries = 0
    while not locate_and_click_image('outlook/options.png', threshold=0.5, num_clicks=2) and tries < 6:
        sleep(.2)
        tries +=1
    tries = 0
    tries = 0
    while not locate_and_click_image('outlook/delay_delivery.png', threshold=0.5, num_clicks=2) and tries < 6:
        sleep(.2)
        tries += 1
    tries = 0
    while not locate_and_click_image('data/do_not_deliver_before.png', threshold=0.5, num_clicks=3) and tries < 6:
        sleep(.2)
        tries +=1
    write(calendar)
    sleep(.2)
    press('tab')  #clicking time entry box
    write(hour)
    press('enter')
    tries = 0
    while not locate_and_click_image('outlook/send.png', threshold=0.8, num_clicks=3) and tries < 6:
        sleep(.2)
        tries +=1



def strip_in_parenthesis(test_str):
    """ removes text in parenthesis"""
    ret = ''
    skip1c = 0
    skip2c = 0
    for i in test_str:
        if i == '[':
            skip1c += 1
        elif i == '(':
            skip2c += 1
        elif i == ']' and skip1c > 0:
            skip1c -= 1
        elif i == ')' and skip2c > 0:
            skip2c -= 1
        elif skip1c == 0 and skip2c == 0:
            ret += i
    return ret




def increment_time(counter, day, hour, year, month):
    """handles incrementing time"""
    increment_day = False
    if counter == 20:
        hour += 1
        counter = 0
    if hour >= 18:
        hour = 8
        increment_day = True
    elif hour < 8:
        hour = 8

    # Convert day, month, and year to a datetime object
    current_date = datetime(year, month, day)

    # Increment day if necessary
    if increment_day:
        current_date += timedelta(days=1)
        day = current_date.day
        month = current_date.month
        year = current_date.year
        counter = 0

    # Increment day by 2 if it's Saturday
    if current_date.weekday() == 5:  # Saturday
        current_date += timedelta(days=2)
        day = current_date.day
        month = current_date.month
        year = current_date.year
        counter = 0

    time_obj = datetime.strptime(str(hour), "%H")
    formatted_hour = time_obj.strftime("%H:%M")

    counter += 1
    return formatted_hour, counter, day, hour, year, month

def load_date_data(filename):
    """loads date data from file"""
    # Initialize variables
    day = None
    month = None
    year = None
    counter = None
    hour = None
    # Open the file in read mode
    try:
        with open(filename, 'r') as file:
            # Read the lines from the file
            lines = file.readlines()
            today = datetime.today()
        # Iterate through the lines and extract values
            for line in lines:
            # Split the line into label and value
                label, value = line.strip().split(': ')
                # Assign value to the corresponding variable
                if label == 'Day':
                    day = int(value)
                elif label == 'Month':
                    month = int(value)
                elif label == 'Year':
                    year = int(value)
                elif label == 'Hour':
                    hour = int(value)
                elif label == 'Counter':
                    counter = int(value)

        # Check if all variables are initialized
        if day is not None and hour is not None and counter is not None and year is not None and month is not None:
            loaded_date = datetime(year, month, day)
            today_date = datetime(today.year, today.month, today.day)
            if loaded_date < today_date:
                year = today.year
                month = today.month
                day = today.day
                hour = today.hour
                counter = 0
            elif loaded_date == today_date and hour < today.hour:
                hour = today.hour
                counter = 0

            return hour, day, month, year, counter
        else:
            today = datetime.today()
            if today.weekday() == 5:  # Saturday
                today += timedelta(days=2)
            hour = today.hour
            if hour >= 18 or hour <= 7:
                hour = 8
            day = today.day
            month = today.month
            year = today.year
            counter = 0
            return hour, day, month, year, counter
    except FileNotFoundError:
        today = datetime.today()
        if today.weekday() == 5:  # Saturday
            today += timedelta(days=2)
        hour = today.hour
        if hour <= 7:
            hour = 8
        day = today.day
        month = today.month
        year = today.year
        counter = 0
        return hour, day, month, year, counter

def save_date_data(filename, day, month, year, hour, counter=None):
    """saves the date data to the file"""
    # Open the file in write mode
    with open(filename, 'w') as file:
        # Write the variables to the file
        file.write(f"Day: {day}\n")
        file.write(f"Month: {month}\n")
        file.write(f"Year: {year}\n")
        file.write(f"Hour: {hour}\n")
        if counter is not None:
            file.write(f"Counter: {counter}\n")



def launch_outlook():
    """launches outlook app"""
    hotkey('winleft', 'r')
    sleep(1)
    write('outlook')
    sleep(1)
    press('enter')
    sleep(15)

import os

def create_sent_file_path(filepath):
    """creates the filepath for the sent_file using the input file name"""
    # Split the filepath into directory and filename
    directory, filename = os.path.split(filepath)

    # Split the filename into name and extension
    name, extension = os.path.splitext(filename)

    # Append "_sent" to the name
    sent_filename = f"{name}_sent{extension}"

    # Join the directory and the new filename
    sent_file_path = os.path.join(directory, sent_filename)

    # Replace backslashes with forward slashes
    sent_file_path = sent_file_path.replace("\\", "/")

    return sent_file_path


import tkinter as tk
from tkinter import Canvas, filedialog, messagebox, PhotoImage
import os

class FileSelectionWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Automated Outlook Emailer Scheduler and Sender from Clipboard")
        self.root.configure(bg="white")

        self.message_line1 = "Please copy the message body into the clipboard"
        self.message_line2 = "(Ctrl+C or Copy) before clicking the 'Start' button."
        self.message_line3 = "Move the cursor to the top left corner of the screen"
        self.message_line4 = "to stop and disconnect external monitors before starting"

        self.message_label1 = tk.Label(self.root, text=self.message_line1, bg="white", fg="dark blue")
        self.message_label1.grid(row=1, column=0, columnspan=2, pady=5)

        self.message_label2 = tk.Label(self.root, text=self.message_line2, bg="white", fg="dark blue")
        self.message_label2.grid(row=2, column=0, columnspan=2, pady=5)

        self.greeting_label = tk.Label(self.root, text="Email Greeting:", bg="white", fg="dark blue")
        self.greeting_label.grid(row=3, column=0, pady=5)
        self.greeting_entry = tk.Entry(self.root, bg="white", fg="dark blue", width=20)
        self.greeting_entry.grid(row=3, column=1, pady=5)
        self.greeting_entry.insert(tk.END, "Dear")

        self.greeting_info_label = tk.Label(self.root, text="This is the word before recipient's name.", bg="white", fg="dark blue")
        self.greeting_info_label.grid(row=4, column=0, columnspan=2, pady=5)

        self.canvas = Canvas(self.root, height=200, width=200, bg="white", highlightthickness=0)
        self.logo_img = PhotoImage(file="data/R.png")
        self.canvas.create_image(100, 100, image=self.logo_img)
        self.canvas.grid(row=5, column=0, columnspan=2)

        self.select_button = tk.Button(self.root, text="Select File", command=self.select, bg="white", fg="dark blue", bd=2, relief=tk.RAISED)
        self.select_button.grid(row=6, column=0, columnspan=2, pady=10)

        self.column_requirement_label = tk.Label(self.root, text="xlsx must have columns: First Name, Company Name, and Email Address", bg="white", fg="dark blue")
        self.column_requirement_label.grid(row=7, column=0, columnspan=2, pady=5)

        self.file_path_label = tk.Label(self.root, text="", bg="white", fg="dark blue")
        self.file_path_label.grid(row=8, column=0, columnspan=2)

        self.start_button = tk.Button(self.root, text="Start", command=self.start, bg="white", fg="dark blue", bd=2, relief=tk.RAISED)
        self.start_button.grid(row=9, column=0, padx=10, pady=10)

        self.quit_button = tk.Button(self.root, text="Quit", command=self.quit, bg="white", fg="dark blue", bd=2, relief=tk.RAISED)
        self.quit_button.grid(row=9, column=1, padx=10, pady=10)
        self.message_label3 = tk.Label(self.root, text=self.message_line3, bg="white", fg="red")
        self.message_label3.grid(row=10, column=0, columnspan=2, pady=5)

        self.message_label4 = tk.Label(self.root, text=self.message_line4, bg="white", fg="red")
        self.message_label4.grid(row=11, column=0, columnspan=2, pady=5)

        self.selected_file = None
        self.user_start = False

    def select(self):
        self.selected_file = filedialog.askopenfilename(title="Select File")
        print("Selected file:", self.selected_file)
        if not self.selected_file.endswith('.xlsx'):
            messagebox.showinfo(title="Oops", message="Please make select an .xlsx file")
            self.selected_file = None
        else:
            # Extract file name from the selected file path
            file_name = os.path.basename(self.selected_file)
            self.file_path_label.config(text="Selected file: " + file_name)

    def start(self):
        if self.selected_file:
            greeting = self.greeting_entry.get()
            self.user_start = True
            self.root.quit()
        else:
            messagebox.showinfo(title="Error", message="No file selected.")

    def quit(self):
        self.root.quit()

    def run(self):
        self.root.mainloop()
        return self.selected_file, self.user_start, self.greeting_entry.get()

def enter_email(text):
    """copy paste and tabs down to subject"""
    write(text)
    sleep(2)
    press('enter')
    press('tab')
    # extra tab to skip over cc
    press('tab')


def enter_subject(text):
    """copy paste and tabs down"""
    write(text)
    sleep(.2)
    press('enter')



def make_subject(company):
    subject = f"{company} & {COMPANY_NAME}"
    return subject


def enter_subject_and_message_from_clipboard(subject, greeting, name):
    write(subject)
    press('tab')
    write(f"{greeting} {name},")
    sleep(.2)
    press('enter', 2)
    hotkey('ctrl', 'v')
    sleep(.01)

def automated_email_clip(filepath,greeting):
    """sends personalized emails with the body containing the contents of the clipboard"""
    launch_outlook()
    email_data = pandas.read_excel(filepath, engine="calamine")
    email_data = email_data[email_data['Email Address'].notna()]   # handle blank etries
    receiver_email_list = [item for item in email_data["Email Address"]]
    sent_num = 0   # for persisting
    #initialize sent file path
    sent_file_path = create_sent_file_path(filepath)
    for recipient in receiver_email_list:
        try:
            sent_email_data = pandas.read_excel(sent_file_path, engine="calamine")
        except:
            sent_email_data = pandas.DataFrame()
        if (sent_email_data == recipient).any().any():
            print(f"you already sent an email to {recipient}")
            sent_num += 1
            continue
        else:
            name_row = email_data[email_data['Email Address'] == recipient]
            name = name_row["First Name"].iloc[0]
            name = name.title()   #Format Name to titlecase
            hour, day, month, year, counter = load_date_data('date_data.txt')
            company = name_row["Company Name"].iloc[0]
            company = strip_in_parenthesis(company)
            formatted_hour, counter, day, hour, year, month = increment_time(counter, day, hour, year, month)
            calendar = f"{month}/{day}/{year}"
            subject = make_subject(company)
            findandclick_image_color('outlook/new_email.png')
            # copy and paste the recipient#
            sleep(.2)
            enter_email(recipient)
            # copy and paste the subject#
            enter_subject_and_message_from_clipboard(subject, greeting, name)
            sleep(.1)

            scheduled_send_outlook(calendar, formatted_hour)

            email_data = email_data.drop(email_data[email_data["Email Address"] == recipient].index)
            
            s = pandas.Series({'Email': recipient, 'First Name': name, 'Send Date ': calendar, 'Time Sent': formatted_hour,
                               'subject': subject})
            sent_email_data = pandas.concat([sent_email_data, s.to_frame().T])
            sent_email_data.to_excel(sent_file_path, engine='xlsxwriter', index=False)

            sleep(1.5)
            save_date_data('date_data.txt', day, month, year, hour, counter)
            print(f"{recipient},{counter},{formatted_hour},{calendar}")
            sent_num +=1
            print('you can also look at the sent_email list for a complete list')

    if sent_num == len(receiver_email_list):
        return True
    if sent_num != len(receiver_email_list):
        return False

def main_clipboard():
    file_window = FileSelectionWindow()
    file_path, user_started,greeting = file_window.run()
    if file_path is None:
        return
    if user_started:
        while automated_email_clip(file_path,greeting) is None:
            try:
                automated_email_clip(file_path,greeting)
            except:
                automated_email_clip(file_path,greeting)









if __name__ == "__main__":
    sleep(5)
    print(locate_and_click_image('data/do_not_deliver_before.png', threshold=0.7, num_clicks=3))
