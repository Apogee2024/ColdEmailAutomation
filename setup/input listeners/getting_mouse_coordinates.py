import win32api
import win32con
import time
import pyautogui
y_list = []
x_list = []


def click_event(event):
    """appends the coordinates of the current click location to the list"""
    global y_list
    global x_list
    x_list.append(event[0])
    y_list.append(event[1])
    return(event[0],event[1])


def get_mouse_position():
    """gets the mouse position"""
    return win32api.GetCursorPos()

def main():
    """waits 5 seconds before starting logging for user to switch applications"""
    global x_list
    global y_list
    list = []
    time.sleep(5)
    while True:
        if win32api.GetKeyState(win32con.VK_LBUTTON) < 0:
            x= click_event(get_mouse_position())
            time.sleep(0.5)  # Add a small delay to avoid detecting multiple clicks in quick succession
            list.append(x)
            print(x_list,y_list)
            print(list)

if __name__ == "__main__":
    main()
