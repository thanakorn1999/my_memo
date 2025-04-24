import os
import argparse
from dotenv import load_dotenv
import pyperclip
import keyboard
import pygetwindow as gw
import pandas as pd
import re
from sys import platform
if platform =='win32':
    import win32api
    win32api.LoadKeyboardLayout('00000409',1) # to switch to english
import time
import pyautogui as pg
import random
import string
import ctypes
import threading
parser = argparse.ArgumentParser(description='Script so useful.')
parser.add_argument("--t", type=str, default='')

args = parser.parse_args()
vda = ctypes.CDLL('./VirtualDesktopAccessor.dll')
if platform =='win32':
    cm_paste ="ctrl+v"
    cm_del ="\b"
elif platform =='darwin':
    cm_paste ="command+v"
    cm_del ="delete"

default_history='012345678910'
history =default_history
# ? AUTO-FILL

def generate_random_data():
    random_number = random.randint(1, 100)
    # random_string = ''.join(random.choices(string.ascii_letters + string.digits, k=6))
    letters = string.ascii_lowercase.replace('e', 'a').replace('E', 'A')  # Remove 'e' and 'E'
    random_string=  ''.join(random.choice(letters) for _ in range(10))
    return random_number, random_string

def auto_fill():
    pg.click(pg.position())
    pg.keyDown('up')
    pg.keyDown('enter')
    random_number, random_string = generate_random_data()
    text = str(random_number)+random_string
    print(text)
    # pg.write(text, interval=0.001) 
    pg.write(text) 
    pg.keyDown('enter')
    pg.keyDown('esc')
    
def export_joget():
    pg.keyDown('tab')
    pg.keyDown('space')      
    pg.keyUp('space')  
    pg.keyDown('tab')
    pg.keyDown('space')      
    pg.keyUp('space')  
    pg.keyDown('tab')
    pg.keyDown('space')      
    pg.keyUp('space')  
    pg.keyDown('tab')
    pg.keyDown('space')      
    pg.keyUp('space')  
    pg.keyDown('tab')
    pg.keyDown('space')      
    pg.keyUp('space')   
    pg.keyDown('tab')
    pg.keyDown('space')    
    pg.keyUp('space')  
    pg.keyDown('tab')
    pg.keyDown('space')    
    pg.keyUp('space')  
    pg.keyDown('tab')
    pg.keyDown('space')    
    pg.keyUp('space')  
    pg.keyDown('tab')
    pg.keyDown('space')    
    pg.keyUp('space')  
    pg.keyDown('tab')
    pg.keyDown('space')    
    pg.keyUp('space')         

def auto_fill_next_fiel():
    # pg.click(pg.position())
    pg.keyDown('tab')
    pg.keyDown('up')
    pg.keyDown('up')
    pg.keyDown('enter')
    random_number, random_string = generate_random_data()
    text = str(random_number)+random_string
    print(text)
    # pg.write(text, interval=0.001) 
    pg.write(text) 
    pg.keyDown('enter')
    pg.keyDown('esc')


def back_space(action):
    global history, cm_del
    if action :
        keyboard.send(cm_del)
    history='0'+history[:len(history)-1]

def write(replacement,command):
    global history, cm_paste, cm_del, default_history
    history = default_history

    save_old_copy = pyperclip.paste()
    for n in range(len(command)):
        keyboard.send(cm_del)

    pyperclip.copy(replacement)
    keyboard.send(cm_paste)
    #! need to update logic error
    time.sleep(0.2)
    pyperclip.copy(save_old_copy)

def chang_command_mac(command_list):
    new_command_list =[]
    for command in command_list:
        text = command
        if '!' in text:
            text = re.sub('!', '1', text)
        if '@' in text:
            text = re.sub('@', '2', text)
        if '#' in text:
            text = re.sub('#', '3', text)
        if '%' in text:
            text = re.sub('%', '5', text)
        if '_' in text:
            text = re.sub('_', '-', text)
        new_command_list.append(text)

    print(new_command_list)
    return new_command_list

def load_data_excel(platform):
    load_dotenv()
    list_memo = os.getenv('LIST_MEMO').split(',')
    print('list_memo',list_memo)
    
    data = pd.concat([pd.read_excel('./memo/memo.xlsx', sheet_name = sheet) for sheet in list_memo], ignore_index = True)

    data['command'] = data['command'].astype(str)

    data['message'] = data['message'].astype(str)

    # print(data.to_string())

    data = data.dropna()

    lenght_command=[]
    for text in data['command']:
        lenght_command.append(len(str(text)))

    data['lenght_command'] = lenght_command
    data = data.sort_values(by='lenght_command', ascending=False)

    message_json = data['message'].values.tolist()
    command_list = data['command'].values.tolist()
    
    #! need to update logic error
    if platform == 'darwin':
        command_list = chang_command_mac(command_list)

    print(command_list)
    return command_list, message_json

def check_map_command(history):
    global command_list,message_json
    is_match = False
    for index,command in enumerate(command_list) :
        if command in history :
            write(message_json[index],command_list[index])
            is_match = True
            break
    
    if not is_match:
        back_space(True)
        print('not match')
        for command in command_list:
            print(command)

def released(release):
    global cm_del ,history

    if(len(release))==1:
        history=history[1:]+release
        print(history)
    elif release=='right shift':
        check_map_command(history)
    elif release=='delete':
        back_space(False)
    elif release=='f4':
        # tab 
        # space
        export_joget()
    elif release=='print screen':
        auto_fill()
    elif release=='home':
        auto_fill_next_fiel()
    else :
        print(release)



# Define constants
SW_SHOWNORMAL = 1
SW_SHOWMINIMIZED = 2
SW_SHOWMAXIMIZED = 3

# Load user32.dll
user32 = ctypes.windll.user32
# 
# 
# 
backend_identifier = ""
frontend_identifier = ""
test_identifier = ""
front_test_google_identifier = ""
front_google_identifier = ""

# 
# lang = 'OSS' 
# lang = 'OEE-JAME' 
lang = args.t
if lang == "OSS":
    backend_identifier = "efund-backend - Visual Studio Code"
    frontend_identifier = "efund-frontend - Visual Studio Code"
    test_identifier = "oss-test-cyp - Visual Studio Code"
    front_test_google_identifier = "oss-test-cyp - Google Chrome"
    front_google_identifier = "e-Fund | ERC - Google Chrome"
    front_error = "An error occurred |"
elif lang == "RPGG":
    # backend_identifier = "ui - Visual Studio Code"
    backend_identifier = "react-game-backend - Visual Studio Code"
    frontend_identifier = "rgame-front - Visual Studio Code"
    test_identifier = ""
    front_google_identifier = "Vite + React + TS"
    # 
    front_test_google_identifier = "oss-test-cyp - Google Chrome"
    front_error = "An error occurred |"
elif lang == "ELI":
    backend_identifier = "backend - Visual Studio Code"
    frontend_identifier = "frontend - Visual Studio Code"
    test_identifier = ""
    front_google_identifier = "e-Licensing | ERC - Google Chrome"
    front_test_google_identifier = "oss-test-cyp - Google Chrome"
    front_error = "An error occurred |"

def switch_to_window(window_identifier):
    try:
        windows = gw.getWindowsWithTitle(window_identifier)
        if windows:
            window = windows[0]
            if window.isMinimized:
                window.restore()  # Restore the window if it's minimized
                time.sleep(0.2)  # Adding a small delay to ensure the window is restored

            hwnd = window._hWnd
            desktop_index = vda.GetWindowDesktopNumber(hwnd)
            current_desktop_index = vda.GetCurrentDesktopNumber()

            if desktop_index != current_desktop_index:
                vda.GoToDesktopNumber(desktop_index)

            # Bring window to foreground using ctypes
            user32.ShowWindow(hwnd, SW_SHOWMAXIMIZED)
            user32.SetForegroundWindow(hwnd)
            time.sleep(0.2)  # Adding a small delay to ensure the window is activated
            print(f"Switched to: {window.title}")  # Show actual window title after switch
        else:
            print(f"Window not found with identifier: {window_identifier}")
    except gw.PyGetWindowException as e:
        if e.args[0].endswith("The operation completed successfully."):
            print("Window activated successfully despite the error message.")
        else:
            print(f"Error: {e}")
            raise
def on_activate_f():
    switch_to_window(frontend_identifier)
    switch_to_window(frontend_identifier)
def on_activate_b():
    switch_to_window(backend_identifier)
    switch_to_window(backend_identifier)
def on_activate_t():
    switch_to_window(test_identifier)
    switch_to_window(test_identifier)
def on_activate_ft():
    switch_to_window(front_test_google_identifier)
    switch_to_window(front_test_google_identifier)

def on_activate_ff():
    switch_to_window(front_google_identifier)
    switch_to_window(front_google_identifier)
    switch_to_window(front_error)
    switch_to_window(front_error)


# def main():
#     keyboard.add_hotkey('alt+1', on_activate_f)
#     keyboard.add_hotkey('alt+2', on_activate_b)
#     keyboard.add_hotkey('alt+3', on_activate_t)
#     keyboard.add_hotkey('alt+w+3', on_activate_ft)
#     keyboard.add_hotkey('alt+w+1', on_activate_ff)
#     keyboard.on_release(lambda e: released( e.name ))
#     keyboard.wait()

# command_list ,message_json= load_data_excel(platform)

# def load_and_run():
#     global command_list, message_json
#     command_list, message_json = load_data_excel(platform)
#     main()

# # ใช้ thread เพื่อแยกการโหลดข้อมูลออกจาก main loop ของ keyboard
# threading.Thread(target=load_and_run, daemon=True).start()

# # main()
def main():
    keyboard.add_hotkey('alt+1', on_activate_f)
    keyboard.add_hotkey('alt+2', on_activate_b)
    keyboard.add_hotkey('alt+3', on_activate_t)
    keyboard.add_hotkey('alt+w+3', on_activate_ft)
    keyboard.add_hotkey('alt+w+1', on_activate_ff)
    keyboard.on_release(lambda e: released(e.name))
    print("Hotkeys registered, waiting for key events...")
    keyboard.wait()  # Main loop, stays here

def load_data():
    global command_list, message_json
    print("Loading data...")
    command_list, message_json = load_data_excel(platform)
    print("Data loaded!")

# โหลดข้อมูลก่อนใน thread แยก เพื่อไม่ให้บล็อค main thread
threading.Thread(target=load_data, daemon=True).start()

# รัน main loop ของ keyboard ใน main thread
main()