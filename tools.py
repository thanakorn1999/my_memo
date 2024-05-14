import os
from dotenv import load_dotenv
import pyperclip
import keyboard
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
    elif release=='f2':
        auto_fill()
    elif release=='f4':
        auto_fill_next_fiel()
    else :
        print(release)

def main():
    keyboard.on_release(lambda e: released( e.name ))
    keyboard.wait()



command_list ,message_json= load_data_excel(platform)

main()
