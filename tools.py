
# from email.headerregistry import AddressHeader
# from ssl import ALERT_DESCRIPTION_ACCESS_DENIED
# from urllib.request import AbstractDigestAuthHandler
import pyperclip
import keyboard
import pandas as pd
import re
from sys import platform
if platform =='win32':
    import win32api
    win32api.LoadKeyboardLayout('00000409',1) # to switch to english
import time

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
        # if '$' in text:
        #     text = re.sub('$', '4', text)
        if '%' in text:
            text = re.sub('%', '5', text)
        if '_' in text:
            text = re.sub('_', '-', text)
        new_command_list.append(text)

    print(new_command_list)
    return new_command_list

def load_data_excel(platform):
    # data = pd.read_excel('./memo/memo.xlsx')
    sheets = ['react', 'node']
    
    data = pd.concat([pd.read_excel('./memo/memo.xlsx', sheet_name = sheet) for sheet in sheets], ignore_index = True)

    data['command'] = data['command'].astype(str)
    data['message'] = data['message'].astype(str)

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

    return command_list, message_json




def check_map_command(history):
    global command_list,message_json
    is_match = False
    for index,command in enumerate(command_list) :
        print(command)
        if command in history :
            write(message_json[index],command_list[index])
            is_match = True
            break
    
    if not is_match:
        back_space(True)

def released(release):
    global cm_del,history

    if(len(release))==1:
        history=history[1:]+release
        print(history)
    elif release=='right shift':
        check_map_command(history)
    # elif release=='backspace':delete
    elif release=='delete':
        back_space(False)
    else :
        print(release)



def main(history):
    keyboard.on_release(lambda e: released( e.name )) # old logic keyboard.add_abbreviation('@@', 'my.long.email@example.com')
    keyboard.wait()


if platform =='win32':
    cm_paste ="ctrl+v"
    cm_del ="\b"
elif platform =='darwin':
    cm_paste ="command+v"
    cm_del ="delete"

command_list ,message_json= load_data_excel(platform)
default_history='0123456'
history =default_history

main(history)
