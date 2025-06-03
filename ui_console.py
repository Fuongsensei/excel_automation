#pylint:disable = all
import sys, os, time, re
from colorama import init, Fore
from blessed import Terminal
import getpass
import constains

init()

if os.name == "nt":
    os.environ.setdefault("TERM", "xterm")
def apply_color(text : str | list ):
    if isinstance(text, str):
        return Fore.GREEN + text + Fore.RESET
    return ''.join([f"|{Fore.GREEN}{char}{Fore.RESET}|" for char in text])

def print_authors():
    msg : str = 'Dev by phuongnguyen1183 using Python, compiled to C'
    for c in msg.upper():
        sys.stdout.write(apply_color(c))
        sys.stdout.flush()
        time.sleep(0.02)
    print('\n'*2)

def get_des_path(callback):
    path_file : str = rf"C:\Users\{getpass.getuser()}\Documents\path.txt"
    if os.path.exists(path_file):
        with open(path_file) as f:
            path = f.read()
        if input(f"Đường dẫn hiện tại là {path}. Thay đổi? [{apply_color('Y')}/N]: ").upper() == 'Y':
            os.system('cls'); return callback()
        os.system('cls'); return path
    else:
        new_path : str = input(f"Dán đường dẫn file {apply_color('Report Scan Verify Shiftly (RCV)')}: ")
        with open(path_file, 'w') as f: f.write(new_path)
        os.system('cls')
        return new_path

def change_des_path():
    new_path : str = input("Nhập đường dẫn mới: ")
    os.system('cls')
    return new_path

def get_list_sap():
    sap_input: str  = input("Nhập số SAP cách nhau bởi ký tự không phải số: ")
    os.system('cls')
    sap_list : list = [s.strip() for s in re.split(r'\D+', sap_input)]
    confirm : str = input(f"Danh sách SAP: {apply_color(sap_list)} — Nhấn Enter để xác nhận, N để nhập lại: ")
    os.system('cls')
    return sap_list if confirm.upper() != 'N' else get_list_sap()

def print_loading():
    
    while not constains.is_event.is_set():
        start = constains.progress
        constains.done.wait()
        
        for i in range(start ,constains.progress + 1):
            bar = '█' * i
            sys.stdout.write(f'\rLoading: {bar:░<100} {i}%')
            sys.stdout.flush()
            time.sleep(0.03)
        constains.done.clear()
    print('\n' * 10)

def ask_user(question) ->bool:
    user_input : str = input(f"\n {question} ? [{apply_color('Y')}/'N]:     ")
    os.system('cls')
    return True if user_input.upper() != 'N' else False
