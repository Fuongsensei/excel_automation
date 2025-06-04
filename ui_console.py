#pylint:disable = all
import sys, os, time, re
from colorama import init, Fore
from blessed import Terminal
import getpass
import constains
init()

if os.name == "nt":
    os.environ.setdefault("TERM", "xterm")

term = Terminal()

def print_center_noitice(noitice:str)-> None:
    print(term.center(term.black_on_white(noitice)))

pcn = print_center_noitice
def apply_color(text : str | list ):
    if isinstance(text, str):
        return Fore.GREEN + text + Fore.RESET
    return ''.join([f"|{Fore.GREEN}{char}{Fore.RESET}|" for char in text])


def print_authors():
    msg : str = (' '*35)+'Dev by phuong_nguyen1183 using Python, compiled to C'+(' '*35)
    for c in msg.upper():
        sys.stdout.write(apply_color(c))
        sys.stdout.flush()
        time.sleep(0.015)
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


from rich.console import Console
from rich.table import Table
from rich.text import Text
from rich.padding import Padding

def print_user_table_clean(data):
    console = Console()
    table = Table(
        show_header=True,
        header_style="bold white on black",
        border_style="white", 
        padding=(0, 0)
    )

    table.add_column("STT", justify="center")
    for col in data.columns:
        table.add_column(str(col), justify="center")

    # Hàng trắng đầu tiên
    top_blank = [Text("") for _ in range(len(data.columns) + 1)]
    table.add_row(*top_blank)

    for i, row in data.iterrows():
        styled_row = [Text(str(i), style="black on white")]
        for cell in row:
            value = str(cell) if str(cell) != "nan" else " "
            styled_row.append(Text(value, style="black on white"))
        table.add_row(*styled_row)

        if i < len(data) - 1:
            table.add_row(*top_blank)
    padded_table = Padding(table, (0, 2)) 
    console.print(padded_table)


def  save_selected_keyins(data) -> list[int]:
        from getpass import getuser
        path:str = rf"C:\Users\{getpass.getuser()}\Documents\keyins_list.txt"
        is_exist : bool = os.path.exists(path)
        if not is_exist:
                user_input :str  =input(f'Vui lòng chọn SAP keyins từ {data.index.min()} tới {data.index.max()} ')
                user_input = user_input.strip()
                keyins_list : list[int ]= [int(i) for i in re.split(r'\D+',user_input)]
                with open(path,mode='w') as file:
                        for i in keyins_list:
                                file.write(f'{str(i)}'+'\n')
                        return keyins_list
        else: 
                with open(path,mode='r') as file:
                        data : list [str] = file.readlines()
                        keyins_list_mode_r = [int(i) for i in data]
                        is_reselect = input(f'Danh sách keyins bạn đã chọn lúc nãy {apply_color(keyins_list_mode_r)} bạn có muốn chọn lại ? [{apply_color('Y')}]/N:    ')
                        if is_reselect.upper() != 'Y': return keyins_list_mode_r
                        else : return reselect_keyins_users(path)

                        


def reselect_keyins_users(path:str)-> list[int]:
    user_input = input('Vui lòng nhập SAP keyins mới :   ').strip()
    users = [user.strip() for user in re.split(r'\D+',user_input)]
    data = [int(user) for user in users]
    with open(path,mode='w') as file:
        for i in data:
            file.write(f'{str(i)}'+'\n')
    return data


