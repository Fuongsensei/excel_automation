#pylint:disable = all
import sys, os, time, re
from colorama import init, Fore
import getpass
import constains
from rich.console import Console
from rich.table import Table
from rich.padding import Padding
from rich.align import Align
from rich.text import Text
from rich.style import Style
from pandas import DataFrame
from constains import yaml_path
import yaml

import ctypes


init(autoreset=True)  


console = Console()

def print_center_notice(notice: str) -> None:
    # Tạo style: chữ đen trên nền trắng
    style = Style(color="black", bgcolor="white")
    # Tạo đối tượng Text có style
    styled_text = Text(notice, style=style)
    # Căn giữa và in ra
    console.print(Align.center(styled_text))


pcn = print_center_notice

def apply_color(text : str | list ) -> str|list:
    if isinstance(text, str):
        return Fore.GREEN + text + Fore.RESET
    return ''.join([f"|{Fore.GREEN}{char}{Fore.RESET}|" for char in text])


def print_authors()->None:
    msg : str = (' '*35)+'Dev by phuong_nguyen_1183 using Python, compiled to C'+(' '*35)
    for c in msg.upper():
        sys.stdout.write(apply_color(c))
        sys.stdout.flush()
        time.sleep(0.015)
    print('\n'*2)


def get_des_path(callback)->str:
    ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0)
    with open(yaml_path, 'r') as file:
        config = yaml.safe_load(file)
        path_file = config['path_report']
        if path_file is None:
           new_path : str = input(f"Dán đường dẫn file {apply_color('Report Scan Verify Shiftly (RCV)')}: ")
           config['path_report'] = new_path
           with open(yaml_path, 'w') as file:
                yaml.dump(config, file, default_flow_style=False, allow_unicode=True,sort_keys=False)
           os.system('cls'); return new_path.strip('"')
        else:
          user_input :str = input(f"Đường dẫn hiện tại là {apply_color(path_file)} — Nhấn Enter để xác nhận, N thay đường dẫn: ").strip()
          config['path_report'] = path_file if user_input.upper() != 'N' else callback()
          with open(yaml_path,mode='w') as file:
                yaml.dump(config, file, default_flow_style=False, allow_unicode=True,sort_keys=False);ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)
                os.system('cls');return config['path_report'].strip('"')
    



def change_des_path()->str:
    new_path : str = input("Nhập đường dẫn mới: ")
    return new_path.strip('"')
   


def get_list_sap()-> list[str]:
    from users_process import add_user , remove_user, show_users
    ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0)
    with open(yaml_path, 'r', encoding='utf-8') as file:
          config = yaml.safe_load(file)
    data = DataFrame(config['verify'])
    user_input = input(f'Danh sách verify {apply_color(config['selected_verify'])} bạn có muốn chọn lại [{apply_color('Y')}/N]? : ') 
    
    if user_input.upper() == 'Y'or config['selected_verify'][0] =='EMPTY' :
       try:
          print_user_table_clean(data)
          sap_input: str  = input("Nhập số verify cách nhau bởi ký tự không phải số: ").strip()
          os.system('cls')
          selected : list = [int(s.strip()) for s in re.split(r'\D+', sap_input)]
          verify_sap = data['SAP'].iloc[selected].tolist()
          config['selected_verify'] = verify_sap
          confirm : str = input(f"Danh sách số: {apply_color(verify_sap)} — Nhấn Enter để xác nhận, N để nhập lại: ")
          os.system('cls')
          with open(yaml_path, 'w', encoding='utf-8') as file:
              yaml.dump(config, file, default_flow_style=False, allow_unicode=True, sort_keys=False)
          return verify_sap if confirm.upper() != 'N' else get_list_sap();ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)
       except Exception as e:
          if constains.count_get_sap==0:
              print('Thân chưa mà giỡn dữ vậy thoát app nhé')
              time.sleep(2)
              os._exit(0)
              print(f'Nhập đàng hoàng đi bro !{e}' )
              time.sleep(2)
              os.system('cls')
              constains.count_get_sap -= 1
              return get_list_sap()
    elif user_input.upper().strip() == 'ADD':
          os.system('cls')
          add_user()
          return get_list_sap()
    elif user_input.upper().strip() == 'REMOVE':
          os.system('cls')
          remove_user()
          return get_list_sap()
    elif user_input.upper().strip() == 'SHOW':
          os.system('cls')
          show_users(print_user_table_clean,'data_entry')
          show_users(print_user_table_clean,'verify')
          return get_list_sap()
    else:
          os.system('cls');return config['selected_verify'];ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)
       
    

def print_loading()->None:
    
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
    user_input : str = input(f"\n {question}  [{apply_color('Y')}/'N]:     ")
    os.system('cls')
    return True if user_input.upper() != 'N' else False




def print_user_table_clean(data:DataFrame)->None:
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


def  save_selected_keyins(path:str) -> list[int]:
        
        ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0)
        with open(path,mode='r') as file:
                 config = yaml.safe_load(file)
        data: DataFrame = DataFrame(config['data_entry'])
        print()
        print_user_table_clean(data)
        select =  input(f'Các số SAP keyins {apply_color(config['selected_keyins'])} đang được chọn bạn có muốn chọn lại [{apply_color('Y')}/N]? : ').strip()
        if select.upper() == 'Y':
            user_input = input("Nhập các số keyins cách nhau bởi ký tự không phải số: ").strip()
            keyins_list = [int(user.strip()) for user in re.split(r'\D+', user_input)]
            config['selected_keyins'] = data['SAP'][keyins_list].tolist()
            with open (path, mode='w', encoding='utf-8') as file:
                      yaml.dump(config, file, default_flow_style=False, allow_unicode=True,sort_keys=False)
                      ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)
                      return keyins_list
        else:
             ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)
             return config['selected_keyins']
        

save_selected_keyins(yaml_path)

