#pylint:disable = all
import getpass
from stat import FILE_ATTRIBUTE_HIDDEN
import pandas as pd 
import os
from getpass import getuser
import openpyxl
import ctypes 
import time
import yaml
import sys
from collections import OrderedDict as odict
from ui_console import print_user_table_clean as putc , apply_color
from constains import yaml_path





      

def show_users(callback,title:str)->None:
    with open(yaml_path, 'r', encoding='utf-8') as file:
        data = yaml.safe_load(file)
        callback(pd.DataFrame(data[title]))

def get_user(title:str)->pd.DataFrame:
    with open(yaml_path, 'r', encoding='utf-8') as file:
        data = yaml.safe_load(file)
        df = pd.DataFrame(data[title])
        return df

def add_user()-> None:
    user_input = input('Bạn muốn thêm bao nhiêu người dùng?: ').strip()
    for i in range(int(user_input)):
       ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0)
       with open(yaml_path, 'r', encoding='utf-8') as file:
             config = yaml.safe_load(file)
       user_name = input("Nhập tên người dùng: ").strip().upper()
       user_sap = input("Nhập số SAP: ").strip()
       user_shift = input("Nhập ca làm việc (VD: VN02,VN60,VN61,VN62,VN82,VN83): ").strip().upper()
       role = input("Nhập công việc đảm nhận  (VD: Checker, Data entry,Verify): ").strip().upper()
       gender = input("Nhập giới tính (VD: Nam : Male, Nữ : Female): ").strip().upper()
       data = {
       'NAME': user_name,
       'SAP': user_sap,
       'ROLE': role,
       'SHIFT_WORKING': user_shift,
       'GENDER': gender,
       'LOCATION': "TBS",
       'ON_OFF': 'On duty',
       'SLOC': None
              } 
       try:
           if role == 'DATA ENTRY':

               config['data_entry'].append(data) 
               with open(yaml_path,'w', encoding='utf-8') as file:
                                   yaml.dump(config, file, default_flow_style=False, allow_unicode=True,sort_keys=False)
               ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)

           elif role == 'VERIFY':

               config['verify'].append(data) 
               with open(yaml_path,'w', encoding='utf-8') as file:
                                yaml.dump(config, file, default_flow_style=False, allow_unicode=True,sort_keys=False)
               ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)
           os.system('cls')
       except Exception as e:
             os.system('cls')
             print(f"Error adding user: {e}") 


def remove_user()-> None:
    ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0)
    with open(yaml_path, 'r', encoding='utf-8') as file:
          config = yaml.safe_load(file)
    user_sap = input(f"Nhập {apply_color("SAP")} người dùng cần xóa: ").strip()
    role = input("Nhập công việc đảm nhận  (VD: Checker, Data entry,Verify): ").strip().upper()
    try:
        if role == 'DATA ENTRY':
            config['data_entry'] = [user for user in config['data_entry'] if user['SAP'] != user_sap]
            with open(yaml_path,'w', encoding='utf-8') as file:
                                yaml.dump(config, file, default_flow_style=False, allow_unicode=True,sort_keys=False)
            ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)

        elif role == 'VERIFY':
            config['verify'] = [user for user in config['verify'] if user['SAP'] != user_sap]
            with open(yaml_path,'w', encoding='utf-8') as file:
                      yaml.dump(config, file, default_flow_style=False, allow_unicode=True,sort_keys=False)
            ctypes.windll.kernel32.SetFileAttributesW(yaml_path, 0x02)
        os.system('cls')
    except Exception as e:
          print(f"Error removing user: {e}")
          os.system('cls')

