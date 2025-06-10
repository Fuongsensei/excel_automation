#pylint:disable = all
import getpass
import pandas as pd 
import os
from getpass import getuser
all_users_list : list = [{
                                'SAP': '2824206',
                                'NAME': 'NGUYỄN NGỌC TRANG',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN02',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None 
                        },


                        {
                                'SAP': '3349561',
                                'NAME': 'VŨ THỊ HÒE',
                                'GENDER': 'Female',
                                'ROLE': 'Data Entry',
                                'SHIFT WORKING': 'VN02',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '3322285',
                                'NAME': 'NGUYỄN THỊ THANH TÂM',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN02',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '3212044',
                                'NAME': 'TUYẾT LAN',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN02',
                                'LOCATION':'TMS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '3294389',
                                'NAME': 'VÂN ANH',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN02',
                                'LOCATION':'TMS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '691461',
                                'NAME': 'HÀ THỊ THANH VÂN',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN60',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '3307871',
                                'NAME': 'TRẦN THU THẢO',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN61',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '2495170',
                                'NAME': 'NGUYỄN THỊ KIM NGÂN',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN62',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '1080662',
                                'NAME': 'LÊ THỊ THÚY HÀ',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN82',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '2978308',
                                'NAME': 'TRƯƠNG NGỌC MAI DUNG',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN82',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '3307814',
                                'NAME': 'NGUYỄN THỊ KIỀU NƯƠNG',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN82',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '1216105',
                                'NAME': 'PHẠM THỤY THANH TRÚC',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN83',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '3225301',
                                'NAME': 'LÊ THỊ HIẾU',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN83',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None
                        },


                        {
                                'SAP': '3225303',
                                'NAME': 'LÊ CAO THÙY LINH',
                                'GENDER': 'Female',
                                'ROLE': 'Data entry',
                                'SHIFT WORKING': 'VN83',
                                'LOCATION':'TBS',
                                'ON/OFF':'On duty',
                                'Sloc': None 
                        }]


def get_users(users_list_of_dict : list[dict])->pd.DataFrame:
        path :str = rf"C:\Users\{getpass.getuser()}\Documents\user.csv"
        is_exist:bool = os.path.exists(path)
        if is_exist:
                data = pd.read_csv(path)
        
                return data
        else :
                data = pd.DataFrame(all_users_list)
                data.to_csv(path,index=False, encoding='utf-8-sig')
        
                return data







