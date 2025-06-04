#pylint:disable = all
import os
import  shutil, getpass
import datetime as dt
from ui_console import get_list_sap
import sys
count = 0 

def copy_file_from_net(sap_list) -> list[str]:
    global count 
    current_user : str  = getpass.getuser()
    current_time : dt  = dt.date.today()
    short_month : str  = current_time.strftime('%b')
    year : dt = current_time.year
    list_path : list[str] = []
    try:
        for sap in sap_list:
                src : str = rf"\\AWASE1HCMICAP01\AppsData\GR Ver Report\{short_month} {year}\GR Verification {sap}.xlsx"
                dst : str = rf"C:\Users\{current_user}\Documents\GR Verification {sap}.xlsx"
                shutil.copy(src, dst)
                list_path.append(dst)
        return list_path
    except Exception as e:
        os.system('cls')
        print(f"Không tìm thấy file cho SAP {sap}. Vui lòng kiểm tra lại đường dẫn hoặc số SAP ...")
        if count <=10:
            count+=1
            return copy_file_from_net(get_list_sap())
        else:os._exit(0)




