#pylint:disable = all
import os
import  shutil, getpass
import datetime as dt
from ui_console import get_list_sap
count = 0 

def copy_file_from_net(sap_list,path_list) -> list[str]:
    global count 
    current_user : str  = getpass.getuser()
    current_time : dt  = dt.date.today()
    short_month : str  = current_time.strftime('%b')
    year : dt = current_time.year
    try:
        for sap in sap_list:
                src : str = rf"D:\AWASE1HCMICAP01\AppsData\GR Ver Report\{short_month} {year}\GR Verification {sap}.xlsx"
                dst : str = rf"C:\Users\{current_user}\Documents\GR Verification {sap}.xlsx"
                is_exits = os.path.exists(rf"C:\Users\{current_user}\Documents")
                if is_exits:
                    shutil.copy(src, dst)
                    path_list.append(dst)
                else:
                    os.makedirs(rf"C:\Users\{getpass.getuser()}\Documents\Report",exist_ok=True)
                    shutil.copy(src, dst)
                    path_list.append(dst)
        return  path_list
    except Exception as e:
        os.system('cls')
        print(f"Không tìm thấy file cho SAP {sap}. Vui lòng kiểm tra lại đường dẫn hoặc số SAP ...")
        if count <=10:
            count+=1
            return copy_file_from_net(get_list_sap(), path_list)
        else:os._exit(0)




