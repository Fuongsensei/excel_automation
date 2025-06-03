#pylint:disable = all

import xlwings as xw
import psutil
import time as t 
import os
import pandas as pd
import datetime as dt 
import win32com.client
from xlsx2csv import Xlsx2csv
import re
if os.name == "nt":
    os.environ.setdefault("TERM", "xterm")
from blessed import Terminal

def clear_sheet_data(wb) -> None:
        sheet = wb.sheets["Verify data"]
        last_row = sheet.range("A" + str((sheet.cells.last_cell).row)).end('up').row
        if last_row > 2:
                sheet.api.Rows(f"3:{last_row}").Delete()
        sheet.range('A2:N2').clear()


def write_df_to_excel(data, des_path:str, callback_1,callback_2) -> None:
        try:
                wb = xw.Book(des_path,update_links=False)
                callback_1(wb)
                sheet = wb.sheets['Verify data']
                sheet.range('A2:N2').value = data.values
                wb.save()
        except Exception:
                print(f'\nLỗi xung đột: File đang mở — sẽ tự đóng.')
        callback_2("EXCEL.EXE")


def close_excel(app_name: str) -> None:
        for proc in psutil.process_iter():
                if proc.name() == app_name:
                        proc.kill()


def call_macro(des_path,user_input):
        wb = xw.Book(des_path)
        macros : dict = {
        '1': wb.macro("Ngay"),
        '2': wb.macro('DEM')
        }
        macro_delete = wb.macro("Xoa_Data")
        macro_run_data = wb.macro('chay_Data')

        macros[user_input]()
        t.sleep(1)
        macro_delete()
        t.sleep(3)
        macro_run_data()

def check_state_file(path:str) -> bool:
        file_name = os.path.basename(path)
        try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                for wb in excel.Workbooks:
                        if wb.Name == file_name:
                                return True
        except Exception:
                pass  # Excel không mở
        return False


def open_file(path:str) -> None:
        if not check_state_file(path):
                os.startfile(path)
                return
        else :
                print("FILE ĐANG ĐƯỢC MỞ ")
                return



def delete_blank(path):
        wb= xw.Book(path,update_links=False)
        sheet = wb.sheets['GRN (10 so)']
        last_rng : int = (sheet.cells.last_cell).row 
        last_row : int = sheet.range(last_rng,1).end('up').row
        sheet.api.Range(f'A{last_row+1}:A{last_rng}').EntireRow.Delete()
        wb.save()
        print('ĐÃ XÓA CÁC DÒNG TRỐNG')


def get_criteria(path:str):
        import  getpass
        csv_path : str =rf"C:\Users\{getpass.getuser()}\Documents\entered_date.csv"
        
        Xlsx2csv(path,outputencoding='utf-8').convert(csv_path,sheetid=3)
        
        entered_str_date : pd.DataFrame = pd.read_csv(csv_path,usecols=[2],dtype=str)
        entered_date = pd.to_datetime(entered_str_date.iloc[:,0],format='%m-%d-%y')
        
        mask : pd.Series[bool]  = entered_date == entered_date.iloc[0]
        
        is_false : bool = mask.all(axis=0)
        if not is_false:
                max : int = mask.idxmax()
                min : int = mask.idxmin()
                if entered_date[max] > entered_date[min]: return dt.datetime.strftime(entered_date[min],format='%m/%d/%Y')
        
                else:  return dt.datetime.strftime(entered_date[max],format='%m/%d/%Y')
                
                
        else : return None

def delete_entered_on_date(path: str,criteria:str)-> None:
        if not criteria  == None:
                wb = xw.Book(path,update_links=False)
                sheet = wb.sheets['GRN (10 so)']
                last_row : int = sheet.range((sheet.cells.last_cell).row,3).end('up').row
                sheet.api.Range('C1').AutoFilter(Field=3, Criteria1=criteria)
                try:
                        visible_row = sheet.api.Range(f'C2:C{last_row}').SpecialCells(12)
                        visible_row.EntireRow.Delete()
                        print("ĐÃ XÓA NGÀY CŨ")
                        

                except Exception as er:
                        print('KHÔNG TÌM THẤY ')
                
                sheet.api.ShowAllData()
        else : print("KHÔNG CÓ NGÀY CŨ ĐỂ XÓA")
        

def delete_na(path)->None:
        wb = xw.Book(path,update_links=False)
        # Xóa #N/A trong sheet 'Label (16 So)'
        sheet=wb.sheets['Label (16 So)']
        last_row : int = sheet.range((sheet.cells.last_cell).row,25).end('up').row
        print(last_row)
        sheet.api.Range('Y1').AutoFilter(Field=25,Criteria1='#N/A')
        visible_rows = sheet.api.Range(f'Y2:Y{last_row}').SpecialCells(12)
        visible_rows.EntireRow.Delete()
        sheet.api.ShowAllData()
        print('ĐÃ XÓA #N/A')



def write_user_to_sheet(data: pd.DataFrame , path: str):
                wb = xw.Book(path,update_links=False)
                sheet=wb.sheets['User']
                table = sheet.api.ListObjects('Table3')
                for i in range(len(data)-1):
                        print(data.iloc[i].values)
                        print('\n'*3)
                last_row : int = sheet.range((sheet.cells.last_cell).row,1).end('up').row
                print('\n'*5)
                if last_row > 4 :
                        sheet.api.Range(f'A{4+1}:A{last_row}').EntireRow.Delete()
                else : pass
                user_input : str = input(f"VUI LÒNG CHỌN SAP KEYIN TỪ {data.index.max()} ĐẾN {data.index.min()}:       ")
                user_input = user_input.strip()
                choose_list : list[int] = [int(i) for i in re.split(r'\D+',user_input)]
                for i , row in enumerate(choose_list):
                                start = 4 + i
                                sheet.range(f'A{start}:H{start}').value = data.iloc[row].values
                rng = sheet.api.Range(f'A3:H{start}')
                table.Resize(rng)        
                wb.save()





