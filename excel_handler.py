#pylint:disable = all

import xlwings as xw
import  getpass
import time as t 
import os
import pandas as pd
import datetime as dt 
import win32com.client
from xlsx2csv import Xlsx2csv
from ui_console import print_user_table_clean,save_selected_keyins,pcn
import run_sapgui as rsg

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
        except Exception as e:
                pcn(f'{type(e).__name__} - {e}')
                pcn(f'\nLỗi xung đột: File đang mở — sẽ tự đóng.')
                callback_2(wb,des_path)


def close_excel(wb: xw.Book,path: str) -> None:
        wb.save()
        wb.close()
        t.sleep(4)
        wb_again = xw.Book(path,update_links=False)
        

def call_macro(des_path,user_input):
        wb = xw.Book(des_path,update_links = False)
        year = rsg.get_year()
        posting_date = rsg.get_posting_date(user_input)
        entered_date = rsg.get_entered_date()

        rsg.copy_wdid_user(wb)

        file_10 = rsg.get_file_grn(10)
        file_16 = rsg.get_file_grn(16)

        rsg.auto_sapgui_grn_10(year,file_10['file_name'],file_10['file_path'],posting_date,entered_date)
        rsg.copy_grn_10(file_10['file_path'],wb)

        rsg.auto_sapgui_grn_16(year,file_16['file_name'],entered_date)
        rsg.copy_grn_16(file_16['file_path'],wb)


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
                pcn("  FILE ĐANG ĐƯỢC MỞ  ")
                return



def delete_blank(path):
        wb= xw.Book(path,update_links=False)
        sheet = wb.sheets['GRN (10 so)']
        last_rng : int = (sheet.cells.last_cell).row 
        last_row : int = sheet.range(last_rng,1).end('up').row
        sheet.api.Range(f'A{last_row+1}:A{last_rng}').EntireRow.Delete()
        wb.save()
        pcn('  ĐÃ XÓA CÁC DÒNG TRỐNG  ')


def get_criteria(path:str):
        
        try:        
                csv_path : str =rf"C:\Users\{getpass.getuser()}\Documents\entered_date.csv"
        
                Xlsx2csv(path,outputencoding='utf-8').convert(csv_path,sheetid=3)
        except Exception as e:
                
                print(f'Lỗi không tìm thấy file Documents sau đây sẽ tự tạo - {e}' )
                os.makedirs(rf"C:\Users\{getpass.getuser()}\Documents",exist_ok=True)
                Xlsx2csv(path,outputencoding='utf-8').convert(csv_path,sheetid=3)
                
        entered_str_date : pd.DataFrame = pd.read_csv(csv_path,usecols=[2],dtype=str)
        entered_date = pd.to_datetime(entered_str_date.iloc[:,0],format='%m-%d-%y')
        
        mask : pd.Series[bool]  = entered_date == entered_date.iloc[0]
        
        is_false : bool = mask.all(axis=0)
        if not is_false:
                max : int = mask.idxmax()
                min : int = mask.idxmin()
                if entered_date[max] > entered_date[min]:
                        _date : dt.date = entered_date[min]
                        return f'{_date.month}/{_date.day}/{_date.year}'
        
                else:
                        _date : dt.date = entered_date[max]
                        return f'{_date.month}/{_date.day}/{_date.year}'
                
                
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
                        pcn("  ĐÃ XÓA NGÀY CŨ  ")
                        

                except Exception as er:
                        pcn(' KHÔNG TÌM THẤY NGÀY CŨ ')
                
                sheet.api.ShowAllData()
        else : pcn(" KHÔNG CÓ NGÀY CŨ ĐỂ XÓA ")
        

def delete_na(path)->None:
        wb = xw.Book(path,update_links=False)
        # Xóa #N/A trong sheet 'Label (16 So)'
        sheet=wb.sheets['Label (16 So)']
        last_row : int = sheet.range((sheet.cells.last_cell).row,25).end('up').row
        sheet.api.Range('Y1').AutoFilter(Field=25,Criteria1='#N/A')
        try:        
                visible_rows = sheet.api.Range(f'Y2:Y{last_row}').SpecialCells(12)
                visible_rows.EntireRow.Delete()
                sheet.api.ShowAllData()
                pcn(' ĐÃ XÓA #N/A ')
        except Exception :
                if sheet.api.AutoFilter and sheet.api.FilterMode:        
                        sheet.api.ShowAllData()
                pcn(' Không có #N/A để xóa ')
                return



def write_user_to_sheet(data: pd.DataFrame , path: str):
                wb = xw.Book(path,update_links=False)
                sheet=wb.sheets['User']
                table = sheet.api.ListObjects('Table3')
                last_row : int = sheet.range((sheet.cells.last_cell).row,1).end('up').row
                print_user_table_clean(data)
                choose_list = save_selected_keyins(data)
                if last_row > 4 :
                        sheet.api.Range(f'A{4+1}:A{last_row}').EntireRow.Delete()
                else : pass
                for i , row in enumerate(choose_list):
                                start = 4 + i
                                sheet.range(f'A{start}:H{start}').value = data.iloc[row].values
                rng = sheet.api.Range(f'A3:H{start}')
                table.Resize(rng)        
                wb.save()





