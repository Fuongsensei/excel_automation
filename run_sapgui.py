import win32com.client
import time
import datetime as dt
import xlwings as xw 
from datetime import timedelta
import re 
import os
import pandas as pd

year  = dt.date.today().year
wb = xw.Book(r"C:\Users\3601183\Desktop\Report Scan Verify Shiftly (RCV).xlsm",update_links=False)

sheet = wb.sheets['User']
sheet.api.Range('A3:A30').Copy()


def get_file_grn(grn:int) -> str:
        today : dt  = dt.date.today()
        path :str = r'C:\TEMP'
        suffix :str = '.csv'
        file_name :str  =  f"export_{today.month}_{today.day}_{today.year}_[{grn} nums]"
        file_path :str = os.path.join(path,file_name)
        file_path = file_path+suffix
        return {
                    'file_path' :file_path,
                    'file_name' : file_name
        } 



def get_posting_date(user_input:str) -> str:
    today : dt = dt.date.today()
    day : dt = dt.date.strftime(today,format='%m/%d/%Y')
    night = dt.date.strftime((today - timedelta(days=1)),format='%m/%d/%Y')
    posting_date = {
    'start' : day if user_input == '1' else night,
    'end' : day
    }
    return posting_date


def get_entered_date():
    today : dt = dt.date.today()
    entered_date = {
        
        'start' : dt.date.strftime((today - timedelta(days=3)),format='%m/%d/%Y'),
        'end' : posting_date['end']}
    return entered_date

posting_date = get_posting_date('2')
print(posting_date['end'])
entered_date = get_entered_date()
print(entered_date['end'])



def get_session_sap() -> None :
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not SapGuiAuto:
                print("SAP GUI is not running.")
                exit()
        else:
                application = SapGuiAuto.GetScriptingEngine
                connection = application.Children(0)  
                session = connection.Children(0)
        return session

def run_session_sap(f):
    try:
        session =  f()
        return session
    except Exception as e:
        print('Vui lòng kiểm tra SAP đã bật hay chưa')
        os._exit(0)

session  = run_session_sap(get_session_sap)
file_name_grn_10 = get_file_grn(10)
file_name_grn_16 = get_file_grn(16)
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nz_invmvmts"
session.findById("wnd[0]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").Text = "VN01"
session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = posting_date['start']
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = posting_date['end']
session.findById("wnd[0]/usr/btn%_SO_MJAHR_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").Text =year
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").Text = year 
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtSO_CPUDT-LOW").Text = entered_date['start']
session.findById("wnd[0]/usr/ctxtSO_CPUDT-HIGH").Text = entered_date['end']
session.findById("wnd[0]/usr/ctxtSO_BWART-LOW").Text = "101"
session.findById("wnd[0]/usr/ctxtSO_BWART-HIGH").Text = "102"
session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").SetFocus()
session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_SO_UNAME_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtP_LAY01").Text = "/QUANTITY"
session.findById("wnd[0]/usr/ctxtP_LAY01").SetFocus()
session.findById("wnd[0]/usr/ctxtP_LAY01").caretPosition = 9
session.findById("wnd[0]/usr/ctxtSO_AUFNR-LOW").SetFocus()
session.findById("wnd[0]/usr/ctxtSO_AUFNR-LOW").caretPosition = 1
session.findById("wnd[0]/tbar[1]/btn[8]").press()
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").Text = file_name_grn_10['file_name']
session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(file_name_grn_10['file_name'])
session.findById("wnd[1]/tbar[0]/btn[20]").press()
session.findById("wnd[1]/tbar[0]/btn[11]").press()
time.sleep(3)
wb_export_grn_10 = xw.Book(file_name_grn_10['file_path'])
wb_export_grn_10.close()
data_grn_10 = pd.read_csv(file_name_grn_10['file_path'])

data_grn_10[['User Name','Material Document']] = data_grn_10[['User Name','Material Document']].astype(str)

data_grn_10 = data_grn_10.drop(columns=data_grn_10.columns[[98,94]])
data_grn_10['Network'] = "=VLOOKUP(@CN:CN,'Vendor Subcontrac'!A:B,2,0)"
data_grn_10 = data_grn_10.rename(columns={'Activity': "=COUNTIF(CR:CR,'Vendor Subcontrac'!B2)"})
print(data_grn_10)

wb = xw.Book(r"C:\Users\3601183\Desktop\Report Scan Verify Shiftly (RCV).xlsm",update_links=False)
sheet_grn_10 = wb.sheets['GRN (10 so)']
for i in ['A','H','I']:
    sheet_grn_10.range(f'{i}:{i}').number_format = '@'
sheet_grn_10.range('A1').value = data_grn_10.columns.tolist()  
sheet_grn_10.range('A2').value = data_grn_10.values

print(data_grn_10.index.max())
last_row = data_grn_10.index.max()
sheet_grn_10.api.Range(f'A2:A{last_row+2}').Copy()

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzlgrns1"
session.findById("wnd[0]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "VN01"
session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = entered_date['start']
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = entered_date['end']
session.findById("wnd[0]/usr/ctxtS_BWART-LOW").SetFocus()
session.findById("wnd[0]/usr/ctxtS_BWART-LOW").caretPosition = 3
session.findById("wnd[0]/usr/btn%_S_MBLNR_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[16]").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/CUONG1"
session.findById("wnd[0]/usr/ctxtP_LAYOUT").SetFocus()
session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 7
session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").SetFocus()
session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").caretPosition = 1
session.findById("wnd[0]/tbar[1]/btn[8]").press()
session.findById("wnd[0]/usr/cntlCNTNR/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlCNTNR/shellcont/shell").selectContextMenuItem ("&XXL")
session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").Text = file_name_grn_16['file_name']
session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(file_name_grn_16['file_name'])
session.findById("wnd[1]/tbar[0]/btn[20]").press()
session.findById("wnd[1]/tbar[0]/btn[11]").press()

time.sleep(5)
wb_export_grn_16 = xw.Book(file_name_grn_16['file_path'])
wb_export_grn_16.close()
data_grn_16 = pd.read_csv(file_name_grn_16['file_path'])
data_grn_16 = data_grn_16.drop(columns=data_grn_16.columns[23])
data_grn_16['GRN Number'] = data_grn_16['GRN Number'].astype(str)
sheet_grn_16 = wb.sheets['Label (16 So)']
sheet_grn_16.range('F:F').number_format = '@'
sheet_grn_16.range('A1').value = data_grn_16.columns.tolist() 
sheet_grn_16.range('A2').value = data_grn_16.values      
