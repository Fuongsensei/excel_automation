#pylint:disable = all
from datetime import datetime, timedelta
import pandas as pd
import threading as th
import msoffcrypto
import io
from xlsx2csv import Xlsx2csv

df_list : list  = []

password : str = 'J@bil2022'

day : datetime = datetime.today().replace(hour=6,minute=0,second=0)
in_day : datetime = day.replace(hour=23,minute=59,second=59)
night : datetime = (day - timedelta(days=1)).replace(hour=18,minute=0,second=0)

def load_data_with_key(path, password):
        try:
                with open(path, 'rb') as file:
                        office_file = msoffcrypto.OfficeFile(file)
                        office_file.load_key(password)
                        decrypted = io.BytesIO()
                        office_file.decrypt(decrypted)
                        df : pd.DataFrame = pd.read_excel(decrypted)
                        df_list.append(df)
        except Exception:
                df = pd.read_excel(path)
                df_list.append(df)

def create_dataframe(dercrypt,list_file_path : list )-> None:
        threads : list[th.Thread] = []
        for path in list_file_path:
                task = th.Thread(target=dercrypt,args=(path,password))
                task.start()
                threads.append(task)
                
        for i in threads: 
                i.join()


def concat_df(list_data:list[pd.DataFrame],callback) -> pd.DataFrame:
        data_concat = pd.concat(list_data,axis=0)
        return callback(data_concat) 


def resize_dataframe(data_frame:pd.DataFrame) -> pd.DataFrame:
        data : pd.DataFrame = data_frame.drop(columns=data_frame.columns[[14,15]])
        return data


def filter_df(data:pd.DataFrame,start_time:datetime,end_time:datetime,callback)-> pd.DataFrame:
        user_and_time_col : pd.Series = data.iloc[:,5].str.split("|").str[1]
        date_col: pd.Series = pd.to_datetime(user_and_time_col,format='%m/%d/%Y %I:%M:%S %p')
        if data.columns[12] == 'Stk Placement':
                bool_col = data.drop(columns=data.columns[12])
        else : 
                print('VỊ TRÍ CỘT BỊ SAI')
                return 
        mask : pd.Series = ((date_col >= start_time) & (date_col <= end_time) & (bool_col.iloc[:,9:].all(axis=1)))
        return callback(data[mask])


def unique_data(data: pd.DataFrame)-> pd.DataFrame:
        data_unique : pd.Series = data.duplicated(subset=data.columns[0],keep='first')
        return data[~data_unique]


