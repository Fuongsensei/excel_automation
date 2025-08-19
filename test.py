
import xlwings as xw
import pandas as pd
import yaml
import os
from constains import yaml_path

def get_mai_to()->str:  
      with open(yaml_path, mode='r') as file:
        config = yaml.safe_load(file)
        for i,k in enumerate(config['LEAD_EMAIL']):
              print(f"{i+1}. {k}")
        choice = input("Chọn người nhận (nhập số hoặc 'all' để gửi cho tất cả): ").strip()
        if choice.lower() == 'all':
            return config['LEAD_EMAIL']
        else:
            try:
                index = int(choice) - 1
                if 0 <= index < len(config['LEAD_EMAIL']):
                    print(f"Đã chọn: {config['LEAD_EMAIL'][index]}")
                    return [config['LEAD_EMAIL'][index]]
                else:
                    print("Lựa chọn không hợp lệ. Vui lòng thử lại.")
                    os.system('cls')
                    return get_mai_to()
            except ValueError:
                os.system('cls')
                print("Lựa chọn không hợp lệ. Vui lòng nhập một số.")
                return get_mai_to()
get_mai_to()