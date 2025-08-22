import ctypes.wintypes
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.policy import default
from smtplib import SMTP
import os
from webbrowser import get
import getpass4
import smtplib
import yaml
import socket
from constains import yaml_path
import pandas as pd
import datetime 
time = datetime.date.today()
day = time.strftime('%d')
month = time.strftime('%m')
year = time.strftime('%Y')
def get_email_and_password()-> tuple:
    with open(yaml_path,mode='r') as file:
        config = yaml.safe_load(file)
        print(config['email_from'])
        email = config['email_from']
        password = config['password']
    if email == 'EMPTY' and password =='EMPTY':
        email_input = input('Vui lòng nhập email:       ').strip() + '@JABIL.COM'
        password_input = getpass4.getpass('Vui lòng nhập mật khẩu: ').strip()
        config['email_from'] = email_input
        config['password'] =  password_input
        ctypes.windll.kernel32.SetFileAttributesW(yaml_path,0)
        with open(yaml_path,mode='w')  as file:
            yaml.dump(config,file,default_flow_style=False,allow_unicode=True,sort_keys=False)
        return email_input,password_input
    else:return email, password

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
           

my_email = get_email_and_password()[0]
my_password = get_email_and_password()[1]
mail = MIMEMultipart()
mail['From'] = my_email
mail['Subject'] = f'REPORT SCAN VERIFY NGAY {day}-{month}-{year} '
mail['To'] = ', '.join(get_mai_to())
print(mail['To'])
print(mail['Subject'])
file_path = rf"C:\Users\3601183\Desktop\Report Scan Verify Shiftly (RCV).xlsm"
file_name = os.path.basename(file_path)
with open(file_path, 'rb') as attachment:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={file_name}')
    mail.attach(part)

def dispatch_an_email()->None:
    
   
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(my_email,my_password)
        server.send_message(mail)
        print(f"Email sent to {mail['To']} successfully!")

    except smtplib.SMTPAuthenticationError as e:
         print("Lỗi xác thực email hoặc mật khẩu!", e)
    except smtplib.SMTPRecipientsRefused as e:
         print("Địa chỉ nhận không hợp lệ hoặc bị từ chối!", e)
    except smtplib.SMTPSenderRefused as e:
         print("Địa chỉ gửi không hợp lệ!", e)
    except smtplib.SMTPDataError as e:
        print("Lỗi khi gửi nội dung mail!", e)
    except smtplib.SMTPConnectError as e:
        print("Không kết nối được với server SMTP!", e)
    except smtplib.SMTPHeloError as e:
        print("Lỗi chào HELO/EHLO với server!", e)
    except smtplib.SMTPServerDisconnected as e:
        print("Bị ngắt kết nối khỏi server SMTP!", e)
    except smtplib.SMTPException as e:
        print("Lỗi SMTP không xác định!", e)
    except (socket.gaierror, OSError) as e:
        print("Lỗi kết nối mạng hoặc hostname!", e)
    except Exception as e:
        print("Lỗi không xác định khác:", e)
    finally:
       try:
           server.quit()
       except:
           pass



df = pd.read_excel(file_path, sheet_name='Summary',usecols=[3])
df.columns = ['Status']
not_na_df = df['Status'].notna()

filter_df = (df['Status'] != "Verification scan incomplete")

if filter_df.all():
        print(filter_df.all())
        print("All values are 'Scan verification complete'")
        dispatch_an_email()
else: 
        print(filter_df.all())
        print("There are values that are not 'Scan verification complete'")

