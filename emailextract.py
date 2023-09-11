import win32com.client
import os
import openpyxl
import pytz
import schedule
import time

def emailExtract():
    current_directory = os.getcwd()
folder_name = "outlook_attachments"
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
wb = openpyxl.Workbook()
ws = wb.active
header = ["index","Received Time", "Subject", "Sender", "Sender Email Address","Body"]
ws.append(header)

try: 
    i=0
    for  index, message in enumerate(messages):
        
        
        subject = message.Subject  
        try:
            received_time = message.ReceivedTime    
            received_time = received_time.replace(tzinfo=None)
        except Exception as e:
            received_time = "N/A"  
        try:
            sender = message.SenderName
        except Exception as e:
            sender = "N/A"  
        try:
            sender_email_address = message.SenderEmailAddress 
        except Exception as e:
            sender = "N/A"  

        try:
            Body = message.Body 
        except Exception as e:
            sender = "N/A"  
        
        data = [index,received_time, subject, sender, sender_email_address,Body]
        print("api is hit")
        ws.append(data)


except Exception as e:
    print(e)
excel_file_path = "output.xlsx"
wb.save(excel_file_path)
wb.close()


       
schedule.every(1).minutes.do(emailExtract)  # Set the time to 2:30 PM
while True:
    schedule.run_pending()
    time.sleep(1)