exchange_dn = "//O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=3820A0C319C14FD8A95C114D4CF019EF-USHA.SHETTY"

# Split the Exchange DN by "/"
parts = exchange_dn.split("/")
cn_part = None

# Find the part that starts with "cn="
for part in parts:
    if part.startswith("cn="):
        cn_part = part
        break

# Extract the CN value after "cn="
if cn_part:
    cn_value = cn_part.split("cn=")[1]
    email_address = cn_value
    print("Extracted Email Address:", email_address)
else:
    print("Unable to extract email address from Exchange DN.")
#/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=12C3E919CF464AF6ADE4D3B5AF5276A7-VRUSHALI.GH
#/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=F790DE9F3698434BA5601C80EE79AD98-RAVIRAJ.JAD
#ravirajsinhj@skapsindia.com
#raviraj.jadeja@skapsindia.com





import win32com.client
import os
import openpyxl
import pytz
current_directory = os.getcwd()
folder_name = "outlook_attachments"



outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

wb = openpyxl.Workbook()
ws = wb.active

header = ["Received Time", "Subject", "Sender", "Sender Email Address"]
ws.append(header)

try:
    for message in messages:
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
        data = [received_time, subject, sender, sender_email_address]
        # print(data)
        ws.append(data)


except Exception as e:
    print(e)

excel_file_path = "output.xlsx"
wb.save(excel_file_path)

wb.close()
       
