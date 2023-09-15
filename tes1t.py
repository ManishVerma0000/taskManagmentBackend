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

header = ["Received Time", "Subject", "Sender", "Sender Email Address", 'Body']
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

        try:
            body = message.Body
        except Exception as e:
            sender = "N/A"

        data = [received_time, subject, sender, sender_email_address, body]
        # print(data)
        ws.append(data)


except Exception as e:
    print(e)

excel_file_path = "output.xlsx"
wb.save(excel_file_path)

wb.close()
