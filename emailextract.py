import win32com.client
import os
import openpyxl
from openpyxl import Workbook

# Function to save email attachments to a folder and record file paths in an Excel sheet
def save_attachments_and_record_paths():
    folder_name = "outlook_attachments"
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    
    wb = Workbook()
    ws = wb.active
    header = ["Subject", "Attachment Path"]
    ws.append(header)

    try:
        for message in messages:
            subject = message.Subject
            # Check for attachments
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    filename =os.path.join(os.getcwd(),folder_name,attachment.FileName)
                    file_name = attachment.FileName
                    # file_extension = file_name.split('.')[-1]
                    # print(file_extension)
                    
                    if file_name.endswith(".png") or file_name.endswith(".jpg"):
                        print(file_name)

                    else:
                        print(file_name)
                        if(file_name):
                            try:
                                attachment.SaveAsFile(filename)
                                ws.append([subject, filename])
                            except:
                                ws.append([subject,'null'])


                        
                    # Add the file extension to the set
                    # attachment_types.add(file_extension.lower())
                    # if(filename.find('.png')):
                    #     print(filename)
                  
                    # file_extension = filename.split('.')[-1]
                    # print(file_extension,'this is the value of the extension')
                    
                    # if(filename):
                    #     try:
                    #         attachment.SaveAsFile(filename)
                    #         ws.append([subject, filename])
                    #     except :
                    #         ws.append([subject,'null'])
                    # else:
                    #     print("errr")

                excel_file_path = "attachment_paths.xlsx"
                wb.save(excel_file_path)
                  
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    save_attachments_and_record_paths()
