
from flask import request, send_file
import os
import openpyxl
import json
from flask import Flask, jsonify
import pandas as pd
import win32com.client
from flask_cors import CORS
app = Flask(__name__)
CORS(app)

# admin routes =>>>>>>>>


@app.route("/result")
def readExcelFile():
    result = pd.read_excel("output.xlsx")

    json_data = result.to_json(orient='records')
    response = jsonify({'data': json_data})
    response.status_code = 200
    return response


@app.route('/loginuser', methods=["POST"])
def loginuser():
    data = request.get_json()
    print(data['email'])
    print(data['password'])
    if (data['email'] == 'admin@skapsindia.com' and data['password'] == 'ADMIN@123'):
        response = jsonify({"message": "user is login successfully.."})
        response.status_code = 200
        return response
    else:
        response = jsonify({"message": "password is incorrect....."})
        response.status_code = 400
        return response


@app.route('/resource/<int:resource_id>', methods=['GET'])
def finduserDetails(resource_id):
    try:
        print(resource_id, 'this is the value of the resource id')
        excel_file_path = 'output.xlsx'
        df = pd.read_excel(excel_file_path)
        target_index = resource_id
        if 0 <= target_index < len(df):
            data = df.iloc[target_index]
            response = jsonify(data.to_json())
            response.status_code = 200
            return response
    except Exception as e:
        response = jsonify({"message": "interal server error"})
        response.status_code = 400
        return response


@app.route("/emailfetch")
def emailfetch():
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    print(messages)


@app.route("/loginUsr", methods=['POST'])
def UserLogin():
    try:
        email = request.get_json()['email']
        password = request.get_json()['password']

        if (not email and not password):
            response = jsonify({"message": "please  enter all the details"})
        else:
            data = pd.read_excel('userdata.xlsx')
            df = pd.DataFrame(data)
            email_data = df[['email']]
            email_exists = email in df['email'].values
            if (not email_exists):
                response = jsonify({"message": "please enter the valid email"})
                response.status_code = 400
                return response
            else:
                correct_password = df.loc[df['email']
                                          == email, 'password'].values[0]
                if (correct_password == password):
                    response = jsonify({"data": request.get_json()})
                    response.status_code = 200
                    return request.get_json()
                else:
                    response = jsonify({"data": "password mistmatch"})
                    response.status_code = 400
                    return response
    except Exception as e:
        response = jsonify({"message": "error occurs"})
        response.status_code = 400
        return response


@app.route("/RegisterUser", methods=['POST'])
def UserRegistrer():
    try:
        name = request.get_json()['email']
        username = request.get_json()['username']
        password = request.get_json()['password']
        if not name or not username or not password:
            response = jsonify({"message": "please enter all the details"})
            return response
        else:
            data = pd.read_excel('userdata.xlsx')
            df = pd.DataFrame(data)
            email_data = df[['email']]
            email_exists = name in df['email'].values
            print(email_exists)
            if (email_exists):
                response = jsonify({"message": "user is already exist"})
                response.status_code = 400
                return response
            else:
                name = request.get_json()['email']
                username = request.get_json()['username']
                password = request.get_json()['password']
                data_to_insert = [{
                    "email": name,
                    "username": username,
                    "password": password
                }]
                new_data = pd.DataFrame(data_to_insert)
                df = pd.concat([df, new_data], ignore_index=True)
                df.to_excel('userdata.xlsx',  index=False)
                response = jsonify({"message": "registration is completed"})
                response.status_code = 200
                return response
    except Exception as e:
        name = request.get_json()['email']
        username = request.get_json()['username']
        password = request.get_json()['password']
        columns = ['email', 'username', 'password']
        df = pd.DataFrame(columns=columns)
        data_to_insert = [{
            "email": name,
            "username": username,
            "password": password
        }]
        new_data = pd.DataFrame(data_to_insert)
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel('userdata.xlsx',  index=False)
        print(e, 'this is the error in the query')
        response = jsonify({"message": "error occurs"})
        response.status_code = 400
        return response


@app.route('/download/<filename>')
def download_file(filename):
    folder_path = os.path.join('static')
    file_path = os.path.join(folder_path, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "File not found", 404


@app.route('/taskAssign', methods=['POST'])
def TaskAssign():
    try:
        body = request.get_json()['body']
        Subject = request.get_json()['Subject']
        Recivied = request.get_json()['Recivied']
        sender = request.get_json()['sender']
        assignedTo = request.get_json()['assignedTo']
        taskid = request.get_json()['taskid']
        if not taskid:
            response = jsonify({"message": "please enter all the details"})
            response.status_code = 400
            return response
        else:
            data = pd.read_excel('AssignedTo.xlsx')
            df = pd.DataFrame(data)
            print(df['taskid'].values)
            print(taskid, 'this is the value of the taskid')
            email_exists = int(taskid) in df['taskid'].values
            print(email_exists)
            if (email_exists):
                response = jsonify(
                    {"message": "this task is already assigned"})
                response.status_code = 400
                return response
            else:
                body = request.get_json()['body']
                Subject = request.get_json()['Subject']
                Recivied = request.get_json()['Recivied']
                sender = request.get_json()['sender']
                assignedTo = request.get_json()['assignedTo']
                taskid = request.get_json()['taskid']
                data_to_insert = [{
                    "body": body,
                    "Subject": Subject,
                    "Recivied": Recivied,
                    "sender": sender,
                    "assignedTo": assignedTo,
                    "taskid": taskid
                }]
                new_data = pd.DataFrame(data_to_insert)
                df = pd.concat([df, new_data], ignore_index=True)
                df.to_excel('AssignedTo.xlsx',  index=False)
                response = jsonify(
                    {"message": "Assigned is complete "})
                response.status_code = 200
                return response
    except Exception as e:
        body = request.get_json()['body']
        Subject = request.get_json()['Subject']
        Recivied = request.get_json()['Recivied']
        sender = request.get_json()['sender']
        assignedTo = request.get_json()['assignedTo']
        taskid = request.get_json()['taskid']
        columns = ['body', 'Subject', 'Recivied',
                   'sender', 'assignedTo', 'taskid']
        df = pd.DataFrame(columns=columns)
        data_to_insert = [{
            "body": body,
            "Subject": Subject,
            "Recivied": Recivied,
            "sender": sender,
            "assignedTo": assignedTo,
            "taskid": taskid,
        }]
        new_data = pd.DataFrame(data_to_insert)
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel('AssignedTo.xlsx',  index=False)
        print(e, 'this is the error in the query')
        response = jsonify({"message": "error occurs"})
        response.status_code = 400
        return response


@app.route("/totalAssign")
def totalAssign():
    try:
        result = pd.read_excel("AssignedTo.xlsx")
        json_data = result.to_json(orient='records')
        response = jsonify({'data': json_data})
        response.status_code = 200
        return response
    except:
        response = jsonify({"message": "error occurs"})
        response.status_code = 400
        return response


# user routes

@app.route("/userAssignedTask")
def useAssingedTask():
    try:
        response = jsonify({"message": "success"})
        response.status_code(200)
        return response
    except Exception as e:
        response = jsonify({"message": "error occurs"})
        response.status_code = 400
        return response


@app.route("/userList")
def addtheuser():
    try:
        data1 = pd.read_excel("userdata.xlsx")
        result = (data1.to_json(orient='records'))
        print(result, 'this is the result')
        response = jsonify(
            {"message": "the list of the user is", "data": result})
        return response
    except Exception as e:
        response = jsonify({"message": "error occurs"})
        response.status_code(400)
        return response


@app.route("/TaskAddByUser")
def taskAddbyUser():
    try:
        TaskTitle = request.get_json()['TaskTitle']
        AssignedTo = request.get_json()['AssignedTo']
        Status = request.get_json()['Status']
        completion = request.get_json()['completion']
        Priority = request.get_json()['Priority']
        StartDate = request.get_json()['StartDate']
        DueDate = request.get_json()['DueDate']
        CompletedDate = request.get_json()['CompletedDate']
        Remarks = request.get_json()['Remarks']
        Description = request.get_json()['Description']
        if not TaskTitle or not AssignedTo or not Status or not completion or not Priority or not StartDate or not DueDate or not CompletedDate or not Remarks or not Description:
            response = jsonify({"message": "please enter all the details"})
            response.status_code = 400
            return response
        else:
            data_to_insert = [{
                "TaskTitle":  TaskTitle,
                "AssignedTo": AssignedTo,
                "Status": Status,
                "completion": completion,
                "Priority": Priority,
                "StartDate": StartDate,
                "DueDate": DueDate,
                "CompletedDate": CompletedDate,
                "Remarks": Remarks,
                "Description": Description
            }]
            new_data = pd.DataFrame(data_to_insert)
            df = pd.concat([df, new_data], ignore_index=True)
            df.to_excel('AssignedTaskDetails.xlsx',  index=False)
            response = jsonify({"message": "Task is completed"})
            response.status_code = 200
            return response
    except Exception as e:
        TaskTitle = request.get_json()['TaskTitle']
        AssignedTo = request.get_json()['AssignedTo']
        Status = request.get_json()['Status']
        completion = request.get_json()['completion']
        Priority = request.get_json()['Priority']
        StartDate = request.get_json()['StartDate']
        DueDate = request.get_json()['DueDate']
        CompletedDate = request.get_json()['CompletedDate']
        Remarks = request.get_json()['Remarks']
        Description = request.get_json()['Description']
        columns = ['TaskTitle', 'AssignedTo', 'Status', 'completion', 'Priority',
                   'StartDate', 'DueDate', 'CompletedDate', 'Remarks', 'Description']
        df = pd.DataFrame(columns=columns)
        data_to_insert = [{
            "TaskTitle":  TaskTitle,
            "AssignedTo": AssignedTo,
            "Status": Status,
            "completion": completion,
            "Priority": Priority,
            "StartDate": StartDate,
            "DueDate": DueDate,
            "CompletedDate": CompletedDate,
            "Remarks": Remarks,
            "Description": Description
        }]
        new_data = pd.DataFrame(data_to_insert)
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel('AssignedTaskDetails.xlsx',  index=False)
        response = jsonify({"message": "error occurs"})
        response.status_code = 400
        return response


def run_app():
    app.run(debug=True)


if __name__ == "__main__":
    run_app()
