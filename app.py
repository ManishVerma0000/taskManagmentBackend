
from flask import Flask,jsonify
import pandas as pd
import win32com.client
from flask_cors import CORS
app = Flask(__name__)
from flask import request,send_file
import json
import openpyxl
CORS(app)
import os

@app.route("/result")
def readExcelFile():

  result= pd.read_excel("output.xlsx")
  json_data =result.to_json(orient='records')
  response= jsonify({'data':json_data})
  response.status_code=200
  return response


@app.route('/loginuser', methods=["POST"])
def loginuser():
   data= request.get_json()
   print(data['email'])
   print(data['password'])
   if(data['email']=='admin@skapsindia.com'and data['password']=='ADMIN@123'):
      response= jsonify({"message":"user is login successfully.."})
      response.status_code=200
      return response
   else:
      response= jsonify({"message":"password is incorrect....."})
      response.status_code=400
      return response

@app.route('/resource/<int:resource_id>', methods=['GET'])
def finduserDetails(resource_id):
    try:
      print(resource_id,'this is the value of the resource id')
      excel_file_path = 'output.xlsx'
      df = pd.read_excel(excel_file_path)
      target_index = resource_id
      if 0 <= target_index < len(df):
          data=df.iloc[target_index]
          print(data['index'])
          print(data['Received Time'])
          response= jsonify(data.to_json())
          response.status_code=200
          return response
    except Exception as e:
       response=jsonify({"message":"interal server error"})
       response.status_code=400
       return response

@app.route("/emailfetch")
def emailfetch():
  outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
  inbox = outlook.GetDefaultFolder(6)
  messages = inbox.Items
  print(messages)


@app.route("/loginUsr",methods=['POST'])
def UserLogin():
   try:
      email= request.get_json()['email']
      password=request.get_json()['password']

      if( not email and not password):
         response=jsonify({"message":"please  enter all the details"})
      else:
         data= pd.read_excel('userdata.xlsx')
         df = pd.DataFrame(data)
         email_data = df[['email']]
         email_exists = email in df['email'].values
         if( not email_exists):
            response=jsonify({"message":"please enter the valid email"})
            response.status_code=400
            return response
         else:
            correct_password = df.loc[df['email'] == email, 'password'].values[0]
            print(correct_password)
            if(correct_password==password):
               response= jsonify({"data":request.get_json()})
               response.status_code=200
               return request.get_json()
            else:
               response= jsonify({"data":"password mistmatch"})
               response.status_code=400
               return response      
   except Exception as e:
      response=jsonify({"message":"error occurs"})
      response.status_code=400
      return response
   

@app.route("/RegisterUser",methods=['POST'])
def UserRegistrer():
   try:
      name=request.get_json()['email']
      username=request.get_json()['username']     
      password=request.get_json()['password']
      if not name or not username or not password:
         response= jsonify({"message":"please enter all the details"})
         return response
      else:
         data= pd.read_excel('userdata.xlsx')
         df = pd.DataFrame(data)
         email_data = df[['email']]
         email_exists = name in df['email'].values
         if(email_exists):
            response=jsonify({"message":"user is already exist"})
            response.status_code=400
            return response
         else:
              name=request.get_json()['email']
              username=request.get_json()['username']
              password=request.get_json()['password']
              data_to_insert=[{
                  "email":name,
                  "username":username,
                  "password":password
              }]
              new_data = pd.DataFrame(data_to_insert)
              df = pd.concat([df, new_data], ignore_index=True)
              df.to_excel('userdata.xlsx',  index=False) 
              response=jsonify({"message":"registration is completed"}) 
              response.status_code=200
              return response       
   except Exception as e:
      name=request.get_json()['email']
      username=request.get_json()['username']
      password=request.get_json()['password']      
      columns = ['email', 'username', 'password']
      df = pd.DataFrame(columns=columns)
      data_to_insert=[{
         "email":name,
         "username":username,
         "password":password
      }]
      new_data = pd.DataFrame(data_to_insert)
      df = pd.concat([df, new_data], ignore_index=True)
      df.to_excel('userdata.xlsx',  index=False)     
      print(e,'this is the error in the query')
      response=jsonify({"message":"error occurs"})
      response.status_code=400
      return response
   
@app.route('/download/<filename>')
def download_file(filename):
    folder_path = os.path.join('static')
    file_path = os.path.join(folder_path, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "File not found", 404
    


@app.route('/taskAssign')
def TaskAssign(filename):
   try:

      return 'hey'
   except Exception as e:
      response= jsonify({"message":"error occurs"})
      response.status_code=400
   

def run_app():
    app.run(debug=True)



if __name__ == "__main__":
    run_app()
