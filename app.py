
from flask import Flask,jsonify
import pandas as pd
import win32com.client
from flask_cors import CORS
app = Flask(__name__)
from flask import request
import json

CORS(app)


@app.route("/result")
def login():
  print("api is run")
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
def finduser(resource_id):
    try:
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


@app.route("/readexcel")
def readExcel():
   excel_file_path = 'output.xlsx'
   df = pd.read_excel(excel_file_path)
   target_index = 2
   if 0 <= target_index < len(df):
      print(df.iloc[target_index])
      return 'hey'
   
      



# Import 'user_controller' inside a function when needed
def run_app():
    app.run(debug=True)



if __name__ == "__main__":
    run_app()
