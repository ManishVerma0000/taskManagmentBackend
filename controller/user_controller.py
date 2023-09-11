from app import app
from models.user_model import usermodel


user=usermodel()

@app.route("/login")
def login():
    return user.login("heylloo")
    


app.route("/login")
def login():
    return user.login("heylloo")
    




