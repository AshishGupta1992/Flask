from flask import Flask,request

app=Flask(__name__)


@app.route("/")
def index():
    return "Hello %s" % request.method

@app.route("/bacon",methods=['GET','POST'])
def bacon():
    if request.method=='GET':
        return "You are using GET method"
    else:
        return "You are using POST method"


@app.route('/flask/<int:ash>')
def flask(ash):
    return "hey there %s" %ash


if __name__ == "__main__":
    app.run()
