from math import remainder
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from werkzeug.datastructures import FileStorage
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw

UPLOAD_FOLDER = './data'
ALLOWED_EXTENSIONS = {'xlsx'}
# Create a Flask instace
app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create a route decorator


@app.route('/')
def index():
    return render_template("main.html")


# @app.route('/uploader')
# def upload_file():
#     return render_template('main.html')


@app.route('/uploaded', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        f.save(secure_filename(f.filename))
        return render_template("main.html")


@app.route('/graph', methods=['POST'])
def create_graph():
    if request.method == 'POST':
        print(request.form.get('City'))
        save_graph()
        return render_template("main.html")


@app.route('/user/<name>')
def user(name):
    return render_template("user.html", username=name)


def save_graph():
    wb = xw.Book('DummyData.xlsx')

    # Viewing available
    # sheets in it
    wks = xw.sheets
    print("Available sheets :\n", wks)

    # Selecting a sheet
    ws = wks[0]

    cities = ws.range("B3:B6").value

    # Selecting a value
    # from the selected sheet
    total_points = ws.range("S3:S6").value
    print("A value in sheet1 :", total_points)

    plt.bar(cities, total_points)
    plt.title('City vs Total Points Earned')
    plt.xlabel('City')
    plt.ylabel('Total Points Earned')
    plt.savefig('./static/images/graph.png')
