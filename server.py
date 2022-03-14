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

attribute_names = []
attribute_cells = []

city_names = []
city_cells = []


@app.route('/')
def index():
    return render_template("main.html", attribute_cells=attribute_cells, attribute_names=attribute_names,
                           attribute_count=len(attribute_names),
                           city_cells=city_cells, city_names=city_names, city_count=len(city_names))


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        f.save(secure_filename(f.filename))
        city_names, attribute_names = load_data()
        return render_template("main.html", attribute_cells=attribute_cells, attribute_names=attribute_names,
                               attribute_count=len(attribute_names),
                               city_cells=city_cells, city_names=city_names, city_count=len(city_names))


@app.route('/graph', methods=['GET', 'POST'])
def create_graph():
    if request.method == 'POST':
        print(request.form.getlist('Attribute'))
        save_graph()
        city_names, attribute_names = load_data()
        return render_template("show_graph.html", attribute_cells=attribute_cells, attribute_names=attribute_names,
                               attribute_count=len(attribute_names),
                               city_cells=city_cells, city_names=city_names, city_count=len(city_names))


@app.route('/user/<name>')
def user(name):
    return render_template("user.html", username=name)


def load_data():
    wb = xw.Book('DummyData.xlsx')

    # Viewing available
    # sheets in it
    wks = xw.sheets
    print("Available sheets :\n", wks)

    # Selecting a sheet
    ws = wks[0]

    city_names = ws.range("B3:B67").value

    attribute_names = ws.range("C2:P2").value

    wb.close()

    return city_names, attribute_names


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

    wb.close()
