import asyncio
import xlwings as xw
import matplotlib.pyplot as plt
from math import remainder
from typing import final
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from werkzeug.datastructures import FileStorage
import pandas as pd
import matplotlib
from PIL import Image
matplotlib.use('Agg')

UPLOAD_FOLDER = './data'
ALLOWED_EXTENSIONS = {'xlsx'}
# Create a Flask instace
app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

attribute_names = []
attribute_cells = []

city_names = []
city_cells = []
graph_type = ""

for i in range(3, 68):
    city_cells.append(f"B{i}")

for i in 'CDEFGHIJKLMNOP':
    attribute_cells.append(f"{i}2")


@app.route('/')
def index():
    return render_template("main.html", attribute_cells=attribute_cells, attribute_names=attribute_names,
                           attribute_count=len(attribute_names),
                           city_cells=city_cells, city_names=city_names, city_count=len(city_names))


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # print(request.form.getlist('City'))
        # checked_cities = request.form.getlist('City')
        # checked_attributes = request.form.getlist('Attribute')
        # graph_type = request.form.get('graph_type')
        f = request.files['file']
        f.save(secure_filename(f.filename))
        city_names, attribute_names = load_data()
        return render_template("main.html", attribute_cells=attribute_cells, attribute_names=attribute_names,
                               attribute_count=len(attribute_names),
                               city_cells=city_cells, city_names=city_names, city_count=len(city_names))


@app.route('/graph', methods=['GET', 'POST'])
def create_graph():
    if request.method == 'POST':
        checked_attributes = []
        checked_cities = []
        print(request.form.getlist('City'))
        checked_cities = request.form.getlist('City')
        checked_attributes = request.form.getlist('Attribute')
        graph_type = request.form.get('graph_type')
        print(request.form.getlist('Attribute'))
        print(request.form.get('graph_type'))
        city_names, attribute_names = load_data()
        save_graph(graph_type, checked_cities, checked_attributes)
        # checked_attributes.clear()
        # checked_cities.clear()
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


def save_graph(graph_type, checked_cities, checked_attributes):
    wb = xw.Book('DummyData.xlsx')
    total_points = []
    cities = []

    # Viewing available
    # sheets in it
    wks = xw.sheets
    print("Available sheets :\n", wks)

    # Selecting a sheet
    ws = wks[0]

    for city in checked_cities:
        cities.append(ws.range(city).value)
    print(cities)
    print(checked_cities)
    for i in range(0, len(checked_cities)):
        city_points = checked_cities[i][1:]
        total_points.append(
            ws.range(f'{checked_attributes[0][0]}{city_points}').value)
    print(cities)
    print(total_points)
    plt.clf()
    plt.bar(cities, total_points)
    plt.title(f'Cities vs {ws.range(checked_attributes[0]).value}')
    plt.xlabel('Cities')
    plt.ylabel(f'{ws.range(checked_attributes[0]).value}')
    # Dynamically generate whatever name the user sends
    plt.savefig('./static/images/graph.png')
    im = Image.open('./static/images/graph.png')
    im.show()

    # checked_attributes.clear()
    # checked_cities.clear()

    wb.close()
