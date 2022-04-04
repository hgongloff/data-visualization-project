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
import glob

UPLOAD_FOLDER = './data'
ALLOWED_EXTENSIONS = {'xlsx'}
# Create a Flask instace
app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

attribute_cells = []

city_cells = []

graph_type = ""

for i in range(3, 68):
    city_cells.append(f"B{i}")

for i in 'CDEFGHIJKLMNOPRST':
    attribute_cells.append(f"{i}2")


@app.route('/')
def index():
    saved_dataframes, excel_names = load_initial_excel_files()

    attribute_names = []
    city_names = []

    if saved_dataframes:
        attribute_names, city_names = load_initial_cities_attributes(saved_dataframes)

    return render_template("main.html", attribute_cells=attribute_cells, attribute_names=attribute_names,
                           attribute_count=len(attribute_names),
                           city_cells=city_cells, city_names=city_names, city_count=len(city_names),
                           excel_names=excel_names, excel_count=len(excel_names))



def load_initial_excel_files():
    _saved_dataframes = []
    _file_names = []

    path = r'static/stored-data'  # use your path
    all_files = glob.glob(path + "/*.csv")
    for filename in all_files:
        df = pd.read_csv(filename, index_col=None, header=0)
        df.set_index("Cities", inplace=True)
        _file_names.append(filename[19:25])
        _saved_dataframes.append(df)

    # print(saved_dataframes)
    # print(file_names)
    return _saved_dataframes, _file_names

def load_initial_cities_attributes(saved_dataframes):
    _attribute_names = saved_dataframes[0].columns.drop("Unnamed: 0")
    _city_names = saved_dataframes[0].index.values
    return _attribute_names, _city_names

def load_new_excel_file():
    return


