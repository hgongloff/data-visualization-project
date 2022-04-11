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


@app.route('/', methods=['GET', 'POST'])
def index():

    if request.method == 'POST':
        print("Here is a post")
        print(request.form.get('year'))
        print(request.form.get('quarter'))
        print(request.form.getlist('Dataframe'))
        print(request.form.getlist('City'))
        print(request.form.getlist('Attribute'))
        quarter = request.form.get('quarter')
        year = request.form.getlist('year')
        uploaded_file = request.files['file']
        if uploaded_file.filename != '':
            print("Call file")
            uploaded_file.save(uploaded_file.filename)
            if int(year[0]) < 10:
                save_data(uploaded_file.filename, f'FY0{year[0]}{quarter}')
            else:
                save_data(uploaded_file.filename, f'FY{year[0]}{quarter}')

        saved_dataframes, excel_names = load_initial_excel_files()
        print(excel_names)

    attribute_names = []
    city_names = []

    if saved_dataframes:
        attribute_names, city_names = load_initial_cities_attributes(saved_dataframes)

    return render_template("main.html", attribute_names=attribute_names,
                           attribute_count=len(attribute_names), city_names=city_names, city_count=len(city_names),
                           excel_names=excel_names, excel_count=len(excel_names))

@app.route('/graph', methods=['GET', 'POST'])
def graph_page():
    if request.method == 'POST':
        print("Here is a post")
        print(request.form.get('year'))
        print(request.form.getlist('quarter'))
        print(request.form.getlist('Dataframe'))
        print(request.form.getlist('City'))
        print(request.form.getlist('Attribute'))



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

# Save Data from recently gotten excel file
def save_data(file_name, fiscal_date):

    wb = xw.Book(f'{file_name}')

    # Viewing available
    # sheets in it
    wks = xw.sheets
    # print("Available sheets :\n", wks)

    # Selecting a sheet
    ws = wks[0]

    city_cells = []
    attribute_cells = []
    cities = []
    attributes = []
    points = [[]]
    total_points = []

    for i in range(3, 68):
        city_cells.append(f"B{i}")

    for i in 'CDEFGHIJKLMNOPRST':
        attribute_cells.append(f"{i}2")

    for city_index in city_cells:
        cities.append(ws.range(city_index).value)

    attributes.append("Cities")
    for attribute_index in attribute_cells:
        attributes.append(ws.range(attribute_index).value)

    df = pd.DataFrame(
        columns=attributes)
    
    for i in range(3, 68):
        df.at[i-3, "Cities"] = ws.range(f'B{i}').value
        for j in 'CDEFGHIJKLMNOPRST':
            df.at[i-3, ws.range(
                f'{j}2').value] = ws.range(f'{j}{i}').value

    # city_values = ['D2']

    # for i in range(0, len(city_cells)):
    #     city_points = city_cells[i][1:]
    #     total_points.append(ws.range(f'{city_values[0][0]}{city_points}').value)

    # print(df)
    # df.set_index("Cities")
    df.to_csv(f'static/stored-data/{fiscal_date}.csv')
    #print(df)
    xl = xw.apps.active.api
    xl.Quit()

    return df

def load_initial_cities_attributes(saved_dataframes):
    _attribute_names = saved_dataframes[0].columns.drop("Unnamed: 0")
    _city_names = saved_dataframes[0].index.values
    return _attribute_names, _city_names

def load_new_excel_file():
    return

# Split up quarter year string
def split_fy(fy_string):
    print(fy_string)
    new_string = fy_string[0:6]

    return new_string


def make_line_graph(df_list, cities, attributes, fiscal_years):

    df_total = pd.concat(df_list)

    print(df_total)

    df_total = df_total.sort_values("FYQ")

    new_att = []
    for att in attributes:
        for city in cities:
            new_att.append(f'{att}: {city}')


    fig, ax = plt.subplots()
    i = 0
    for city, gp in df_total.groupby('Cities'):
        city_att = [f'{city}: CLIP', f'{city}: Total Points Earned']
        gp.plot(x='FYQ', y=attributes, ax=ax, label=city_att)
        if i == 0:
            i = 1
        else:
            i = 0 
        print("did")

    print(df_total)
    
    plt.savefig('./static/images/graph.png', dpi=200, bbox_inches='tight')
    im = Image.open('./static/images/graph.png')
    im.show()

