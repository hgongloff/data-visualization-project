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
    clear_file = True

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
        if uploaded_file.filename != '' and uploaded_file.filename.endswith('.xlsx'):
            print("Call file")
            uploaded_file.save(uploaded_file.filename)
            if int(year[0]) < 10:
                save_data(uploaded_file.filename, f'FY0{year[0]}{quarter}')
            else:
                save_data(uploaded_file.filename, f'FY{year[0]}{quarter}')

    saved_dataframes, excel_names = load_initial_excel_files()
    print(excel_names)
    uploaded_file = ''

    attribute_names = []
    city_names = []

    if saved_dataframes:
        attribute_names, city_names = load_initial_cities_attributes(saved_dataframes)

    new_excel_names = []
    for name in excel_names:
        new_excel_names.append(name[0:6])
    excel_names = new_excel_names

    return render_template("main.html", attribute_names=attribute_names, clear_file=1,
                           attribute_count=len(attribute_names), city_names=city_names, city_count=len(city_names),
                           excel_names=excel_names, excel_count=len(excel_names))

@app.route('/graph', methods=['GET', 'POST'])
def graph_page():
    fiscal_years = []
    if request.method == 'POST':
        print("Here is a post")
        print(request.form.get('year'))
        print(request.form.getlist('quarter'))
        print(request.form.getlist('Dataframe'))
        print(request.form.getlist('City'))
        print(request.form.getlist('Attribute'))
        print(request.form.getlist('graph_type'))
        print("Here")
        cities = request.form.getlist('City')
        attributes = request.form.getlist('Attribute')
        fiscal_years = request.form.getlist('Dataframe')
        fiscal_years = split_fy(fiscal_years)
        graph_type = request.form.get('graph_type')
        saved_dataframes, excel_names = load_initial_excel_files()

        if len(attributes) and len(cities) and len(fiscal_years):
            df_list = split_dataframe(saved_dataframes, fiscal_years, cities=cities, attributes=attributes)
            attributes.remove('FYQ')
            if graph_type == 'Line':
                print(graph_type)
                make_line_graph(df_list, cities, attributes, fiscal_years)
            elif graph_type == 'Bar':
                make_bar_graph(df_list, cities, attributes, fiscal_years)
            elif graph_type == 'Stacked Bar':
                make_stacked_bar_graph(df_list, cities, attributes, fiscal_years)
            elif graph_type == 'Pie':
                make_pie_chart(df_list, cities, attributes, fiscal_years)

    saved_dataframes, excel_names = load_initial_excel_files()

    if saved_dataframes:
        attribute_names, city_names = load_initial_cities_attributes(saved_dataframes)
    else:
        attribute_names = []
        city_names = []

    new_excel_names = []
    for name in excel_names:
        new_excel_names.append(name[0:6])
    excel_names = new_excel_names

    return render_template("main.html", attribute_names=attribute_names, clear_file=1,
                        attribute_count=len(attribute_names), city_names=city_names, city_count=len(city_names),
                        excel_names=excel_names, excel_count=len(excel_names))



def load_initial_excel_files():
    _saved_dataframes = []
    _file_names = []

    path = r'static/stored-data'  # use your path
    all_files = glob.glob(path + "/*.csv")
    for filename in all_files:
        df = pd.read_csv(filename, index_col=None, header=0)
        f = filename[filename.find('F'):]
        f = f[0:6]
        print(f)
        df['FYQ'] = f
        _file_names.append(filename[filename.find('F'):])
        _saved_dataframes.append(df)

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

    if isinstance(ws.range(f'B1').value, str) and ws.range(f'B1').value.startswith('MEPS'):

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
                val = 0
                if isinstance(ws.range(f'{j}{i}').value, float):
                    val = ws.range(f'{j}{i}').value
                df.at[i-3, ws.range(
                    f'{j}2').value] = val


        df.to_csv(f'static/stored-data/{fiscal_date}.csv')

        xl = xw.apps.active.api
        xl.Quit()

        return df
    else:
        xl = xw.apps.active.api
        xl.Quit()
        return 0

def load_initial_cities_attributes(saved_dataframes):
    _attribute_names = saved_dataframes[0].columns.drop(["Unnamed: 0", "Cities"])
    _city_names = []
    for i in range(65):
        _city_names.append(saved_dataframes[0]['Cities'][i])
    return _attribute_names, _city_names

def load_new_excel_file():
    return

# Split up quarter year string
def split_fy(fy_list):
    new_list = []
    for fy in fy_list:
        new_list.append(fy[0:6])

    return new_list

# Split dataframes into checked cities and attributes
def split_dataframe(dataframes, fiscal_years, cities, attributes):
    df_list = []
    attributes.append('FYQ')
    for df in dataframes:
        df.set_index("Cities", inplace=True)
        df_list.append(df.loc[cities, attributes])
    
    # for i in range(len(fiscal_years)):
    #     df_list[i]['FYQ'] = fiscal_years[i]

    print("Split it")
    print(df_list[0])
    return df_list


def make_line_graph(df_list, cities, attributes, fiscal_years):

    df_total = pd.concat(df_list)

    df_total = df_total[df_total['FYQ'].isin(fiscal_years)]

    df_total = df_total.sort_values("FYQ")


    city_att = []

    city_labels = []

    fig, ax = plt.subplots()
    
    for city, gp in df_total.groupby('Cities'):
        for i in range(len(attributes)):
            city_att.append(f'{city}: {attributes[i]}')
        city_labels.append(city_att)
        city_att = []
     
    i = 0
    for city, gp in df_total.groupby('Cities'):
        gp.plot(x='FYQ', y=attributes, ax=ax, label=city_labels[i])
        i = i + 1


    plt.legend(loc='upper right')
    plt.savefig('./static/images/graph.png', dpi=200, bbox_inches='tight')
    im = Image.open('./static/images/graph.png')
    im.show()


def make_bar_graph(df_list, cities, attributes, fiscal_years):

    df_total = pd.concat(df_list)

    print(df_total)

    df_total = df_total[df_total['FYQ'].isin(fiscal_years)]

    print(df_total)

    df_total = df_total.sort_values("FYQ")

    print("Here it is")
    print(df_total)


    city_att = []

    city_labels = []

    fig, ax = plt.subplots()
    
    for city, gp in df_total.groupby('Cities'):
        for i in range(len(attributes)):
            city_att.append(f'{city}: {attributes[i]}')
        city_labels.append(city_att)
        city_att = []
     
    i = 0
    for city, gp in df_total.groupby('Cities'):
        gp.plot.bar(x='FYQ', y=attributes, label=city_labels[i])
        i = i + 1

    
        plt.savefig('./static/images/graph.png', dpi=200, bbox_inches='tight')
        im = Image.open('./static/images/graph.png')
        im.show()



def make_stacked_bar_graph(df_list, cities, attributes, fiscal_years):

    df_total = pd.concat(df_list)

    print(df_total)

    df_total = df_total[df_total['FYQ'].isin(fiscal_years)]

    print(df_total)

    df_total = df_total.sort_values("FYQ")

    print("Here it is")
    print(df_total)


    city_att = []

    city_labels = []

    fig, ax = plt.subplots()
    
    for city, gp in df_total.groupby('Cities'):
        for i in range(len(attributes)):
            city_att.append(f'{city}: {attributes[i]}')
        city_labels.append(city_att)
        city_att = []
    i = 0
    for city, gp in df_total.groupby('Cities'):
        gp.plot.bar(x='FYQ', y=attributes, label=city_labels[i], stacked=True)
        i = i + 1

    
        plt.savefig('./static/images/graph.png', dpi=200, bbox_inches='tight')
        im = Image.open('./static/images/graph.png')
        im.show()
    



def make_pie_chart(df_list, cities, attributes, fiscal_years):

    df_total = pd.concat(df_list)

    df_total = df_total[df_total['FYQ'].isin(fiscal_years)]

    df_total = df_total.sort_values("FYQ")
    titles=df_total['FYQ'].to_list()
    df_total = df_total.drop('FYQ', 1)


    df_total.T.plot.pie(subplots=True, figsize=(20, 3), legend=False, title=titles)

    
    plt.savefig('./static/images/graph.png', dpi=200)
    im = Image.open('./static/images/graph.png')
    im.show()
