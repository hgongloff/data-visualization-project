import asyncio
from tabnanny import filename_only
from matplotlib.transforms import Bbox
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


# Read in data from available csv files
def read_saved_data():
    saved_dataframes = []
    file_names = []

    path = r'static/stored-data'  # use your path
    all_files = glob.glob(path + "/*.csv")
    for filename in all_files:
        df = pd.read_csv(filename, index_col=None, header=0)
        df.set_index("Cities", inplace=True)
        file_names.append(filename[filename.find('Q'):])
        saved_dataframes.append(df)

    # print(saved_dataframes)
    # print(file_names)
    return saved_dataframes, file_names


# Save Data from recently gotten excel file
def save_data(file_name, fiscal_date):

    wb = xw.Book(f'excel-data/{file_name}')

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
    df.set_index("Cities")
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
    print(df)

    return df

    # df2 = pd.read_csv(f'static/stored-data/{fiscal_date}')

    # print("Here it is")
    # print(df2)

    # df2 = df2[["Atlanta"]]

    # df2.plot()

    # plt.savefig('./static/images/graph.png')
    # im = Image.open('./static/images/graph.png')
    # im.show()


# Split dataframes into checked cities and attributes
def split_dataframe(dataframes, cities, attributes):
    df = dataframes[0]

    df2 = df.loc[cities, attributes]

    return df2


# Split up quarter year string
def split_fy(fy_string):
    print(fy_string)
    new_string = fy_string[4] + fy_string[5] + fy_string[1]

    return new_string

# Called if there is only one dataframe used


def make_single_bar_graph(df):
    df.plot.bar()

    plt.savefig('./static/images/graph.png', dpi=120, bbox_inches='tight')
    im = Image.open('./static/images/graph.png')
    im.show()


# Called if there is only one dataframe used
def make_single_line_graph(df):
    df.plot()

    plt.savefig('./static/images/graph.png', dpi=120, bbox_inches='tight')
    im = Image.open('./static/images/graph.png')
    im.show()


file_names = []
saved_dataframes = []
saved_dataframes, file_names = read_saved_data()

fiscal_years = []

for name in file_names:
    fiscal_years.append(split_fy(name))

print(fiscal_years)
# print(file_names)
# print(saved_dataframes)

#df = save_data('DummyData.xlsx', '2QFY20')

#print(df.loc["Atlanta", "CLIP"])


df = split_dataframe(saved_dataframes, cities=[
                     "Atlanta", "Baltimore"], attributes=["CLIP", "Total Points Earned"])

make_single_line_graph(df)
