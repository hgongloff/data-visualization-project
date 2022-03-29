import asyncio
from tabnanny import filename_only
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
        file_names.append(filename[filename.find('Q'):])
        saved_dataframes.append(df)

    print(saved_dataframes)
    print(file_names)
    return saved_dataframes, file_names


# Save Data from recently gotten excel file
def save_data(file_name, fiscal_date):

    wb = xw.Book(f'excel-data/{file_name}')

    # Viewing available
    # sheets in it
    wks = xw.sheets
    print("Available sheets :\n", wks)

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

    for attribute_index in attribute_cells:
        attributes.append(ws.range(attribute_index).value)

    df = pd.DataFrame(index=cities,
                      columns=attributes)

    for i in range(3, 68):
        for j in 'CDEFGHIJKLMNOPRST':
            df.at[ws.range(f'B{i}').value, ws.range(
                f'{j}2').value] = ws.range(f'{j}{i}').value

    # city_values = ['D2']

    # for i in range(0, len(city_cells)):
    #     city_points = city_cells[i][1:]
    #     total_points.append(ws.range(f'{city_values[0][0]}{city_points}').value)

    print(df)

    df.to_csv(f'static/stored-data/{fiscal_date}.csv')

    return df

    # df2 = pd.read_csv(f'static/stored-data/{fiscal_date}')

    # print("Here it is")
    # print(df2)

    # df2 = df2[["Atlanta"]]

    # df2.plot()

    # plt.savefig('./static/images/graph.png')
    # im = Image.open('./static/images/graph.png')
    # im.show()


file_names = []
saved_dataframes = []
#saved_dataframes, file_names = read_saved_data()

print(file_names)
print(saved_dataframes)

save_data('DummyData.xlsx', '4Q20')
