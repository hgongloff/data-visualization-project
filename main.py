import asyncio
from hashlib import new
from tabnanny import filename_only
from attr import attr, attrib
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

    return df



# Split dataframes into checked cities and attributes
def split_dataframe(dataframes, fiscal_years, cities, attributes):
    df_list = []
    for df in dataframes:
        df.set_index("Cities", inplace=True)
        df_list.append(df.loc[cities, attributes])
        #print("done")
    
    for i in range(len(df_list)):
        df_list[i]['FYQ'] = fiscal_years[i]
        #print(df_list[i])

    #df2 = df.loc[cities, attributes]

    return df_list


# Split up quarter year string
def split_fy(fy_string):
    print(fy_string)
    new_string = fy_string[4] + fy_string[5] + fy_string[1]

    return new_string

def convert_qfy(qfy):
    new_string = f"Q{qfy[2]}FY{qfy[0]}{qfy[1]}"

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

def make_bar_graph(df_list):
    df1 = df_list[0]
    df2 = df_list[1]

    df1['Key'] = "Q2FY20"
    df2['Key'] = "Q3FY21"

    df3 = pd.concat([df1, df2])

    df_group = df3.groupby(['Cities', 'Key'])

    df_plot = df_group.sum().unstack().plot.bar()
    plt.savefig('./static/images/graph.png', dpi=120, bbox_inches='tight')
    im = Image.open('./static/images/graph.png')
    im.show()

def make_att_label(attributes):
    att_label = []
    for att in attributes:
        att_label.append(f"Baltimor e{att}")

    return att_label


def make_line_graph(df_list, cities, attributes, fiscal_years):
    df1 = df_list[0]
    df2 = df_list[1]

    df_total = pd.concat(df_list)

    print(df_total)

    att_label = make_att_label(attributes)

    df_total = df_total.sort_values("FYQ")

    #df_total = df_total.groupby(['Cities'])

    #df_total.plot(x='FYQ', y=attributes)
    #df_total.pivot(index='FYQ', columns=attributes, values=attributes).plot()
    new_att = []
    for att in attributes:
        for city in cities:
            new_att.append(f'{att}: {city}')


    fig, ax = plt.subplots()
    for city, gp in df_total.groupby('Cities'):
        gp.plot(x='FYQ', y='CLIP', ax=ax, label=f'{city}')

    print(df_total)
    #new_df = df1.loc[cities, attributes]


    #print(df1.iloc[df1[cities[0]]][attributes[0]])
    #print(df1.index[df1['Cities'==cities[0]]])

    # Go through and find the cities needed to get values for graph
    # for i in range(65):
    #     #print(df1.iloc[i].name)
    #     if df1.iloc[i].name in cities:
    #         print(df1.iloc[i].name)
    #         new_df['Cities']
    #         for att in attributes:
    #             print(f"attribute: {att} value: {df1.iloc[1][att]}")


    #print(df1.iloc[1].name)
    #print(df1.iloc[1][attributes[0]])

    # df1['Key'] = "Q2FY20"
    # df2['Key'] = "Q3FY21"

    # df3 = pd.concat([df1, df2])

    # df_group = df3.groupby(['Cities', 'Key'])

    # df_plot = df_group.sum().unstack().plot()
    
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

#temp_df = save_data('DummyData.xlsx', 'Q3FY18')

#print(df.loc["Atlanta", "CLIP"])


df = split_dataframe(saved_dataframes, fiscal_years, cities=[
                     "Atlanta", "Baltimore", "Dallas"], attributes=["CLIP", "Total Points Earned"])



# make_single_bar_graph(df[0])
#make_single_line_graph(df[0])

cities = ["Atlanta", "Baltimore", "Dallas"]
attributes = ["CLIP"]

make_line_graph(df, cities, attributes, fiscal_years)


