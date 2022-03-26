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


wb = xw.Book('excel-data/DummyData.xlsx')

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
fiscal_date = "Q3FY18"


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

df.to_csv(f'static/stored-data/{fiscal_date}')

df2 = pd.read_csv(f'static/stored-data/{fiscal_date}')

print("Here it is")
print(df2)
# Selecting a value
# # from the selected sheet
# total_points = ws.range("S3:S4").value
# print("A value in sheet1 :", total_points)

# plt.bar(cities, total_points)
# plt.title('City vs Total Points Earned')
# plt.xlabel('City')
# plt.ylabel('Total Points Earned')
# plt.show()
