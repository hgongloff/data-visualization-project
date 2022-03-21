import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw


wb = xw.Book('data/DummyData.xlsx')

# Viewing available
# sheets in it
wks = xw.sheets
print("Available sheets :\n", wks)

# Selecting a sheet
ws = wks[0]

city_cells = ['B6', 'B8', 'B10']
cities = []
total_points = []

for city in city_cells:
    cities.append(ws.range(city).value)

city_values = ['D2']

for i in range(0, len(city_cells)):
    city_points = city_cells[i][1:]
    total_points.append(ws.range(f'{city_values[0][0]}{city_points}').value)


# Selecting a value
# # from the selected sheet
# total_points = ws.range("S3:S4").value
# print("A value in sheet1 :", total_points)


plt.bar(cities, total_points)
plt.title('City vs Total Points Earned')
plt.xlabel('City')
plt.ylabel('Total Points Earned')
plt.show()
