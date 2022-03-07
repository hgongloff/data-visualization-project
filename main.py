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

cities = []
s = 'B7'
cities.append(ws[s].value)


# Selecting a value
# from the selected sheet
total_points = ws.range("S3:S10").value
print("A value in sheet1 :", total_points)


plt.bar(cities, total_points)
plt.title('City vs Total Points Earned')
plt.xlabel('City')
plt.ylabel('Total Points Earned')
plt.show()
