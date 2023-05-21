#import libraries and report if not installed
try:
    import pandas as pd
except ImportError:
    print('Pandas is not installed')
try:
    import openpyxl
except ImportError:
    print('Openpyxl is not installed')
try:
    import numpy as np
except ImportError:
    print('Numpy is not installed')

# Read the excel file
workbook1 = openpyxl.load_workbook(filename = 'data/supermarkt_sales_1.xlsx')
workbook2 = openpyxl.load_workbook(filename = 'data/supermarkt_sales_2.xlsx')

# Get the data from both files using pandas
sheet1 = workbook1.active
data1 = pd.DataFrame(sheet1.values)

sheet2 = workbook2.active
data2 = pd.DataFrame(sheet2.values)

title_row = 4

# Compare the data
comparison_values = data1.values == data2.values
rows, cols = np.where(comparison_values == False)

#Generate report showing the cells that differ
for row, col in zip(rows, cols):
    print(f'Cells ({row + 1}, {col + 1}) are different')
    print(f'Cell {data1.iat[row, col]} != {data2.iat[row, col]}')
    print(f'Row: {row + 1}, {data1.iat[3, col]} are different')
    
print(data1.head())

