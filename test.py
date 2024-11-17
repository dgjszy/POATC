import pandas as pd
import re

file_path = 'Unloading-statistics.xlsx'
xls = pd.ExcelFile(file_path)

sheet_names = xls.sheet_names
# df = pd.read_excel(xls, sheet_name=sheet_names[-1])
print(sheet_names.index('4.2'))

