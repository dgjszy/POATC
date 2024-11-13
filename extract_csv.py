import pandas as pd

file_path = 'Unloading-statistics.xlsx'
xls = pd.ExcelFile(file_path)

sheet_names = xls.sheet_names


#  ([^\s，]+?)(装|出)(([^\s出]*?)(\d+)车，?)+
for sheet_name in sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    # date = df.iloc[]