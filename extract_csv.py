import pandas as pd
import re

file_path = 'Unloading-statistics.xlsx'
xls = pd.ExcelFile(file_path)

sheet_names = xls.sheet_names

data = {
    '日期':[],
    '队组':[],
    '班次':[],
    '物料':[],
    '数量':[],
    '出发点':[],
    '目的地':[]
}

# 5.9表之后与5.9及之前格式不一样，还有4.2
diff_sheet_index_1 = 26
diff_sheet_index_2 = 63

for index, sheet_name in enumerate(sheet_names):
    if index <= diff_sheet_index_2:
        date_index = (0, 4) #日期位置
        mate_info_index = 4
        mate_num_index = 3
        if index <= diff_sheet_index_1:
            loc_index = 5
        else:
            loc_index = 6
    else:
        date_index = (0, 3) #日期位置
        mate_info_index = 3
        mate_num_index = 2
        loc_index = 5
    # 按顺序读取每一张表格
    df = pd.read_excel(xls, sheet_name=sheet_name)
    # 从表格获得日期信息
    date = df.iloc[date_index[0],date_index[1]].split('：')[-1]
    if date[-1] == '日':
        date.rstrip('日')
    date = date.replace('年', '.')
    date = date.replace('月', '.')
    # 早班、中班与夜班的索引
    index_morning = df.iloc[:, 0].index[df.iloc[:, 0] == '早班'][0]
    index_afternoon = df.iloc[:, 0].index[df.iloc[:, 0] == '中班'][0]
    index_night = df.iloc[:, 0].index[df.iloc[:, 0] == '夜班'][0]

    for i in range(index_morning, df.shape[0]):
        if index_morning <= i < index_afternoon:
            shift = '早班'
        elif index_afternoon <= i < index_night:
            shift = '中班'
        elif i >= index_night:
            shift = '夜班'

        if not pd.isna(df.iloc[i,mate_info_index]):
            text = df.iloc[i,mate_info_index]
            matches2 = re.findall(r'([^\s，]+?)(装|出)(((([^\s出，]+?)(\d+)车，?)+)|(((\d+)车([^\s出，]+)，?)+))', text)
            if not pd.isna(df.iloc[i, mate_num_index]):
                # 运入
                matches = re.findall(r'([^\s，]+?)(\d+)车(，|、)?', text)
                for j in range(0, len(matches)):
                    data['日期'].append(date)
                    data['队组'].append(df.iloc[i,1])
                    data['班次'].append(shift)
                    data['物料'].append(matches[j][0])
                    data['数量'].append(matches[j][1])
                    data['出发点'].append('井口')
                    data['目的地'].append(df.iloc[i,loc_index])
            elif  matches2:
                # 运出
                matches = matches2
                for j in range(0, len(matches)):
                    match_material = re.findall(r'([^\s，计、）]+?)(\d+)车(，|、)?', matches[j][2])
                    for k in range(0, len(match_material)):
                        data['日期'].append(date)
                        data['队组'].append(df.iloc[i,1])
                        data['班次'].append(shift)
                        data['物料'].append(match_material[k][0])
                        data['数量'].append(match_material[k][1])
                        data['出发点'].append(matches[j][0])
                        data['目的地'].append('井口')

cleaned_df = pd.DataFrame(data)
ouput_csv_path = 'cleaned_data.csv'
cleaned_df.to_csv(ouput_csv_path, index=False)