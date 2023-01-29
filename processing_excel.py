import glob
import pandas as pd
from datetime import date
import numpy as np

today = date.today()
sheetname = 'podzial '+str(today)
file = glob.glob(
    'dir\files*.xlsx')
file.sort

df = pd.read_excel(file[0], sheet_name='Name')
cutoff_data = df.index[df['column1'] == 'data_openings'].tolist()
current_data = df.iloc[:, [
    0,
    11,
    10,
    6,
    9]]
current_data = current_data[0:cutoff_data[0]]
current_data.columns = [
    'data1',
    'data2',
    'data3',
    'data4',
    'data5']

for column in current_data.columns:
    current_data[column] = current_data[column].str.strip()

future_data = df.iloc[:, [
    0,
    11,
    10,
    6,
    9,
    21]]
future_data = future_data[cutoff_data[0]+1:]
future_data.columns = [
    'data1',
    'data2',
    'data3',
    'data4',
    'data5',
    'data6']
empty_data_cells = np.where(pd.isnull(future_data['Data otwarcia']))[0]
future_data = future_data[:empty_data_cells[0]]
future_data['data6'] = pd.to_datetime(
    future_data['data6'], errors='coerce').dt.strftime('%Y-%m-%d')

for column in future_data.columns:
    if column != 'data6':
        future_data[column] = future_data[column].str.strip()
    else:
        pass
writer = pd.ExcelWriter(sheetname+'.xlsx', engine='xlsxwriter')

current_data.to_excel(writer, sheet_name='Current', index=False)
future_data.to_excel(writer, sheet_name='Future', index=False)

writer.save()
