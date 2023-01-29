import glob
import pandas as pd
from datetime import date
import numpy as np


def get_file():
    file = glob.glob(
        'dir\files*.xlsx')
    file.sort
    df = pd.read_excel(file[0], sheet_name='name')
    return df


def strip_columns_ex(excel_sheet):
    for column in excel_sheet.columns:
        excel_sheet[column] = excel_sheet[column].str.strip()
    return excel_sheet


def get_current_stores(raw_excel_data, future_opening_cell):
    current_stores_ex = raw_excel_data.iloc[:, [
        0,
        11,
        10,
        6,
        9]]
    current_stores_ex = current_stores_ex[0:future_opening_cell]

    current_stores_ex.columns = [
        'column_1',
        'column_2',
        'column_3',
        'column_4',
        'column_5']

    current_stores_ex = strip_columns_ex(current_stores_ex)
    return current_stores_ex


def get_future_stores(raw_excel_data, future_opening_cell):
    future_stores_ex = raw_excel_data.iloc[:, [
        0,
        11,
        10,
        6,
        9,
        21]]

    future_stores_ex = future_stores_ex[future_opening_cell+1:]
    future_stores_ex.columns = [
        'column_1',
        'column_2',
        'column_3',
        'column_4',
        'column_5',
        'column_6']

    empty_data_cells = np.where(
        pd.isnull(future_stores_ex['column_6']))[0]

    future_stores_ex = future_stores_ex[:empty_data_cells[0]]
    future_stores_ex['column_6'] = pd.to_datetime(
        future_stores_ex['column_6'], errors='coerce').dt.strftime('%Y-%m-%d')
    future_stores_ex = strip_columns_ex(future_stores_ex)
    return future_stores_ex


def write_excel(current_stores, future_stores):
    today = date.today()
    sheet_name = f'name_{today}'
    writer = pd.ExcelWriter(sheet_name+'.xlsx', engine='xlsxwriter')
    current_stores.to_excel(writer, sheet_ncolumn_4e='Current', index=False)
    future_stores.to_excel(writer, sheet_ncolumn_4e='Future', index=False)
    writer.save()


def main():
    main_file = get_file()
    future_opening_possition = main_file.index[main_file['column_1']
                                               == 'Future'].tolist()
    current_stores = get_current_stores(main_file, future_opening_possition[0])
    future_stores = get_future_stores(main_file, future_opening_possition[0])
    write_excel(current_stores, future_stores)


if __name__ == '__main__':
    main()
