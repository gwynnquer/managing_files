import pandas as pd
import teradata as td
import openpyxl
from openpyxl.styles import Side
import os
import shutil
import datetime
import calendar
import win32com.client as win32

LAST_REPORING_DATE = (datetime.datetime.now() -
                      datetime.timedelta(days=1)).date()

FIRST_DAY_REPORTING_MONTH = datetime.date(
    LAST_REPORING_DATE.year, LAST_REPORING_DATE.month, 1)

FIRST_DAY_PREVIOUS_MONTH = datetime.date(LAST_REPORING_DATE.year, LAST_REPORING_DATE.month - 1,
                                         1) if LAST_REPORING_DATE.month > 1 else datetime.date(LAST_REPORING_DATE.year - 1, 12, 1)

LAST_DAY_PREVIOUS_MONTH = datetime.date(FIRST_DAY_PREVIOUS_MONTH.year, FIRST_DAY_PREVIOUS_MONTH.month, calendar.monthrange(
    FIRST_DAY_PREVIOUS_MONTH.year, FIRST_DAY_PREVIOUS_MONTH.month)[1])

SHEET_NAMES = ['Current month', 'Previous month']

CURRENT_DATA_INFO = f'{SHEET_NAMES[0]}, dane od {FIRST_DAY_REPORTING_MONTH} do {LAST_REPORING_DATE}'
PREVIOUS_DATA_INFO = f'{SHEET_NAMES[1]}, dane od {FIRST_DAY_PREVIOUS_MONTH} do {LAST_DAY_PREVIOUS_MONTH}'

SAMPLE_FILE = 'eastore_sample.xlsx'

THIN_B = openpyxl.styles.Border(left=Side(style='thin'),
                                right=Side(style='thin'),
                                top=Side(style='thin'),
                                bottom=Side(style='thin'))

DOTTED_B = openpyxl.styles.Border(left=openpyxl.styles.Side(style="dotted"),
                                  right=openpyxl.styles.Side(style="dotted"),
                                  top=openpyxl.styles.Side(style="dotted"),
                                  bottom=openpyxl.styles.Side(style="dotted"))

DARK_BLUE_FONT = openpyxl.styles.Font(color='00008B', bold=True)

BLACK_FONT = openpyxl.styles.Font(color='000000', bold=False)

BLUE_FILL = openpyxl.styles.PatternFill(
    start_color="DAEEF3", end_color="DAEEF3", fill_type="solid")

WHITE_FILL = openpyxl.styles.PatternFill(
    start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")


class ProcessExcel():
    CREATED = 0

    def __init__(self, id, df):
        # id = R022, self = df_R022
        self.file_name = f'{id}.xlsx'
        self.file_dir = f'store\{self.file_name}'
        self.full_dir = rf'some_private_dir\{self.file_dir}'

        self.df_current = df.query(
            'Order_Date >= @FIRST_DAY_REPORTING_MONTH and Data_zamówienia <= @LAST_REPORING_DATE')

        self.df_previous = df.query(
            'Order_Date >= @FIRST_DAY_PREVIOUS_MONTH and Data_zamówienia <= @LAST_DAY_PREVIOUS_MONTH')
        self.update_data()
        self.sent_via_outlook()

    def update_data(self):

        if not os.path.exists(self.file_dir):
            self.create_file()

        wb = openpyxl.load_workbook(self.file_dir)
        ws_current = wb.worksheets[0]
        ws_current['B2'].value = CURRENT_DATA_INFO

        if self.CREATED == 0:
            self.new_rows(ws=ws_current, df=self.df_current,
                          fill=WHITE_FILL, font=BLACK_FONT, border=DOTTED_B)

        self.format_table(ws=ws_current, df=self.df_current,
                          fill=BLUE_FILL, font=DARK_BLUE_FONT, border=THIN_B)

        wb.save(self.file_dir)

        return

    def create_file(self):

        shutil.copy2(SAMPLE_FILE, self.file_dir)
        self.load_previous_month_data()
        self.CREATED = 1

        return

    def load_previous_month_data(self):

        wb = openpyxl.load_workbook(self.file_dir)
        ws_previous = wb.worksheets[1]
        ws_previous['B2'].value = PREVIOUS_DATA_INFO

        self.format_table(ws=ws_previous, df=self.df_previous,
                          fill=WHITE_FILL, font=BLACK_FONT, border=DOTTED_B)

        return

    def new_rows(self, ws, df, fill, font, border):

        max_row = ws.max_row
        cell_range = ws[f"B4:G{max_row}"]

        for row in cell_range:
            for cell in row:
                cell.fill = fill
                cell.border = border
                cell.font = font

        df_rows = len(df)
        ws.insert_rows(idx=4, amount=df_rows)

        return

    def format_table(self, ws, df, fill, font, border):

        for r in range(len(df)):
            if r >= len(df):
                break
            for c in range(len(df.columns)):
                ws.cell(row=r+4, column=c +
                        2).value = df.iloc[r, c]
                ws.cell(row=r+4, column=c+2).border = border
                ws.cell(row=r+4, column=c+2).fill = fill
                ws.cell(row=r+4, column=c+2).font = font

        return

    def sent_via_outlook(self, id=None):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'example.mail@example.com'
        mail.Subject = f'sometitle {LAST_REPORING_DATE}'
        mail.HTMLBody = f'''First sheet shows report from {FIRST_DAY_REPORTING_MONTH} to {LAST_REPORING_DATE}, 
            new data are at the top of the table. 
            Second sheet shows previous data from {FIRST_DAY_PREVIOUS_MONTH} to {LAST_DAY_PREVIOUS_MONTH}. 
            <br> <i>Message has been generated automaticly.</i>'''
        mail.Attachments.Add(self.full_dir)
        mail.Send()

        return


def get_sql_data():
    udaExec = td.UdaExec(appName="name", version="1.0", logConsole=False)
    session = udaExec.connect(
        method="method", system="system", username="user", password="pass")

    query_call = "CALL DATA_BASE.ONLY_NEWEST_DATA()"
    session.execute(query_call)

    query = "SELECT * FROM DATA_BASE.TABLE"
    df = pd.read_sql(query, session)

    session.close()

    processed_df = df.rename(columns={'ORDER_DATE': 'Order_Date'})

    processed_df = processed_df.sort_values(by=['Shop', 'Order_Date'])

    return processed_df


def main():

    new_data = get_sql_data()
    stores = new_data['Shop'].unique()

    for store in stores:
      store_df = new_data[new_data['Shop] == store]
      ProcessExcel(id=shop, df=store_df)


if __name__ == '__main__':
    main()
