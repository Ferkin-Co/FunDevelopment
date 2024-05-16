from openpyxl.styles import PatternFill, Font, Side, Border, numbers, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from io import BytesIO
from typing import Any
import sqlalchemy as sqla
import datetime as dt
import urllib.parse
import pandas as pd
import openpyxl

# CONSTANTS FOR COLUMNS AND SHEETS
# Sheets
ACTIVE_SHEETS = ["Summary", "Loader Rejects PIF Close", "Treatment Complete", "Active Under RWCS",
                 "Active Under SID", "Cont Writeoff not in Aging", "Active"]
ONE_PAGER_SHEETS = ["Placements by Month", "Placements & Performance by Cat"]

# Bold these rows
ROWS_TO_BOLD = ["Summary", "Total", "Treatment", "Collection", "Grand Total",
                "Sub Total of Collection Phases", "Difference"]
# Apply $ to values in these columns
DOLLARS_COL = ["Original Assigned", "Additional Charges", "Total Assigned", "Paid", "Credit",
               "Stop Amt", "Resolved", "$Exhausted", "$Current", "At ARM Active",
               "Total Assigned to ARM - Current & Active $", "Reconciliation", "Dollars", "ARM",
               "SID Aging", "Difference", "AR Total Balance", "Total Open AR", "AmtDueInARM"
               "DueAmtInClientAging"]
# Apply % to values in these columns
PERC_COL = ["Paid %", "Credit %", "Stop Amt %", "Resolved %", "Exhausted %", "Resolved & $Exhuasted %",
            "Active %"]

# Color these headers in OnePager reports
COLOR_OP_SHEET_1 = [5, 12, 14, 15, 16, 21, 22, 23]
COLOR_OP_SHEET_2 = [9, 16, 18, 19, 20, 25, 26, 27]
DROP_SORT_COL = ["Sort", "sort"]

def sql_connection() -> str:  # establish connection to sql server

    sql_server = "armsql"
    sql_database = "CSS_Local"
    driver = "ODBC Driver 17 for SQL Server"
    connect_setup = f'DRIVER={driver};SERVER={sql_server};DATABASE={sql_database};Trusted_Connection=yes;'

    connect_param = urllib.parse.quote_plus(connect_setup)
    connect_to = f"mssql+pyodbc://@{sql_server}/{sql_database}?driver={driver}&odbc_connect={connect_param}"

    return connect_to


class ExcelReport:
    def __init__(self, workbook_name: str):
        self.workbook_name = workbook_name
        self.book = openpyxl.Workbook()

    def save(self):
        self.book.save(self.workbook_name)


class StyleExcel:
    def __init__(self, excel_sheet: Any):
        self.excel_sheet = excel_sheet

    def find_tables(self):
        tables = []

        non_empty_rows = [row for row in self.excel_sheet.iter_rows() if any(cell.value is not None for cell in row)]

        if not non_empty_rows:
            return tables

        current_block = (non_empty_rows[0][0].row, non_empty_rows[0][-1].row)

        for row in non_empty_rows:
            if row[0].row > current_block[1] + 1:
                tables.append(current_block)
                current_block = (row[0].row, row[0].row)

            current_block = (current_block[0], row[0].row)

        tables.append(current_block)
        return tables

    def color_headers(self, sheet_num: list[int], columns: list[list], font_size: list[int] = None):
        blue_fill = PatternFill(start_color = "4C68A2", end_color="4C68A2",
                                fill_type="solid")
        green_fill = PatternFill(start_color="00B050", end_color="00B050",
        fill_type="solid")

        for idx, sheet in enumerate(self.excel_sheet):
            if font_size is None:
                font_num = 11
            else:
                font_num = font_size[idx]

            for cell in self.excel_sheet[idx][sheet_num[idx]]:
                cell.fill = blue_fill
                cell.font = Font(bold=True, color="FFFFFF", size=font_num)

            for col in columns[idx]:
                self.excel_sheet[idx].cell(row=1, column=col).fill = green_fill

    def adjust_col_width(self, width_modifier=0.0):
        for col in self.excel_sheet.columns:
            max_length = 0
            column = get_column_letter(col[0].column)

            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

                adjusted_width = (max_length+width_modifier)
                self.excel_sheet.column_dimensions[column].width = adjusted_width

    def bold_rows(self, bold_values: list[list], font_size: list[int] = None):
        for idx, sheet in enumerate(self.excel_sheet):
            if font_size is None:
                bold_font = Font(bold=True, size=11)
            else:
                bold_font = Font(bold=True, size=font_size[idx])

            for row in self.excel_sheet[idx].iter_rows():
                if idx < len(bold_values):
                    if row[0].value in bold_values[idx]:
                        for cell in row:
                            cell.font = bold_font
                    elif row[0].value in bold_values[idx-1]:
                        for cell in row:
                            cell.font = bold_font

    def bold_column_dollars(self, font_size=11):
        bold_font = Font(bold=True, size=font_size)

        for row in self.excel_sheet.iter_rows():
            for cell in row:
                if cell.value == "Dollars":
                    cell_num = cell.row

                    for cells_below in range(cell_num +1,
                                             self.excel_sheet.max_row + 1):
                        self.excel_sheet['B' + str(cells_below)].font = bold_font

    def font_size8(self):
        font_size = Font(size=8)

        for row in self.excel_sheet.iter_rows():
            for cell in row:
                cell.font = font_size

    def outer_borders(self):
        tables = self.find_tables()

        thin_b = Side(style='thin')

        border = Border(left=thin_b,
                        right=thin_b,
                        top=thin_b,
                        bottom=thin_b)

        for start_row, end_row in tables:
            min_col = min((cell.column for row in self.excel_sheet.iter_rows(
                min_row=start_row, max_row=end_row)
                           for cell in row if cell.value is not None))
            max_col = max((cell.column for row in self.excel_sheet.iter_rows(
                min_row=start_row, max_row=end_row)
                           for cell in row if cell.value is not None))

            for row in self.excel_sheet.iter_riws(min_row=start_row, max_row=end_row, min_col=min_col, max_col=max_col):
                for cell in row:
                    cell.border = Border(top=border.top if cell.row == start_row else None,
                                         bottom=border.bottom if cell.row == end_row else None,
                                         left=border.left if cell.column == min_col else None,
                                         right=border.right if cell.column == max_col else None
                                         )

    def cell_to_percent_currency(self):
        for sheet in self.excel_sheet:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and '$' in cell.value:
                        dollar_cell = cell.value.replace('$', '').replace(',', '')
                        try:
                            cell.value = float(dollar_cell)
                            cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                        except ValueError:
                            pass

        for sheet in self.excel_sheet:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.endswith('%'):
                        percent_cell = cell.value.replace('%', '')
                        try:
                            cell.value = float(percent_cell)
                            cell.number_format = numbers.FORMAT_PERCENTAGE
                        except ValueError:
                            pass

    def center_align_start_date(self, start_row=2):
        align_col = 'B'
        end_row = None

        for row in range(start_row, len(self.excel_sheet[align_col]) + 1):
            if self.excel_sheet[align_col + str(row)].value == 'Accounts':
                end_row = row -2
                break
        for row in range(start_row, end_row + 1):
            self.excel_sheet[align_col + str(row)].alignment = Alignmnet(horizontal='center')

    def insert_image(self, start_row=2, column='', picture=''):
        end_row = None

        for row in range(start_row, len(self.excel_sheet[column]) + 1):
            if self.excel_sheet[column + str(row)].value is None:
                end_row = row + 1
                break
            else:
                end_row = start_row + 3

        png = picture
        my_png = openpyxl.drawing.image.Image(png)
        self.excel_sheet.add_image(my_png, column + str(end_row))

class SQLQuery:
    def __init__(self, connection_string: str):
        self.engine = sqla.create_engine(connection_string)

    def execute_query(self, query: list[str]):
        with self.engine.connect() as conn:
            df = pd.read_sql(query, conn)

        return df

class OPReport:
    def __init__(self, sql_query: SQLQuery):
        self.sql_query = sql_query

    def create_op_canada(self, excel_report: ExcelReport):
        query = ["select * from dbo.tmp_OP_SID_OPCA_T1_S1",
                 "select * from dbo.tmp_OP_SID_OPCA_T1_S2",
                 "select * from dbo.tmp_OP_SID_OPCA_T1_S3",
                 "select * from dbo.tmp_OP_SID_OPCA_T1_S4",
                 "select * from dbo.tmp_OP_SID_OPCA_T1_S5",
                 "select * from dbo.tmp_OP_SID_OPCA_T2_S1",
                 "select * from dbo.tmp_OP_SID_OPCA_T2_S2",
                 "select * from dbo.tmp_OP_SID_OPCA_T2_S3"]
        dataframe_list = []

        for q in query:
            df = self.sql_query.execute_query(query)
            for col in DROP_SORT_COL:
                df = df.drop(columns=col, errors='ignore')





