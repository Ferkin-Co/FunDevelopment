from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import QTextOption
import sqlalchemy as sqla
from io import StringIO
import traceback as trb
import urllib.parse
import pandas as pd
import tempfile
import chardet
import glob
import csv
import re
import os


def sql_connection():
    # sql_server = "armsqlprod.database.windows.net"
    # sql_server = "armsqldev"
    # sql_database = "sql-reports-prod"
    # sql_database = ""
    # driver = "ODBC Driver 17 for SQL Server"
    # connect_setup = f'DRIVER={{{driver}}};SERVER={sql_server};DATABASE={sql_database};Authentication=ActiveDirectoryIntegrated;Connection Timeout=30'
    # connect_setup = f'DRIVER={{{driver}}};SERVER={sql_server};DATABASE={sql_database};Trusted_Connection=yes;Connection Timeout=30'

    sql_server = "armsql"
    sql_database = "CSS_Local"
    driver = "ODBC Driver 17 for SQL Server"

    connect_setup = f'DRIVER={{{driver}}};SERVER={sql_server};DATABASE={sql_database};Trusted_Connection=yes;Connection Timeout=30'

    connect_param = urllib.parse.quote_plus(connect_setup)
    connect_to = f"mssql+pyodbc://@{sql_server}/{sql_database}?driver={{{driver}}}&odbc_connect={connect_param}"
    return connect_to


def sql_query(search_date_val):  # execute sql query and store to dataframe
    engine = sqla.create_engine(sql_connection())
    query = f"""SELECT [DBT_CLIENT] AS 'Company',[DBT_NO] AS 'Debtor No',[DBT_CLNT_ACNT1] AS 'Account Number',
    [DBT_COMPANY] AS 'Company Code',[DBT_REFERRAL_AMT] AS 'BalanceDue',[DBT_REFERRAL_DATE] AS 'Referral Date'
     FROM [dbo].[DBT] WHERE DBT_REFERRAL_DATE = '{search_date_val}'
     AND DBT_NDPOST_DATE = '{search_date_val}' AND [DBT_COMPANY] NOT LIKE 'PURGE'
     ORDER BY DBT_REFERRAL_DATE DESC;"""

    # query = f"""SELECT * FROM openquery(css, 'SELECT DBT_CLIENT AS "Company",
    # DBT_NO AS "Debtor No",
    # DBT_CLNT_ACNT1 AS "Account Number",
    # DBT_COMPANY AS "Company Code",
    # DBT_REFERRAL_AMT AS "BalanceDue",
    # TO_CHAR(DBT_REFERRAL_DATE, ''YYYY-MM-DD'') AS "Referral Date"
    # FROM css.DBT
    # WHERE TO_CHAR(DBT_REFERRAL_DATE, ''YYYY-MM-DD'') = ''{search_date_val}''
    #     AND TO_CHAR(DBT_NDPOST_DATE, ''YYYY-MM-DD'') = ''{search_date_val}''
    #     AND DBT_COMPANY NOT LIKE ''PURG%''
    # ORDER BY DBT_REFERRAL_DATE DESC')"""

    with engine.connect() as conn:
        call_sp = sqla.text(query)
        adhoc_df = pd.read_sql(call_sp, conn)

    return adhoc_df




def get_directory_path():
    path_dialog = QFileDialog(None)
    path_dialog.setFileMode(QFileDialog.FileMode.Directory)

    if path_dialog.exec() == QFileDialog.DialogCode.Accepted:
        return path_dialog.selectedFiles()[0]
    return "Path not selected"




def sum_nb_dataframe(newbus):  # Sum newbus balances and invoices
    # start = time.time()

    newbus_dataframe = newbus
    # Reset index so index starts back at 0
    newbus_dataframe.reset_index(drop=True, inplace=True)
    balance_sum = newbus_dataframe["BalanceDue"].sum().round(2)
    invoice_sum = newbus_dataframe["InvoiceAmt"].sum().round(2)
    total_nb_sum = balance_sum + invoice_sum
    newbus_dataframe["Total NewBus Sum"] = total_nb_sum
    # Keep only first instance of total newbus sum, all others become NaN/NULL
    newbus_dataframe.loc[1:, "Total NewBus Sum"] = None

    # print(f"sum_nb took {time.time() - start} seconds")

    return newbus_dataframe


def compare_dataframes(newbus, impact, search_date, user_export_path):  # compare newbus report with impact newbus report

    newbus_dataframe = newbus
    adhoc_dataframe = impact

    # create result dataframe and convert datatypes and get difference
    result_dataframe = pd.concat([newbus_dataframe, adhoc_dataframe["Total AdHoc Sum"]], axis=1)
    result_dataframe["Total NewBus Sum"] = result_dataframe["Total NewBus Sum"].astype(float).round(2)
    result_dataframe["Total AdHoc Sum"] = result_dataframe["Total AdHoc Sum"].astype(float).round(2)


    difference = result_dataframe["Total NewBus Sum"]-result_dataframe["Total AdHoc Sum"]
    round_diff = round(difference, 2)
    result_dataframe["Total Difference"] = round_diff

    if result_dataframe["Total Difference"].iloc[0] != 0:
        difference_accounts = extract_difference_accounts(newbus_dataframe, adhoc_dataframe).getvalue()
        with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', prefix=f"NewBus_Difference_Accounts_{search_date}") as temp:
            temp.write(difference_accounts.encode())
            temp_file = temp.name

        diff_acct_df = pd.read_csv(temp_file)
        export_path = user_export_path

        export_name = os.path.join(export_path, f'NB_Report_Difference_accts {search_date}.csv')
        export_csv = diff_acct_df.to_csv(export_name, index=False)

        os.unlink(temp_file)

    # Check if newbus files are in impact report
    result_dataframe["MATCH"] = result_dataframe["Unique ID"].isin(adhoc_dataframe["Unique ID"])
    # Fill all NaN/NULL/NA with 0
    result_dataframe[['BalanceDue', 'InvoiceAmt']] = result_dataframe[['BalanceDue', 'InvoiceAmt']].fillna(value=0)
    # Drop Unique ID column
    result_dataframe.drop(columns=["Unique ID"], inplace=True)

    return result_dataframe


def extract_difference_accounts(local, impact):
    impact_df = impact
    local_df = local
    # combine invoice and balance columns together and ensure that NaN values are = 0
    # if NaN values are not converted, entire column becomes NaN
    local_df["BalanceDue"] = local["BalanceDue"].fillna(0) + local["InvoiceAmt"].fillna(0)

    # Filter columns by columns
    local_df = local_df[["Unique ID", "BalanceDue"]]
    impact_df = impact_df[["Unique ID", "BalanceDue"]]

    # Aggregate combining sum of all balances related to UniqueID
    filtered_local_df = local_df.groupby(["Unique ID"]).agg({"BalanceDue":"sum"}).reset_index()
    filtered_impact_df = impact_df.groupby(["Unique ID"]).agg({"BalanceDue":"sum"}).reset_index()

    # Round to 2 decimal places and add Location column for local and impact dataframes
    filtered_local_df["BalanceDue"] = filtered_local_df["BalanceDue"].astype(float).round(2)
    filtered_local_df["Location"] = "Local"
    filtered_impact_df["BalanceDue"] = filtered_impact_df["BalanceDue"].astype(float).round(2)
    filtered_impact_df["Location"] = "Impact"

    # Concatenate the dataframes
    combined_df = pd.concat([filtered_impact_df, filtered_local_df])

    # Filter out rows with duplicated 'Unique ID'
    duplicated_df = combined_df

    # Group by 'Unique ID' and 'Location', and calculate the sum of 'BalanceDue'
    grouped_df = duplicated_df.groupby(['Unique ID', 'Location']).agg({'BalanceDue': 'sum'}).reset_index()

    # Pivot the dataframe to have 'Location' as columns
    pivot_df = grouped_df.pivot(index='Unique ID', columns='Location', values='BalanceDue').reset_index()
    pivot_df.fillna(0, inplace=True)

    # Calculate the difference between 'Local' and 'Impact'
    pivot_df['Difference'] = pivot_df['Local'] - pivot_df['Impact']

    # Filter out rows where 'Difference' is not zero
    result_df = pivot_df[pivot_df['Difference'] != 0]

    text_stream = StringIO()
    result_df.to_csv(text_stream, index=False)
    text_stream.seek(0)

    return text_stream


def output_results(local, impact, search_date, user_export_path):  # output
    file_date = search_date

    result_dataframe = compare_dataframes(local, impact, search_date, user_export_path)

    result_dataframe.columns = result_dataframe.columns.str.replace(' ','_')

    export_path = user_export_path

    export_name = os.path.join(export_path, f'NB_Report_RESULT {file_date}.csv')
    export_csv = result_dataframe.to_csv(export_name, index=False)

    return



class CompletionWindow(QDialog):
    def __init__(self, dataframe=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Completed")
        self.ok_button = QPushButton("OK", self)

        self.ok_button.clicked.connect(self.close)

        self.layout = QVBoxLayout(self)
        complete_label = QLabel("Process Completed")
        self.abnormal_label = None
        self.table = None
        try:
            if dataframe is not None:
                self.abnormal_label = QLabel("Abnormal Lines Detected")
                self.table = QTableWidget()
                self.table.setColumnCount(len(dataframe.columns))
                self.table.setRowCount(len(dataframe))
                self.table.setHorizontalHeaderLabels(dataframe.columns)
                for idx in range(len(dataframe.index)):
                    for col in range(len(dataframe.columns)):
                        # self.table.setItem(idx, col, QTableWidgetItem(str(dataframe.iat[idx, col])))
                        item = QTableWidgetItem(str(dataframe.iat[idx, col]))
                        if col == 1:
                            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                        else:
                            item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                        self.table.setItem(idx, col, item)

            self.layout.addWidget(complete_label)
            self.layout.addWidget(self.abnormal_label)
            self.layout.addWidget(self.table)
            self.layout.addWidget(self.ok_button)
        except Exception as e:
            print(f"{e}")

class Worker(QThread):
    progress = pyqtSignal(int)

    def __init__(self, new_bus_option=None, export_path=None, sql_date=None):
        super().__init__(parent=None)
        self.new_bus_option = new_bus_option
        self.export_path = export_path
        self.sql_date = sql_date

    def run(self):
        try:
            self.progress.emit(0)
            for i in range(24):
                self.progress.emit(i)
            impact_df = self.new_bus_option.process_adhoc_report(sql_query(self.sql_date))
            self.progress.emit(25)
            dataframes = self.new_bus_option.set_nb_dataframe(self.new_bus_option.concat_nb_dataframes(), impact_df)
            self.progress.emit(50)
            sum_nb_df = sum_nb_dataframe(dataframes)
            self.progress.emit(75)
            res_df = output_results(sum_nb_df, impact_df, self.sql_date, self.export_path)
            self.progress.emit(80)

            return res_df

        except Exception as e:
            print(f"{e}")

        finally:
            self.progress.emit(100)
            self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(parent=None)
        self.setWindowTitle("Test")
        self.setMinimumSize(600, 400)
        layout = QGridLayout(self)

        self.new_bus_option = None
        self.export_path = None
        self.popup = None
        self.progress_bar = QProgressBar(None)

        self.progress_bar.setRange(0, 100)
        self.progress_bar.hide()

        self.sql_date = QDateEdit(self)
        default_date = QDateTime(2024, 1, 1, 0, 0)
        self.sql_date.setDateTime(default_date)
        self.sql_date.setFixedSize(QSize(170, 30))
        self.sql_date.setDisplayFormat("yyyy-MM-dd")

        nb_button = QPushButton("Select NewBus path")
        exp_button = QPushButton("Select Export path")
        date_button = QPushButton("Select Search Date")
        quit_button = QPushButton("Quit")
        execute_button = QPushButton("Execute")

        nb_button.clicked.connect(self.click_nb_path)
        nb_button.setFixedSize(QSize(120, 30))
        exp_button.clicked.connect(self.click_export_path)
        exp_button.setFixedSize(QSize(120, 30))

        date_button.clicked.connect(self.confirm_date)
        date_button.setFixedSize(QSize(120, 30))

        quit_button.clicked.connect(self.close)
        quit_button.setFixedSize(QSize(100, 30))

        execute_button.clicked.connect(self.execute_nb)
        execute_button.setFixedSize(QSize(100, 30))

        self.path_label_1 = QLabel("Path 1: Not Selected")
        self.path_label_2 = QLabel("Path 2: Not Selected")
        self.search_date_label = QLabel("Enter Search Date")
        self.search_date_label.setWordWrap(True)

        layout.addWidget(self.path_label_1, 0, 0)
        layout.addWidget(nb_button, 0, 1)

        layout.addWidget(self.path_label_2, 1, 0)
        layout.addWidget(exp_button, 1, 1)

        layout.addWidget(self.search_date_label, 2, 0)
        layout.addWidget(self.sql_date, 2, 1)

        layout.addWidget(execute_button, 3, 1)
        layout.addWidget(quit_button, 3, 0)

        central_widget = QWidget()
        central_widget.setLayout(layout)

        self.setCentralWidget(central_widget)


    def confirm_date(self):
        search_date_val = self.sql_date.dateTime()
        date_var = search_date_val.toPyDateTime().date()

        return date_var

    def click_nb_path(self):
        try:
            path = get_directory_path()
            self.path_label_1.setText(f"Path 1: \n{path}")
            self.new_bus_option = NewBusReport(path)
        except Exception as e:
            print(f"{e}")

    def click_export_path(self):
        path = get_directory_path()
        self.path_label_2.setText(f"Path 2: \n{path}")
        self.export_path = path


    def execute_nb(self):
        if self.new_bus_option is not None and self.export_path is not None:
            try:
                self.worker = Worker(new_bus_option=self.new_bus_option,
                                           export_path=self.export_path, sql_date=self.confirm_date())
                self.worker.progress.connect(self.progress_bar.setValue)
                self.worker.start()
                self.progress_bar.show()
                self.worker.finished.connect(self.completion_popup)
            except Exception as e:
                print(f"Error: {e}")

    def completion_popup(self):
        # self.popup = CompletionWindow(self.new_bus_files.get_abnormal_text())
        try:
            abnormal_text_df = self.new_bus_option.get_abnormal_text()
            self.popup = CompletionWindow(abnormal_text_df)
            self.popup.move(self.pos())
            self.popup.show()
            self.progress_bar.hide()
        except Exception as e:
            traceback_info = trb.format_exc()
            print(f"{e}\n{traceback_info}")



class NewBusReport:
    def __init__(self, nb_path):
        self.nb_dataframe = pd.DataFrame()
        self.path = nb_path
        self.completion_popup = CompletionWindow()
        self.abnormal_text_df = None

    def search_nb_directory(self):
        path = self.path
        # Read ONLY files with this naming pattern
        file_patterns = ["**NEW*.csv", "**NB*.csv", "**Missing*.csv"]
        # Store files into list
        combined_file_list = [file for pattern in file_patterns for file in glob.glob(os.path.join(path, pattern))]

        filtered_list = list(set(combined_file_list))  # Dedupe list of files

        # print(f"search_nb_dir took {time.time() - start} seconds")

        return filtered_list

    def read_nb_files(self):
        newbus_list = self.search_nb_directory()

        dataframe_list = []
        abnormal_text = set()

        for file in newbus_list:  # search inside each file for an uneven amount of quotation marks, then appends the text into a list
            filename = file
            raw_data = open(filename, 'rb').read()
            detect_encoding = chardet.detect(raw_data)
            res_encoding = detect_encoding['encoding']
            with open(file, 'r', encoding=res_encoding) as f:
                lines = csv.reader(f)  # reads the csv file and returns each row as a list of fields.
                # If there are 3 columns, each row will be a list of 3 strings
                # the variable lines becomes a reader object, and the for loops re-assigns
                # this reader object to the next iteration of 'f' (file)
                for index, line in enumerate(lines):
                    line_str = ','.join(line)
                    if len(line) > 97:
                        abnormal_text.add((file, index + 1))
                        continue
                    elif re.search(r'",(?!\s")', line_str):  # search for bad line commas
                        abnormal_text.add((file, index + 1))
                        continue
            try:
                # read csv file into dataframe and skip bad lines as they are logged via pre-process check
                nb_dataframe = pd.read_csv(file, encoding=res_encoding, converters={"Account#": str},
                                           on_bad_lines='warn', engine='python')
                if "rec_type" in nb_dataframe.columns:
                    nb_dataframe.drop(columns=["rec_type"], axis=1, inplace=True)
                nb_dataframe.insert(0, "Filename", filename)
                dataframe_list.append(nb_dataframe)
            except Exception as e:
                print(f"An error occurred: {e}")

        if abnormal_text:
            self.abnormal_text_df = pd.DataFrame(abnormal_text, index=None, columns=["Filename", "Line"])

        return dataframe_list

    def concat_nb_dataframes(self):
        concat_df = pd.concat(self.read_nb_files(), axis=0, ignore_index=True)
        return concat_df

    def get_abnormal_text(self):
        if self.abnormal_text_df is not None and not self.abnormal_text_df.empty:
            return self.abnormal_text_df

    def process_adhoc_report(self, impact_df):  # add unique identifier to adhoc report and sums
        adhoc_df = impact_df
        if "Unique ID" not in adhoc_df.columns:
            adhoc_df.insert(2, "Unique ID", adhoc_df["Company"] + '-' + adhoc_df["Account Number"].astype(str))
        total_sum = adhoc_df["BalanceDue"].sum()
        adhoc_df.at[0, "Total AdHoc Sum"] = total_sum
        adhoc_df.loc[1:, "Total AdHoc Sum"] = None
        dataframe = adhoc_df.copy()

        return dataframe

    def set_nb_dataframe(self, local_df, impact):  # format newbus dataframe

        newbus_dataframe = local_df.fillna(0)
        adhoc_df = impact
        adhoc_dataframe = adhoc_df.copy()

        # Format columns in newbus dataframe
        newbus_dataframe.rename(columns={"Account#": "Account Number"}, inplace=True)
        newbus_dataframe = newbus_dataframe[["Filename", "Company", "Account Number",
                                             "LastName", "BalanceDue", "InvoiceAmt"]]
        newbus_dataframe.loc[:, "Company"] = newbus_dataframe["Company"].str.slice(start=0, stop=6)

        # Merge adhoc and newbus dataframe to align indexes
        newbus_dataframe.insert(1, "Unique ID",
                                (newbus_dataframe["Company"] + '-' + newbus_dataframe["Account Number"].astype(str)))

        adhoc_dataframe.rename(columns={"Company": "AdHoc Company", "Account Number": "AdHoc Account Number",
                                        "BalanceDue": "AdHoc BalanceDue"}, inplace=True)
        aligned_df = newbus_dataframe.merge(
            adhoc_dataframe[['AdHoc Company', 'AdHoc Account Number', 'AdHoc BalanceDue']],
            left_on=["Company", "Account Number"], right_on=["AdHoc Company", "AdHoc Account Number"], how='left')
        aligned_df.sort_values(by=["Company", "AdHoc Company"], ascending=True, inplace=True)

        # check for duplicate account numbers (mainly for QUENCH)
        is_duplicate = aligned_df.duplicated(subset=["AdHoc Company", "AdHoc Account Number"])
        # if duplicate, change values to 0 and keep first account number value
        aligned_df.loc[is_duplicate, "AdHoc BalanceDue"] = 0

        return aligned_df


app = QApplication([])
window = MainWindow()
window.show()
app.exec()
