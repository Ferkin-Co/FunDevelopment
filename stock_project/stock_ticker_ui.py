from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (QWidget, QApplication, QLineEdit, QLabel,
                             QPushButton, QHBoxLayout, QVBoxLayout, QComboBox, QFormLayout, QSpacerItem, QSizePolicy,
                             QDialog, QTableWidget, QTableWidgetItem, QScrollArea, QDialogButtonBox, QMessageBox)

import stock_query
import traceback
import warnings
import logging
import sys


def error_window(err_info=None):
    dialog = QDialog()
    dialog.setWindowTitle("!Error!")

    layout = QVBoxLayout(dialog)

    error_message = QLabel(err_info)
    error_message.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
    error_message.setWordWrap(True)

    scroll_bar = QScrollArea()
    scroll_bar.setWidgetResizable(True)
    scroll_bar.setWidget(error_message)

    layout.addWidget(scroll_bar)

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok, dialog)
    buttons.accepted.connect(dialog.accept)
    layout.addWidget(buttons)

    dialog.setLayout(layout)
    dialog.exec()


def delisted_window(ticker_name=None):
    dialog = QDialog()
    dialog.setWindowTitle("!Delisted!")

    layout = QVBoxLayout(dialog)

    delisted_msg = QLabel(f"{ticker_name} Delisted")
    delisted_msg.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
    delisted_msg.setWordWrap(True)

    scroll_bar = QScrollArea()
    scroll_bar.setWidgetResizable(True)
    scroll_bar.setWidget(delisted_msg)

    layout.addWidget(scroll_bar)

    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok, dialog)
    buttons.accepted.connect(dialog.accept)
    layout.addWidget(buttons)

    dialog.setLayout(layout)
    dialog.exec()

def confirm_window(choice=None, ticker=None):
    if choice.lower() == "quit":
        confirm_win = QMessageBox.question(None, "Quit", "Would you like to Quit?",
                                              QMessageBox.StandardButton.Ok |
                                              QMessageBox.StandardButton.Cancel)
        if confirm_win == QMessageBox.StandardButton.Ok:
            return sys.exit()
    if choice.lower() == "execute":
        confirm_win = QMessageBox.question(None, "Execute", f"Would you like to query data for {ticker}?",
                                              QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
        if confirm_win == QMessageBox.StandardButton.Ok:
            return True

def click_exit():
    exit_window = confirm_window("quit")
    return exit_window

def click_confirm_execute(stock_ticker):
    execute_window = confirm_window(choice="execute", ticker=stock_ticker)
    return execute_window

class CompletionPopup(QDialog):
    def __init__(self, stock_data=None, meanchange=None, totalchange=None, totalpercent=None, ticker_name=None):
        super().__init__()

        stock_name = ticker_name
        self.setWindowTitle(stock_name)
        self.ok_button = QPushButton("OK", self)

        self.ok_button.clicked.connect(self.close)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Average Change: {meanchange}"))
        layout.addWidget(QLabel(f"Total Change: {totalchange}"))
        layout.addWidget(QLabel(f"Total Percent Change: {totalpercent:.2f}%"))

        table = QTableWidget(self)
        table.setRowCount(len(stock_data))
        table.setColumnCount(len(stock_data.columns))
        table.setHorizontalHeaderLabels(stock_data.columns)
        for idx, row in stock_data.iterrows():
            for col, ele in enumerate(stock_data.columns):
                element = QTableWidgetItem(str(stock_data.iat[idx, col]))
                element.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                table.setItem(idx, col, element)  # add element to table position idx, col
        table.setWordWrap(True)
        table.resizeRowsToContents()
        table.resizeColumnsToContents()
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        layout.addWidget(table)
        layout.addWidget(self.ok_button)
        self.setLayout(layout)


class StockWorker(QThread):
    finished = pyqtSignal(tuple)
    error_message = pyqtSignal(str)
    delisted_message = pyqtSignal(str)
    def __init__(self, stock_ticker=None, timeframe=None):
        super().__init__()
        self.ticker_name = stock_ticker
        self.timeframe = timeframe
        self.stockquery = stock_query.StockQuery(self.ticker_name, self.timeframe)
    def run(self):
        try:
            stock_data = self.stockquery.create_stock_dataframe()
            if stock_data.empty:
                return self.delisted_message.emit(self.ticker_name)
            mean_change = self.stockquery.get_mean_change(stock_data)
            total_change = self.stockquery.get_total_change(stock_data)
            total_percent_change = self.stockquery.get_percent_change(stock_data)

            return self.finished.emit((stock_data, mean_change, total_change, total_percent_change))
        except Exception as e:
            traceback_info = traceback.format_exc()
            self.error_message.emit(traceback_info)


class StockUI(QWidget):
    def __init__(self, main_window):
        super().__init__()
        # Constant dict for time range selection
        self.TIME_RANGE = {'1 Day':'1d', '5 Days':'5d',
                           '1 Month':'1mo', '3 Months':'3mo', '1 Year': '1y'}
        self.main_window = main_window
        self.setWindowTitle("Stock Ticker Query")

        # TIME LABEL AND COMBO BOX
        self.time_label = QLabel("Select time frame")
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.time_combobox = QComboBox(self)
        [self.time_combobox.addItem(ele) for ele in self.TIME_RANGE]
        self.time_combobox.setFixedSize(100, 35)
        self.time_combobox.activated.connect(self.get_current_timerange)

        # Company Label and Stock Ticker Text Box
        self.company_label = QLabel("Query Stock\n\nEnter company stock ticker\n(Max 4 characters)")
        self.company_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.user_text_box = QLineEdit()
        self.user_text_box.setFixedSize(60, 30)
        self.user_text_box.setMaxLength(4)

        # Execute and Quit buttons
        self.execute_button = QPushButton("Execute")
        self.execute_button.setFixedSize(60, 30)
        self.execute_button.clicked.connect(self.execute_stockquery)
        self.quit_button = QPushButton("Quit")
        self.quit_button.setFixedSize(60, 30)
        self.quit_button.clicked.connect(click_exit)

        # Establish form layout and call stock_ui function
        self.layout = QFormLayout()
        self.stock_ui()


    def stock_ui(self):
        # TEXT BOX
        text_box_layout = QHBoxLayout()
        text_box_layout.addStretch(1)
        text_box_layout.addWidget(self.user_text_box)
        text_box_layout.addStretch(1)

        # Add widgets to UI layout
        self.layout.addRow(self.company_label)
        self.layout.addRow(text_box_layout)
        self.layout.addRow(self.time_label)

        # Establish combo box layout and widget, and center ui
        combo_box_layout = QVBoxLayout()
        combo_box_inner_layout = QHBoxLayout()
        combo_box_inner_layout.addStretch(1)
        combo_box_inner_layout.addWidget(self.time_combobox)
        combo_box_inner_layout.addStretch(1)
        combo_box_layout.addLayout(combo_box_inner_layout)
        # Add combobox layout to UI
        self.layout.addRow(combo_box_layout)

        # Create vertical spacing
        self.layout.addItem(QSpacerItem(20, 20, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))

        # Create layout for quit and execute buttons, create 60px horizontal space
        hbox = QHBoxLayout()
        hbox.addWidget(self.quit_button)
        hbox.addSpacing(60)
        hbox.addWidget(self.execute_button)
        # Add button layout to stock ui layout
        self.layout.addRow(hbox)

        # set stock ui layout as current primary layout
        self.setLayout(self.layout)

    def get_current_timerange(self):
        current_select = self.time_combobox.currentText()
        current_timeframe = self.TIME_RANGE.get(current_select)
        return current_timeframe

    def get_current_ticker(self):
        current_ticker = self.user_text_box.text().upper()
        return current_ticker

    def execute_stockquery(self):
        try:
            time_range = self.get_current_timerange()
            ticker = self.get_current_ticker()
            if click_confirm_execute(ticker):
                self.worker = StockWorker(stock_ticker=ticker, timeframe=time_range)
                self.worker.error_message.connect(error_window)
                self.worker.delisted_message.connect(self.main_window.delisted_popup)
                self.worker.finished.connect(self.main_window.stock_completion_popup)
                self.worker.start()
        except Exception as e:
            traceback_info = traceback.format_exc()
            error_text = f"Error: {e}\n\n{traceback_info}"
            return error_window(error_text)



class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Stock Data Query")
        self.stockui = StockUI(self)
        self.ticker_name = self.stockui.get_current_ticker()

        hbox = QHBoxLayout()
        hbox.addWidget(self.stockui)

        self.setLayout(hbox)
        self.show()

    def stock_completion_popup(self, result):
        stock_data, mean_change, total_change, total_percent_change = result
        stock_ticker = self.stockui.get_current_ticker()
        self.popup = CompletionPopup(stock_data, mean_change, total_change, total_percent_change, stock_ticker)
        self.popup.show()

    def delisted_popup(self, tickername):
        delisted_window(tickername)



def main():
    warnings.filterwarnings("ignore", category=RuntimeWarning)
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

main()