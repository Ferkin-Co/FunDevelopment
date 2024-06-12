import stock_query
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QWidget, QFormLayout, QApplication, QLineEdit, QLabel, QPushButton, QHBoxLayout, QVBoxLayout


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Stock Ticker Query")
        self.layout = QVBoxLayout()
        self.company_label = QLabel("Query Stock\n\nEnter company stock ticker\n(Max 4 characters)")
        self.company_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.execute_button = QPushButton("Execute")
        self.quit_button = QPushButton("Quit")

        # USER TEXT BOX
        self.user_text_box = QLineEdit()
        self.user_text_box.setFixedWidth(60)
        self.user_text_box.setMaxLength(4)

        text_box_layout = QHBoxLayout()
        text_box_layout.addStretch(1)
        text_box_layout.addWidget(self.user_text_box)
        text_box_layout.addStretch(1)

        self.layout.addWidget(self.company_label, alignment=Qt.AlignmentFlag.AlignCenter)
        self.layout.addLayout(text_box_layout)
        self.layout.addSpacing(20)


        hbox = QHBoxLayout()
        hbox.addWidget(self.quit_button)
        hbox.addSpacing(60)
        hbox.addWidget(self.execute_button)

        self.layout.addLayout(hbox)


        self.setLayout(self.layout)
        self.show()

app = QApplication([])
window = MainWindow()
window.show()
app.exec()