from datetime import datetime
import sys
import json
import design
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import QtCore
import openpyxl
from openpyxl.styles import *

class NumericDelegate(QItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        regex = QRegExp('[0-9]+')
        validator = QRegExpValidator(regex, editor)
        editor.setValidator(validator)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        editor.setText(str(value))

    def setModelData(self, editor, model, index):
        value = editor.text()
        model.setData(index, int(value), Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

class MainWindow(QMainWindow):
    def __init__(self, new_window):
        super().__init__()

        # Load the data from the JSON file
        with open('docs/drinks.json') as f:
            data = json.load(f)

        # Create the model for the table
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(['Name', 'Popisano Stanje'])

        for section in data:
            for item in data[section]:
                row = [QStandardItem(item['naziv']), QStandardItem()]
                row[1].setData(item['prodaja'], Qt.DisplayRole)
                self.model.appendRow(row)

        # Create a proxy model for filtering the data
        self.proxy_model = QSortFilterProxyModel()
        self.proxy_model.setSourceModel(self.model)
        self.proxy_model.setFilterKeyColumn(0)

        # Create the search field
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText('Search...')
        self.search_field.textChanged.connect(self.filter_data)
        self.search_field.setStyleSheet(design.field_fancy())

        # Create the table view
        self.table_view = QTableView()
        self.table_view.setModel(self.proxy_model)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        self.table_view.setStyleSheet(design.table_fancy())

        # Set the delegate for the second column of the table
        delegate = NumericDelegate()
        self.table_view.setItemDelegateForColumn(1, delegate)

        # Create the buttons
        self.save_button = QPushButton('Save')
        self.save_button.clicked.connect(self.save_data)
        self.save_button.setStyleSheet(design.make_fancy())

        self.submit_button = QPushButton('Submit')
        self.submit_button.clicked.connect(self.submit_data)
        self.submit_button.setStyleSheet(design.make_fancy())

        # Create the layout
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.submit_button)
        button_layout.setAlignment(Qt.AlignBottom)

        self.back_button = QPushButton("Korak Nazad", self)
        self.back_button.move(150, 300)
        # Connect the back button to the function that closes the current window and opens the main window
        self.back_button.clicked.connect(lambda: self.back_to_main(new_window))
        self.back_button.setStyleSheet(design.make_fancy_back_button())

        #layout.addStretch(1)

        layout = QVBoxLayout()
        layout.addWidget(self.search_field)
        layout.addWidget(self.table_view, 1)
        layout.addLayout(button_layout)
        layout.addWidget(self.back_button)
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        
        self.central_widget = QWidget()
        self.central_widget.setLayout(layout)
        # Set the central widget and window properties
        self.setCentralWidget(self.central_widget)
        self.setWindowTitle("Caffe Bar Casper")
        self.setWindowIcon(QIcon("pics/logo.png"))
        self.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)
        available_geometry = QGuiApplication.primaryScreen().availableGeometry()
        self.setGeometry(available_geometry)
        self.showMaximized()

        # Set the background image
        palette = self.palette()
        background_image = QPixmap("pics/background.png")
        palette.setBrush(QPalette.Window, QBrush(background_image))
        self.setPalette(palette)
        new_window.hide()

    def filter_data(self, text):
        self.proxy_model.setFilterRegExp(text)
    
    def back_to_main(self, current_window): #povratak sa drugog window-a na main
        # Close the current window and show the main window
        self.close()
        current_window.show()

    def save_data(self, info = 0):
        # Save the data to a file
        for row in range(self.model.rowCount()):
            name = self.model.item(row, 0).text()
            popis = self.model.item(row, 1).text()
            with open("docs/drinks.json", "r") as f:
                data = json.load(f)
            
            for section in data:
                for item in data[section]:
                    if item["naziv"].lower() == name.lower():
                        item["prodaja"] = int(popis)

            jsonString = json.dumps(data)
            
            with open("docs/drinks.json", "w") as f:
                f.write(jsonString)

        # Show a message box to confirm the save
        if(info ==0):
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Obavestenje")
            msg_box.setText('Popisano stanje je uspesno sacuvano!')
            msg_box.exec_()

    def restart_popis(self):
        with open("docs/drinks.json", "r") as f:
            data = json.load(f)
            
        for section in data:
            for item in data[section]:
                item["prodaja"] = 0

        jsonString = json.dumps(data)
        
        with open("docs/drinks.json", "w") as f:
            f.write(jsonString)

    def submit_data(self):
        self.save_data(1)
        self.restart_popis()
        # Get the data from the table
        data = []
        for row in range(self.model.rowCount()):
            name = self.model.item(row, 0).text()
            num_items = self.model.item(row, 1).text()
            data.append({"naziv":name, "prodaja":num_items})

        # create a new workbook
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
        
        grey_fill = PatternFill(start_color='d3d3d3', end_color='d3d3d3', fill_type='solid')

        worksheet['A1'] = 'Naziv Artikla'
        worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet['A1'].font = Font(bold=True)
        worksheet['A1'].border = thin_border
        worksheet['A1'].fill = grey_fill
        worksheet['B1'] = 'Popisano Stanje Artikla'
        worksheet['B1'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet['B1'].font = Font(bold=True)
        worksheet['B1'].border = thin_border
        worksheet['B1'].fill = grey_fill

        
        worksheet.column_dimensions['A'].width = 35
        worksheet.column_dimensions['B'].width = 20

        count = 1
        for i in range(len(data)):
            count = count + 1
            worksheet['A' + str(count)] = data[i]["naziv"].capitalize()
            worksheet['A'+ str(count)].alignment = Alignment(horizontal='center', vertical='center')
            worksheet['A'+ str(count)].font = Font(bold=True)
            worksheet['A'+ str(count)].border = thin_border
            worksheet['B' + str(count)] = data[i]["prodaja"]
            worksheet['B'+ str(count)].alignment = Alignment(horizontal='center', vertical='center')
            worksheet['B'+ str(count)].font = Font(bold=True)
            worksheet['B'+ str(count)].border = thin_border
        
        workbook.save('popis.xlsx')

        # set up the SMTP server
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        smtp_username = 'trbovicdusica@gmail.com'
        smtp_password = 'yzerjqneezzkmdey'
        sender_email = smtp_username
        receiver_email = 'trbovicdusica@gmail.com'

        # create the message object
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = 'Popisano stanje artikala - ' + str(datetime.now().day) + "." + str(datetime.now().month) + "." + str(datetime.now().year)

        # open the file in bynary
        binary_file = open('popis.xlsx', 'rb')

        # create the attachment object
        payload = MIMEBase('application', 'octate-stream', Name='popis.xlsx')
        payload.set_payload((binary_file).read())

        # encode the attachment in base64
        encoders.encode_base64(payload)

        # add header with pdf name
        payload.add_header('Content-Decomposition', 'attachment', filename='popis.xlsx')
        msg.attach(payload)

        # create SMTP session
        session = smtplib.SMTP(smtp_server, smtp_port)

        # start TLS for security
        session.starttls()

        # login to the email server
        session.login(smtp_username, smtp_password)

        # send the email
        text = msg.as_string()
        session.sendmail(sender_email, receiver_email, text)

        # end the session
        session.quit()

        msg_box = QMessageBox()
        msg_box.setWindowTitle("Obavestenje")
        msg_box.setText('Popisano stanje je Uspesno poslato na email adresu: casper.trbovic@gmail.com.')
        msg_box.exec_()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())