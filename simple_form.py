# Import widgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QLabel, QRadioButton, QTextEdit, QPushButton, QLineEdit, QMessageBox, QFileDialog, QTableWidget
from PyQt5 import uic
import sys
from PyQt5 import QtCore
import pandas as pd
from tabulate import tabulate

filename = "No file chosen"
formOutput = []
masterList = {'name': [],
            'address': [],
            'contact': [],
            'start_time': [],
            'date': [],
            'sex': [],
            'comments': []
            } # A dictionary of lists
class UI(QMainWindow):
    def __init__(self):
        super(UI, self).__init__()
        # UI to be loaded
        uic.loadUi("simple_form.ui", self)
        # Define widgets here
        self.submit = self.findChild(QPushButton, "pushButton")
        self.name = self.findChild(QLineEdit, "lineEdit")
        self.address = self.findChild(QLineEdit, "lineEdit_2")
        self.contact = self.findChild(QLineEdit, "lineEdit_3")
        self.startTime = self.findChild(QLineEdit, "lineEdit_4")
        self.date = self.findChild(QLineEdit, "lineEdit_5")
        self.sexMale = self.findChild(QRadioButton, "radioButton")
        self.sexFemale = self.findChild(QRadioButton, "radioButton_2")
        self.comments = self.findChild(QTextEdit, "textEdit")
        self.outputExcelButton = self.findChild(QPushButton, "pushButton_2")
        self.filenameOutput = self.findChild(QLabel, "label_9")
        self.importExcelButton = self.findChild(QPushButton, "pushButton_3")
        self.chooseFile = self.findChild(QPushButton, "pushButton_4")
        self.tableOutput = self.findChild(QLabel, "label_10")
        self.tableRefresh = self.findChild(QPushButton, "pushButton_5")

        # Define events here
        self.submit.clicked.connect(self.submitForm)
        self.outputExcelButton.clicked.connect(self.outputExcel)
        self.chooseFile.clicked.connect(self.chooseFileName)
        self.importExcelButton.clicked.connect(self.importExcel)
        self.tableRefresh.clicked.connect(self.showTable)
        # Show the app
        self.show()
    
    def errorDialog(self, text):
        msg = QMessageBox()
        msg.setWindowTitle("Error")
        msg.setIcon(QMessageBox.Critical)
        msg.setText(text)
        msg.exec_()
    
    def infoDialog(self, text):
        msg = QMessageBox()
        msg.setWindowTitle("Error")
        msg.setIcon(QMessageBox.Information)
        msg.setText(text)
        msg.exec_()
    
    def confirmAction(self, text):
        reply = QMessageBox.question(self, 'Confirmation Prompt', text,QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            return True
        else:
            return False


    def clearInputs(self):
        self.name.setText("")
        self.address.setText("")
        self.contact.setText("")
        self.startTime.setText("")
        self.date.setText("")
        self.sexMale.setChecked(False)
        self.sexFemale.setChecked(False)
        self.comments.setText("")
    
    def submitForm(self):
        sex = ""
        error = False
        if self.sexMale.isChecked():
            sex = "Male"
        if self.sexFemale.isChecked():
            sex = "Female"
        formKeys = ["name","address","contact","start_time","date","sex","comments"]
        formValues = [self.name.text(), self.address.text(), self.contact.text(), self.startTime.text(), self.date.text(), sex, self.comments.toPlainText()]
        for x in formValues:
            if x == '':
                self.errorDialog("Fill all data required.")
                error = True
                break
        if not error:
            conf = self.confirmAction("Are you sure you want to submit the form?")
            if conf:
                for i in range(0,len(formValues)):
                    masterList[formKeys[i]].append(formValues[i])
                formOutput.clear() # Remember to clear after!
                self.clearInputs()
                print(masterList)
            else:
                print("Aborted.")

    def outputExcel(self):
        df = pd.DataFrame(masterList)
        df2 = pd.DataFrame(masterList)
        with pd.ExcelWriter('output.xlsx') as writer:
            df.to_excel(writer, sheet_name='What da sheet 1', index=False)
            df2.to_excel(writer, sheet_name="second wat da sheet", index=False)
        self.infoDialog("Done.")

    def chooseFileName(self):
        global filename
        filename, _ = QFileDialog.getOpenFileName(self, "Open File", "","Excel Files (*.xlsx)") #;;CSV Files (*.csv) if we wanna add csv to files accepted
        if filename:
            self.filenameOutput.setText(filename)
        else:
            self.filenameOutput.setText("No file chosen")

    def importExcel(self):
        print("filename is: ",filename)
        try:
            df = pd.read_excel(filename)
            masterlistDict = df.to_dict('list')
            print(masterlistDict)
            global masterList
            masterList = masterlistDict.copy()
        except:
            print("Error - File not found.")
    
    def showTable(self):
        df = pd.DataFrame(masterList)
        x = tabulate(df, headers = 'keys', tablefmt = 'psql')
        print(x)
        self.tableOutput.setText(x)

    
# Initialization
if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
app = QApplication(sys.argv)
UIWindow = UI()
app.exec_()