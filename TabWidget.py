from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
from PyQt5.Qt import QTabWidget, QWidget, QFormLayout, QLineEdit, QApplication,\
    QPushButton, QTextEdit, QComboBox
from functools import partial
from openpyxl import Workbook
import glob
import openpyxl
from datetime import date
    
class tabFrame(QTabWidget):
    def __init__(self, parent = None):
        super(tabFrame,self).__init__(parent)
        self.createTab = QWidget()
        self.insertTab = QWidget()
        self.editTab = QWidget()
        self.deleteTab = QWidget()
        self.viewTab = QWidget()
        
        self.addTab(self.createTab, "CREATE")
        self.addTab(self.insertTab, "INSERT")
        self.addTab(self.editTab, "EDIT")
        self.addTab(self.deleteTab,"DELETE")
        self.addTab(self.viewTab, "VIEW")
        
        self.createTabUI()
        self.insertTabUI()
        #self.editTabUI()
        #self.deleteTabUI()
        #self.viewTabUI()
        
        self.setWindowTitle("PERSONAL ORGANIZER")
        self.resize(500,300)
        
    def createTabUI(self):
        layout = QFormLayout()
        #Text Edit to enter the name of new Worksheet to create
        self.newWorkbookName = QTextEdit()
        layout.addRow("Name your Workbook", self.newWorkbookName)
        #Push button
        pushButton = QPushButton()
        pushButton.setText("ADD Workbook")
        layout.addRow("", pushButton)
        pushButton.clicked.connect(self.onclick)
        #Text Edit to show list of available Worksheets
        self.listWorkbooks = QTextEdit()
        self.listWorkbooks.setReadOnly(True)
        self.listWorkbooks.setPlainText(self.availableWorkbooks())
        layout.addRow("List of available Work sheets", self.listWorkbooks)
        
        self.createTab.setLayout(layout)
        
    def onclick(self):
        wb = Workbook()
        wb.save(filename = self.newWorkbookName.toPlainText()+".xlsx")
        self.listWorkbooks.setPlainText(self.availableWorkbooks())
        self.newWorkbookName.clear()
        
    def availableWorkbooks(self):
        workbooks = []
        for file in glob.glob("*.xlsx"):
            workbooks.append(file)
        print(''.join(str(workbook) for workbook in workbooks))
        return ' \n'.join(str(workbook) for workbook in workbooks)
    
    def insertTabUI(self):
        layout = QFormLayout()
        self.dropDownWBList = QComboBox()
        workbooks = []
        for file in glob.glob("*.xlsx"):
            workbooks.append(file)
        for workbook in workbooks:
            self.dropDownWBList.addItem(workbook)
        layout.addRow("Workbooks available", self.dropDownWBList)
        
        #Insert Text Edit to enter the row
        self.textToEnter = QTextEdit()
        layout.addRow("Enter Text",self.textToEnter)
        
        #Push button
        pushButton = QPushButton()
        pushButton.setText("INSERT")
        layout.addRow("", pushButton)
        pushButton.clicked.connect(self.onInsertTextclick)
        layout.addRow("Enter Text to insert", pushButton)
        
        self.insertTab.setLayout(layout)
        
    def onInsertTextclick(self):
        text = self.textToEnter.toPlainText()
        wb = openpyxl.load_workbook(str(self.dropDownWBList.currentText()))
        sheet = wb.active
        sheet.append((date.today(),text))
        print(text)
        self.textToEnter.clear()
        wb.save(str(self.dropDownWBList.currentText()))
        wb.close()

def main():
    app = QApplication(sys.argv)
    program = tabFrame()
    program.show()
    sys.exit(app.exec_())
    
if __name__ == '__main__':
    main()