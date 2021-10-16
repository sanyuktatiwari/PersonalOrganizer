'''
This Project has been created to organize and maintain daily official activities.
Creator : Sanyukta Tiwari
Last Modified date : 19/9/2021
'''

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QVBoxLayout, QWidget, QPushButton, QLabel, QGridLayout, QLineEdit
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from openpyxl import load_workbook
from datetime import date
from PyQt5.Qt import QWidget
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
from cProfile import label

def onclick():
    text = textEdit.toPlainText()
    wb = load_workbook("D:\Personal Organizer\MyActivities.xlsx")
    sheet = wb.active
    sheet.append((date.today(),text))
    print(text)
    textEdit.clear()
    wb.save("D:\Personal Organizer\MyActivities.xlsx")
    wb.close()

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Personal Organizer")
        layout = QGridLayout()
        self.setLayout(layout)
        label1 = QLabel("Widget in Tab 1.")
        label2 = QLabel("Widget in Tab 2.")
        tabwidget = QTabWidget()
        tabwidget.addTab(label1, "Tab 1")
        tabwidget.addTab(label2, "Tab 2")
        layout.addWidget(tabwidget, 0, 0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()