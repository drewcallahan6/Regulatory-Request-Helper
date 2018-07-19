import sys
import os
from PyQt4 import QtGui, QtCore
from PyQt4.QtCore import pyqtSlot
from PyQt4.QtCore import *
from PyQt4.QtGui import * 
from PyQt4.QtGui import *

#Excel Reading and Writing Modules
import xlrd
import xlwt
import datetime
from xlutils.copy import copy

class Window(QtGui.QMainWindow):

    def __init__(self):
        super(Window,self).__init__()
        self.setGeometry(50,50,1100,1500)
        self.setWindowTitle("Regulatory Helper")
        self.setWindowIcon(QtGui.QIcon('griffith-logo.png'))

        palette = QtGui.QPalette()
        myPixmap = QtGui.QPixmap('green-wave.jpg')
        myScaledPixmap = myPixmap.scaled(self.size(), QtCore.Qt.KeepAspectRatio, transformMode = QtCore.Qt.SmoothTransformation)
        palette.setBrush(QtGui.QPalette.Window, QtGui.QBrush(myScaledPixmap))
        self.setPalette(palette)

        #font size 
        newfont = QtGui.QFont("Times",16)

        self.uploadButton = QtGui.QPushButton('Convert to One', self)
        self.uploadButton.setFont(newfont)
        self.uploadButton.resize(250,32)
        # hBoxLayout = QtGui.QHBoxLayout()
        # hBoxLayout.addWidget(self.uploadButton)
        # self.setLayout(hBoxLayout)
        self.connect(self.uploadButton, QtCore.SIGNAL('clicked()'), self.converter)
        self.uploadButton.move(600,100)


        self.home()

    def home(self):
 
        pic = QtGui.QLabel(self)
        pic.setGeometry(10, 10, 500, 500)
        pic.setPixmap(QtGui.QPixmap(os.getcwd() + "/bigger-griffith-logo.png"))
        pic.move(600,275)
        
        newfont = QtGui.QFont("Times",16)
        self.Label1 = QLabel(self)
        self.Label1.setText('Please Copy the paths for all Requests Below:')
        self.Label1.setFont(newfont)
        self.Label1.resize(525,32)
        self.Label1.move(50, 50)

        self.line = QPlainTextEdit(self)
        self.line.resize(300, 375)
        self.line.move(150, 90)

        newfont2 = QtGui.QFont("Times", 12)
        self.StatusLabel = QLabel(self)
        self.StatusLabel.setText("Current Status: No Files Set to be Converted")
        self.StatusLabel.setFont(newfont2)
        self.StatusLabel.resize(500,32)
        self.StatusLabel.move(550,200)

        self.progress = QtGui.QProgressBar(self)
        self.progress.setGeometry(200,80,300,20)
        self.progress.move(600,250)

        self.show()
    
    def converter(self):
        ## take names of all the files and add them to a list
        self.StatusLabel.setText("Current Status: Taking File Names")
        paths = self.line.toPlainText()
        file_names = []
        num_quotations = 0
        indiv_path = ''
        for i in paths:
            if i == '"':
                num_quotations += 1
            if num_quotations == 2:
                if i != '"':
                    indiv_path += i
                ###to get rid of '\n'
                if indiv_path[0:1] == "\n":
                    indiv_path = indiv_path[1:]
                file_names.append(indiv_path)
                indiv_path = ''
                num_quotations = 0
            else:
                if i != '"':
                    indiv_path += i
        self.progress.setValue(10)

        ## take out appropriate cells from each file and add them to new spreadsheet

        ###setting up stuff for progress bar
        file_len_pro_val = 90/len(file_names)
        current_pro_val = 10
        self.StatusLabel.setText("Current Status: Reading each File")
        to_copy = xlwt.Workbook()
        sheet1 = to_copy.add_sheet("Copy These Values into Database", cell_overwrite_ok =True)
        # for n in range(1,1000):
        #     print(n)
        #     sheet1.write(n,52, "|||")
        statement = "COPY UP UNTIL HERE(NOT INCLUDING THIS COLUMN)"
        sheet1.write(0,52,statement)

        index = 0
        for i in file_names:
            index += 1
            self.StatusLabel.setText("Current Status: Copying Data from File " + str(index))
            condition = "not break"
            book = xlrd.open_workbook(i)
            sheet_form = book.sheet_by_index(0)

            for row in range(16,80):
                if condition == "break":
                    break
                if index == 1:
                    row1 = row-15
                else:
                    row1 += 1
                for col in range(27,80):
                    x = sheet_form.cell(row,col)
                    x = x.value
                    ###testing if it is a date and fixing it
                    if col == 39 or col == 40 and x != '':
                        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(x, book.datemode)
                        year = str(year)
                        year = year[-2:]
                        year = int(year)
                        x = str(month) + '/' + str(day) + '/' + str(year)

                    if col == 27:
                        if x == '':
                            condition = "break"
                            row1 -= 1
                            break

                    if x == 0.0 or x == "0/0/0":
                        x = ''
                    col1 = col-27

                    if type(x) != xlrd.sheet.Cell:
                        sheet1.write(row1,col1,x)

            current_pro_val += file_len_pro_val
            self.progress.setValue(current_pro_val)

        ##saves to same folder as application
        to_copy.save("to_copy.xls")
        self.StatusLabel.setText("Current Status: All Files Copied and Ready in to_copy.xls")

        ##adding to database-example1
        # self.add_to_database()

    #optional addition of saving automatically to database
    # def add_to_database(self):
    #     self.progress.setValue(80)
    #     self.StatusLabel.setText("Current Status: Moving to Master Database (~10 sec)")
    #     ##where you would put network drive path
    #     path_to_database = "C:/Users/andrewbc/Documents/Griffith Stuff/Regulatory/database-example.xls"
    #     database = xlrd.open_workbook(path_to_database, formatting_info=True)
    #     master1 = database.sheet_by_index(2)
    #     database_copy = copy(database)
    #     master = database_copy.get_sheet(2)

    #     to_copy = xlrd.open_workbook("to_copy.xls")
    #     sheet1 = to_copy.sheet_by_index(0)

    #     num_rows = sheet1.nrows
    #     num_cols = sheet1.ncols 
    #     ##find opening in master database
    #     for i in range(7,10000):
    #         cell = master1.cell(i,5)
    #         cell = cell.value
    #         if cell == '':
    #             index_opening = i
    #             break
    #     index_opening -= 1

    #     ##writing in new_data
    #     ###font
    #     font = xlwt.Font()
    #     font.name = 'Arial'
    #     font.height = 160
    #     style = xlwt.XFStyle() 
    #     style.font = font 
    #     for row in range (1,(num_rows)):
    #         for col in range(0,(num_cols)):
    #             cell_data_copy = sheet1.cell(row,col)
    #             cell_data_copy = cell_data_copy.value
    #             master.write((row+index_opening), (col+3), cell_data_copy, xlwt.easyxf("font: name Arial, height 160; align: horiz center"))

    #     database_copy.save("C:/Users/andrewbc/Documents/Griffith Stuff/Regulatory/database-example1.xls")
    #     self.progress.setValue(100)
    #     self.StatusLabel.setText("Current Status: All Information Copied to Master Database")

#creating instance of the application        
def run():
    app = QtGui.QApplication(sys.argv)
    GUI = Window()
    sys.exit(app.exec_())

#automatically running the program
run()
