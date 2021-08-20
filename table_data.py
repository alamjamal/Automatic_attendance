import sys
import os
import pandas as pd
import io
# import numpy as np

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDesktopWidget, QTableWidgetItem,QDialog,QTableWidget, QApplication,QFileDialog,QMainWindow,QLabel,QLineEdit,QWidget,QMessageBox
from PyQt5.uic import  loadUiType
from PyQt5.QtGui import QIcon
from os import path
def resource_path(relative_path):
	base_path=getattr(sys, '_MEIPASS',os.path.dirname(os.path.abspath(__file__)))
	return os.path.join(base_path, relative_path)

ui,_=loadUiType(resource_path('demo.ui'))

class MainApp(QMainWindow, ui):
	def __init__(self):
		QMainWindow.__init__(self)
		self.setupUi(self)
		self.path=f'C:\\Users\\{os.getlogin()}\\Desktop'
		self.pushButton.clicked.connect(self.File_XLSX)
		self.pushButton_2.clicked.connect(self.File_TXT)
		self.pushButton_3.clicked.connect(self.CompareAndGenerate)
		self.pushButton_4.clicked.connect(self.closeFun)
		self.pushButton_5.clicked.connect(self.minimizeWidow)
		self.Combotext=""
		self.comboBox.activated[str].connect(self.onChangedCombo)
		self.df1=""
		self.df2=""
		self.words=""
		self.count=0
		self.excel_path=""
		self.pushButton.setEnabled(False)
		self.pushButton_2.setEnabled(False)
		self.pushButton_3.setEnabled(False)


	def onChangedCombo(self,text):
		if text=="Select Month" or text== "":
			self.ShowMessageBox('Select Month', 'Please Select Month Correctly')
			return
		self.Combotext=text
		self.pushButton.setEnabled(True)

	def File_XLSX(self):
		try:
			if self.Combotext=="Select Month":
				return
			fname = QFileDialog.getOpenFileName(self, "Select File",f'{self.path}',"*.xlsx")
			if fname == ('', ''):
				return
			self.path=fname[0]
			self.excel_path=fname[0]
			# print(self.excel_path)
			with pd.ExcelFile(self.excel_path) as reader:

				self.df2= pd.read_excel(reader, sheet_name=self.Combotext, skiprows= 1, usecols='B:AI')

			self.df2 = pd.DataFrame(self.df2)
			self.df2=self.df2.fillna(' ')
			self.table_2.setColumnCount(len(self.df2.columns))
			self.table_2.setRowCount(len(self.df2.index))
			self.table_2.setHorizontalHeaderLabels(self.df2.columns.astype(str))
			for i in range(len(self.df2.index)):
				for j in range(len(self.df2.columns)):
					self.table_2.setItem(i,j,QTableWidgetItem(str(self.df2.iloc[i, j])))
			
			self.count+=1
			val=int(self.count/3*100)
			self.progressBar.setValue(val)

			self.pushButton.setEnabled(False)
			self.pushButton_2.setEnabled(True)
			self.pushButton_3.setEnabled(False)

			return self.df2

		# except (AttributeError,ValueError,IOError,AssertionError) as e:
		# 	print(e)
		except PermissionError as e:
			self.ShowMessageBox('Something Went Wrong', 'Please Close Excel File')
		except ValueError as e:
			self.ShowMessageBox('Something Went Wrong', 'Please Enter Valid Month Name')


	def File_TXT(self):
		try:
			fname = QFileDialog.getOpenFileNames(self, "Select File",f'{self.path}',"*.txt")
			if fname == ('', ''):
				return
			self.path=fname[0]
			path=fname[0]
			myDict = dict()
			import re
			count=0
			for i in range(len(path)):
				with open(path[i], 'r') as file:
					text = file.read()
					self.words = sorted(set(re.findall(r'[0-9]{2}[A-Za-z]{4}[0-9]{3}[HYhy]{2}',text)))
					self.words= [x.upper() for x in self.words]
					if self.words:
						name = path[i].split("/")[-1]
						fileName=re.sub("[^0-9]", "", name)
						count+=1
						myDict.update({int(fileName) : self.words})

			myDict=dict(sorted(myDict.items()))
			self.df1 = pd.DataFrame.from_dict(myDict, orient='index')
			self.df1 = self.df1.transpose()
			self.df1=self.df1.fillna(' ')

			self.table_2.setColumnCount(len(self.df1.columns))
			self.table_2.setRowCount(len(self.df1.index))
			self.table_2.setHorizontalHeaderLabels(self.df1.columns.astype(str))
			for i in range(len(self.df1.index)):
				for j in range(len(self.df1.columns)):
					self.table_2.setItem(i,j,QTableWidgetItem(str(self.df1.iloc[i, j])))
			
			self.count+=1  
			val=int(self.count/3*100)
			self.progressBar.setValue(val)

			self.pushButton.setEnabled(False)
			self.pushButton_2.setEnabled(False)
			self.pushButton_3.setEnabled(True)

			return self.df1

		except ZeroDivisionError as e:
			print(e)

	def CompareAndGenerate(self):
		try:
			for c in self.df1:
				self.df2[c] = self.df2.merge(self.df1[c], left_on="Roll Number", right_on=c, how="left")[f"{c}_y"]
				if self.df2[c].any():
					self.df2[c].fillna('A', inplace=True)
					self.df2[c] = self.df2[c].where(self.df2[c] == 'A', 'P')
			self.df2=self.df2.fillna(' ')
			self.df2['Total']=self.df2.eq('P').sum(1)

			# self.df2.to_excel('excel_file.xlsx', index=False)
			from openpyxl import load_workbook
			with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='a') as writer:
				writer.book = load_workbook(self.excel_path)
				writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
				# print(writer.sheets)
				self.df2.to_excel(writer, sheet_name=self.Combotext , index=False, startrow=1, startcol=1)

			html = self.df2.to_html()
			with io.open("index2.html", "w", encoding="utf-8") as f:
				f.write(html)

			self.ShowMessageBox('Generated Successfully', 'Congratulations !!!! \nYor File is Generated Successfully....... If Found Any Bug Please Report To \nalam@techanda.tech')

			self.table_2.setColumnCount(len(self.df2.columns))
			self.table_2.setRowCount(len(self.df2.index))
			self.table_2.setHorizontalHeaderLabels(self.df2.columns.astype(str))
			for i in range(len(self.df2.index)):
				for j in range(len(self.df2.columns)):
					self.table_2.setItem(i,j,QTableWidgetItem(str(self.df2.iloc[i, j])))

			self.count+=1
			val=int(self.count/3*100)
			self.progressBar.setValue(val)
			self.count=0

			self.pushButton.setEnabled(True)
			self.pushButton_2.setEnabled(False)
			self.pushButton_3.setEnabled(False)

			self.table_2.setRowCount(0)
			self.table_2.setColumnCount(0)

		# except (AttributeError,ValueError,IOError,AssertionError) as e:
		# 	print(e)

		except PermissionError as e:
			self.ShowMessageBox('Something Went Wrong', 'Please Close Excel File')


	def enableLightMode(self):
		style=open(resource_path('aqua.css'),'r')
		style=style.read()
		self.frame_DashCentral.setStyleSheet(style)


	def ShowMessageBox(self,title, messege):
		msgBox=QMessageBox()
		msgBox.setIcon(QMessageBox.Information)
		msgBox.setWindowTitle(title)
		msgBox.setText(messege)
		msgBox.setStandardButtons(QMessageBox.Ok)
		msgBox.exec_()

	def closeFun(self):
		msgBox = QMessageBox()
		self.setWindowIcon(QIcon(resource_path('icon.ico')))
		msgBox.setText("Close?? Are You Sure?? ")
		msgBox.setIcon(QMessageBox.Warning)
		msgBox.setWindowTitle("Warning!!")
		msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
		# msgBox.buttonClicked.connect(msgButtonClick)
		returnValue = msgBox.exec()
		if returnValue == QMessageBox.Ok:
			self.close()

	def minimizeWidow(self):
		self.showMinimized()
# background-color: rgb(40, 44, 52);
def main():
	app=QApplication(sys.argv)
	form=MainApp()
	# form.setWindowFlags(Qt.FramelessWindowHint)
	# qtRectangle = form.frameGeometry()
	# centerPoint = QDesktopWidget().availableGeometry().center()
	# qtRectangle.moveCenter(centerPoint)
	# form.move(qtRectangle.topLeft())
	form.show()
	app.exec_()


if __name__ == '__main__':
	main()