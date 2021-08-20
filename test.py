from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QAction, QTableWidget,QTableWidgetItem,QVBoxLayout
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
import sys

data = {'col1':['1','2','3'],
		'col2':['1','alam','1','3'],
		'col3':['1']}
 
class TableView(QTableWidget):
	def __init__(self, data, *args):
		QTableWidget.__init__(self, *args)
		self.data = data
		self.setData()
		self.resizeColumnsToContents()
		self.resizeRowsToContents()
 
	def setData(self): 
		horHeaders = []
		for n, key in enumerate(sorted(self.data)):
			horHeaders.append(key)
			print(n,key)
			for m, item in enumerate(self.data[key]):
				print(m,item)
				newitem = QTableWidgetItem(item)
				self.setItem(m, n, newitem)
		self.setHorizontalHeaderLabels(horHeaders)
 
def main(args):
	app = QApplication(args)
	table = TableView(data, 10, 10)
	table.show()
	sys.exit(app.exec_())
 
if __name__=="__main__":
	main(sys.argv)