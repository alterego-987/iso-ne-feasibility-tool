from cProfile import label
import ctypes
import sys
from turtle import goto
from PyQt5 import uic, QtWidgets
from ISO_NE_Flow_Tool_v7_0 import *

class appScreen(QtWidgets.QMainWindow):
    def __init__(self):
        super(appScreen,self).__init__()
        uic.loadUi('feasibilitytool.ui',self)
        self.show()
        self.run.clicked.connect(self.runFunction)
        self.load.clicked.connect(self.loadFunction)
        self.clear.clicked.connect(self.clearlist)
        self.discharge = self.findChild(QtWidgets.QRadioButton, 'dischargingButton')
        self.charge = self.findChild(QtWidgets.QRadioButton, 'chargingButton')
        
    def runFunction(self):
        busNumber = int(self.busNoInput.text())
        projSize = int(self.projSizeInput.text())
        if self.discharge.isChecked():
            charging = 'N'
        if self.charge.isChecked():
            charging = 'Y'

        file = self.listFiles.item(0).text()
        inputParameters = (busNumber,projSize,charging)
        writeExcel(file,inputParameters)
        self.outMessage.setText('redispatch_' + file + ' created!')

    def loadFunction(self):
        ls = QtWidgets.QFileDialog.getOpenFileName(
            parent=self,
            caption='Select the file',
        )
        
        filename = ls[0].split('/')
        print(filename[-1])
        self.listFiles.addItem(filename[-1])
    
    def clearlist(self):
        self.listFiles.clear()

# main
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = appScreen()
    sys.exit(app.exec())