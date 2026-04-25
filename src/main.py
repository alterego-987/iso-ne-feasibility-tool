import sys
from PyQt5 import uic, QtWidgets
from src.excel_writer import writeExcel

class AppScreen(QtWidgets.QMainWindow):
    def __init__(self):
        super(AppScreen, self).__init__()
        # Load the UI file from the correct relative path
        uic.loadUi('ui/feasibilitytool.ui', self)
        self.show()
        
        # Connect buttons
        self.run.clicked.connect(self.runFunction)
        self.load.clicked.connect(self.loadFunction)
        self.clear.clicked.connect(self.clearlist)
        
        self.discharge = self.findChild(QtWidgets.QRadioButton, 'dischargingButton')
        self.charge = self.findChild(QtWidgets.QRadioButton, 'chargingButton')
        
        # Keep track of absolute file paths loaded
        self.loaded_files = []
        
    def runFunction(self):
        try:
            busNumber = int(self.busNoInput.text())
            projSize = int(self.projSizeInput.text())
        except ValueError:
            self.showError("Please enter valid integers for Bus Number and Project Size.")
            return

        if self.discharge.isChecked():
            charging = 'N'
        elif self.charge.isChecked():
            charging = 'Y'
        else:
            self.showError('Please select Charging or Discharging.')
            return

        if not self.loaded_files:
            self.showError('Please load an Excel file first.')
            return

        # Execute for the first loaded file
        file_path = self.loaded_files[0]
        inputParameters = (busNumber, projSize, charging)
        
        try:
            writeExcel(file_path, inputParameters)
            # Extracted filename for display
            filename = file_path.split('/')[-1]
            self.outMessage.setText(f'redispatch_{filename} created!')
        except Exception as e:
            self.showError(f"Error during execution: {str(e)}")

    def loadFunction(self):
        file_dialog = QtWidgets.QFileDialog.getOpenFileName(
            parent=self,
            caption='Select the Excel file',
            filter="Excel Files (*.xlsx *.xls)"
        )
        
        file_path = file_dialog[0]
        if file_path:
            self.loaded_files.append(file_path)
            filename = file_path.split('/')[-1]
            self.listFiles.addItem(filename)
    
    def clearlist(self):
        self.listFiles.clear()
        self.loaded_files.clear()
        self.outMessage.clear()

    def showError(self, message):
        msgBox = QtWidgets.QMessageBox()
        msgBox.setIcon(QtWidgets.QMessageBox.Warning)
        msgBox.setText(message)
        msgBox.setWindowTitle("Error")
        msgBox.exec_()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = AppScreen()
    sys.exit(app.exec())
