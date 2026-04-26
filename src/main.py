import sys
import os
import subprocess

# Add the parent directory to sys.path so 'src' can be resolved
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PyQt5 import QtWidgets, QtCore, QtGui
from src.excel_writer import writeExcel

class AppScreen(QtWidgets.QMainWindow):
    def __init__(self):
        super(AppScreen, self).__init__()
        self.setWindowTitle("ISO-NE Feasibility Tool")
        self.resize(700, 800)
        self.setMinimumSize(500, 600)
        
        self.is_dark_mode = True
        
        # Keep track of absolute file paths loaded
        self.loaded_files = []
        
        # Build UI
        self.init_ui()
        self.apply_theme()
        self.show()

    def init_ui(self):
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QtWidgets.QVBoxLayout(central_widget)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(25)
        
        # --- Header ---
        header_layout = QtWidgets.QHBoxLayout()
        self.titleApp = QtWidgets.QLabel("Feasibility Study Tool")
        self.titleApp.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.titleApp.setObjectName("titleApp")
        
        self.theme_toggle = QtWidgets.QPushButton("Light Mode")
        self.theme_toggle.setObjectName("themeToggle")
        self.theme_toggle.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.theme_toggle.clicked.connect(self.toggle_theme)
        
        header_layout.addWidget(self.titleApp, stretch=1)
        header_layout.addWidget(self.theme_toggle)
        main_layout.addLayout(header_layout)
        
        # --- Control Panel (Glass Card) ---
        control_panel = QtWidgets.QFrame()
        control_panel.setObjectName("glassPanel")
        control_layout = QtWidgets.QGridLayout(control_panel)
        control_layout.setSpacing(15)
        control_layout.setContentsMargins(25, 25, 25, 25)
        
        # Bus Number
        bus_label = QtWidgets.QLabel("Bus Number:")
        self.busNoInput = QtWidgets.QLineEdit()
        self.busNoInput.setPlaceholderText("Enter Bus Number")
        
        # Project Size
        proj_label = QtWidgets.QLabel("Project Size (MW):")
        self.projSizeInput = QtWidgets.QLineEdit()
        self.projSizeInput.setPlaceholderText("Enter Project Size")
        
        # Charging / Discharging
        mode_label = QtWidgets.QLabel("Operation Mode:")
        mode_layout = QtWidgets.QHBoxLayout()
        self.charge = QtWidgets.QRadioButton("Charging")
        self.discharge = QtWidgets.QRadioButton("Discharging")
        self.discharge.setChecked(True)
        mode_layout.addWidget(self.charge)
        mode_layout.addWidget(self.discharge)
        
        control_layout.addWidget(bus_label, 0, 0)
        control_layout.addWidget(self.busNoInput, 0, 1)
        control_layout.addWidget(proj_label, 1, 0)
        control_layout.addWidget(self.projSizeInput, 1, 1)
        control_layout.addWidget(mode_label, 2, 0)
        control_layout.addLayout(mode_layout, 2, 1)
        
        main_layout.addWidget(control_panel)
        
        # --- File Management (Glass Card) ---
        file_panel = QtWidgets.QFrame()
        file_panel.setObjectName("glassPanel")
        file_layout = QtWidgets.QVBoxLayout(file_panel)
        file_layout.setSpacing(15)
        file_layout.setContentsMargins(25, 25, 25, 25)
        
        btn_layout = QtWidgets.QHBoxLayout()
        self.load = QtWidgets.QPushButton("📂 Load Excel File")
        self.load.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.clear = QtWidgets.QPushButton("🗑️ Clear List")
        self.clear.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        btn_layout.addWidget(self.load)
        btn_layout.addWidget(self.clear)
        
        self.listFiles = QtWidgets.QListWidget()
        self.listFiles.setObjectName("listFiles")
        
        file_layout.addLayout(btn_layout)
        file_layout.addWidget(self.listFiles)
        
        main_layout.addWidget(file_panel)
        
        # --- Action Footer ---
        footer_layout = QtWidgets.QVBoxLayout()
        footer_layout.setAlignment(QtCore.Qt.AlignCenter)
        footer_layout.setSpacing(15)
        
        self.run = QtWidgets.QPushButton("🚀 RUN REDISPATCH")
        self.run.setObjectName("runBtn")
        self.run.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.run.setMinimumHeight(55)
        self.run.setMinimumWidth(300)
        
        self.outMessage = QtWidgets.QLabel("")
        self.outMessage.setObjectName("outMessage")
        self.outMessage.setAlignment(QtCore.Qt.AlignCenter)
        
        self.open_folder_btn = QtWidgets.QPushButton("📂 Show Output in Folder")
        self.open_folder_btn.setObjectName("openFolderBtn")
        self.open_folder_btn.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        
        footer_layout.addWidget(self.run)
        footer_layout.addWidget(self.outMessage)
        footer_layout.addWidget(self.open_folder_btn, alignment=QtCore.Qt.AlignCenter)
        
        main_layout.addLayout(footer_layout)
        
        main_layout.addStretch()
        
        # Connections
        self.run.clicked.connect(self.runFunction)
        self.load.clicked.connect(self.loadFunction)
        self.clear.clicked.connect(self.clearlist)
        
    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        if self.is_dark_mode:
            self.theme_toggle.setText("Light Mode")
        else:
            self.theme_toggle.setText("Dark Mode")
        self.apply_theme()
        
    def apply_theme(self):
        # Cross-platform glassmorphic aesthetic using QSS
        if self.is_dark_mode:
            bg_color = "#1e1e2e"
            panel_bg = "rgba(255, 255, 255, 0.05)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "#cdd6f4"
            input_bg = "rgba(0, 0, 0, 0.3)"
            btn_bg = "#89b4fa"
            btn_text = "#11111b"
            btn_hover = "#b4befe"
            run_btn_bg = "#a6e3a1"
            run_btn_text = "#11111b"
            run_btn_hover = "#94e2d5"
        else:
            bg_color = "#eff1f5"
            panel_bg = "rgba(255, 255, 255, 0.7)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#4c4f69"
            input_bg = "rgba(255, 255, 255, 0.9)"
            btn_bg = "#1e66f5"
            btn_text = "#eff1f5"
            btn_hover = "#7287fd"
            run_btn_bg = "#40a02b"
            run_btn_text = "#eff1f5"
            run_btn_hover = "#53b544"

        qss = f"""
        QMainWindow {{
            background-color: {bg_color};
        }}
        QWidget {{
            color: {text_color};
            font-family: BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            font-size: 15px;
        }}
        QLabel#titleApp {{
            font-size: 28px;
            font-weight: bold;
            color: {btn_bg};
        }}
        QFrame#glassPanel {{
            background-color: {panel_bg};
            border: 1px solid {border_color};
            border-radius: 12px;
        }}
        QLineEdit {{
            background-color: {input_bg};
            border: 1px solid {border_color};
            border-radius: 6px;
            padding: 8px;
            color: {text_color};
        }}
        QLineEdit:focus {{
            border: 1px solid {btn_bg};
        }}
        QPushButton {{
            background-color: {btn_bg};
            color: {btn_text};
            border: none;
            border-radius: 6px;
            padding: 10px 16px;
            font-weight: bold;
        }}
        QPushButton:hover {{
            background-color: {btn_hover};
        }}
        QPushButton#themeToggle {{
            background-color: transparent;
            color: {text_color};
            border: none;
            padding: 8px 16px;
            font-weight: bold;
            border-radius: 15px;
        }}
        QPushButton#themeToggle:hover {{
            background-color: {panel_bg};
        }}
        QPushButton#runBtn {{
            background-color: {run_btn_bg};
            color: {run_btn_text};
            font-size: 16px;
            border-radius: 27px; /* Pill shape */
        }}
        QPushButton#runBtn:hover {{
            background-color: {run_btn_hover};
        }}
        QListWidget {{
            background-color: {input_bg};
            border: 1px solid {border_color};
            border-radius: 6px;
            padding: 5px;
            color: {text_color};
        }}
        QLabel#outMessage {{
            font-size: 16px;
            font-weight: bold;
            color: {run_btn_bg};
        }}
        QPushButton#openFolderBtn {{
            background-color: transparent;
            color: {btn_bg};
            border: 1px solid {btn_bg};
            font-size: 14px;
            padding: 6px 16px;
            border-radius: 15px;
        }}
        QPushButton#openFolderBtn:hover {{
            background-color: {panel_bg};
        }}
        QRadioButton {{
            spacing: 8px;
        }}
        QRadioButton::indicator {{
            width: 16px;
            height: 16px;
            border-radius: 9px;
        }}
        QRadioButton::indicator:unchecked {{
            border: 2px solid #888888;
            background-color: {input_bg};
        }}
        QRadioButton::indicator:checked {{
            background-color: {btn_bg};
            border: 4px solid {input_bg};
        }}
        """
        self.setStyleSheet(qss)

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
            output_file = writeExcel(file_path, inputParameters)
            # Extracted filename for display
            filename = os.path.basename(output_file)
            self.outMessage.setText(f'✅ {filename} created successfully!')
            self.last_output_file = output_file
            self.open_folder_btn.setVisible(True)
        except Exception as e:
            self.showError(f"Error during execution: {str(e)}")

    def open_output_folder(self):
        if not hasattr(self, 'last_output_file') or not self.last_output_file:
            return
            
        path = self.last_output_file
        if sys.platform == 'win32':
            subprocess.run(['explorer', '/select,', os.path.normpath(path)])
        elif sys.platform == 'darwin':
            subprocess.run(['open', '-R', path])
        else:
            subprocess.run(['xdg-open', os.path.dirname(path)])

    def loadFunction(self):
        file_dialog = QtWidgets.QFileDialog.getOpenFileName(
            parent=self,
            caption='Select the Excel file',
            filter="Excel Files (*.xlsx *.xls)"
        )
        
        file_path = file_dialog[0]
        if file_path:
            self.loaded_files.append(file_path)
            filename = os.path.basename(file_path)
            self.listFiles.addItem(filename)
    
    def clearlist(self):
        self.listFiles.clear()
        self.loaded_files.clear()
        self.outMessage.setText("")

    def showError(self, message):
        msgBox = QtWidgets.QMessageBox()
        msgBox.setIcon(QtWidgets.QMessageBox.Warning)
        msgBox.setText(message)
        msgBox.setWindowTitle("Error")
        self.apply_theme_to_msgbox(msgBox)
        msgBox.exec_()
        
    def apply_theme_to_msgbox(self, msgBox):
        # Quick hack to apply the style to popups
        msgBox.setStyleSheet(self.styleSheet())

if __name__ == '__main__':
    # Set DPI awareness for high-res screens (Cross-Platform compatibility)
    if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    
    app = QtWidgets.QApplication(sys.argv)
    window = AppScreen()
    sys.exit(app.exec())
