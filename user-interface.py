from PyQt5.QtWidgets import (QAction, QApplication, QFormLayout, QGroupBox,
                             QLabel, QPushButton, QVBoxLayout, QWidget,
                             QMainWindow, QLineEdit)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from processSurveyResultsFile import processExcelFile

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.createUI()





    def createUI(self):
        # Create window
        self.setWindowTitle('FADC Survey Results File Processor')
        self.resize(700, 400)
        self.setMinimumSize(250, 300)
        # Create central widget and layout
        self._centralWidget = QWidget()
        self._verticalLayout = QVBoxLayout()
        self._centralWidget.setLayout(self._verticalLayout)
        # Set central widget
        self.setCentralWidget(self._centralWidget)
        # Vertically center widgets
        self._verticalLayout.addStretch(1)
        self.addText()
        self.addInputText()
        # Vertically center widgets
        self._verticalLayout.addStretch(1)



    def addText(self):
        messageLabel = QLabel(
            'Please enter the name of the Excel file and the Excel Sheet in the fields below'
        )
        messageLabel.setAlignment(Qt.AlignCenter)
        messageLabel.setFixedWidth(350)
        messageLabel.setMinimumHeight(50)
        messageLabel.setWordWrap(True)
        self._verticalLayout.addWidget(messageLabel, alignment=Qt.AlignCenter)


    def addInputText(self):
        groupBox = QGroupBox()
        groupBox.setFixedWidth(550)

        formLayout = QFormLayout()

        excelFileNameLabel = QLabel('Excel File Name')
        excelFileNameField = QLineEdit(self)
        excelFileNameField.setTextMargins(3, 0, 3, 0)
        excelFileNameField.setMinimumWidth(200)
        excelFileNameField.setMaximumWidth(300)
        excelFileNameField.setClearButtonEnabled(True)
        
        excelSheetNameLabel = QLabel('Excel Sheet Name')
        excelSheetNameField = QLineEdit(self)
        excelSheetNameField.setTextMargins(3, 0, 3, 0)
        excelSheetNameField.setMinimumWidth(200)
        excelSheetNameField.setMaximumWidth(300)
        excelSheetNameField.setClearButtonEnabled(True)

        processButton = QPushButton('Process File', self)
        processButton.setMaximumWidth(100)
        processButton.clicked.connect(lambda: self.process_button_click(excelFileNameField, excelSheetNameField))

        submitLabel = QLabel('Open Your Vault')
        submitField = QPushButton()

        formLayout.addRow(excelFileNameLabel, excelFileNameField)
        formLayout.addRow(excelSheetNameLabel, excelSheetNameField)
        formLayout.addRow(processButton)

        groupBox.setLayout(formLayout)
        self._verticalLayout.addWidget(groupBox, alignment=Qt.AlignCenter)

    def process_button_click(excelFileNameField, excelSheetNameField, self):
        fileName = excelFileNameField.text()
        sheetName = excelSheetNameField.text()
        if fileName is not None and fileName != '' and fileName is not None and fileName != '':
            processExcelFile(fileName, sheetName)


if (__name__ == '__main__'):

    application = QApplication([])
    mainWindow = MainWindow()
    mainWindow.show()
    application.exec()