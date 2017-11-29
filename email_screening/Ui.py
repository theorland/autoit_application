
import sys
from PyQt5.QtWidgets import (QMainWindow,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QButtonGroup,QStyle,QApplication,QDesktopWidget)
from PyQt5.Qt  import Qt
from PyQt5.QtGui import QDesktopServices

class Ui(QMainWindow):
    def __init__(self):
        super().__init__()

        self.l_info =QLabel("Here is example of label",self) # type: QLabel


        self.initUI()
    def centerWindow(self):
        desktop = QApplication.desktop()  # type: QDesktopWidget
        self.resize(800, 600)
        self.setGeometry(QStyle.alignedRect( \
            Qt.LeftToRight, Qt.AlignCenter, self.size(), desktop.availableGeometry()))

    def initUI(self):
        ''' FIRST LINE '''
        hbox = QHBoxLayout()

        hbox.addStretch()
        hbox.addWidget(self.l_info,2)
        hbox.addStretch()

        btn1 = QPushButton("reconnect")

        hbox.addWidget(btn1)
        hbox.addStretch()

        btn2 = QPushButton("refresh")

        hbox.addWidget(btn2)
        hbox.addStretch()

        vbox = QVBoxLayout()

        vbox.addStretch()
        vbox.addLayout(hbox)

        self.statusBar().showMessage('Ready')

        statusB = self.statusBar() # type: QStatusBar
        self.setLayout(vbox)

        self.centerWindow()
        self.setWindowTitle("Email Screening")
        self.show()


