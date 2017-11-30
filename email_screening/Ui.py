
import sys
from PyQt5.QtWidgets import (QMainWindow,QDialog)
from PyQt5.QtWidgets import (QHBoxLayout,QVBoxLayout)
from PyQt5.QtWidgets import (QLabel,QPushButton,QButtonGroup,QStyle,QApplication,QDesktopWidget)
from PyQt5.Qt  import Qt
from PyQt5.QtGui import QDesktopServices

class Ui(QDialog):
    def __init__(self):
        super().__init__()

        self.l_info =QLabel("Here is example of label") # type: QLabel


        self.initUI()
    def centerWindow(self):
        desktop = QApplication.desktop()  # type: QDesktopWidget
        self.resize(800, 600)
        self.setGeometry(QStyle.alignedRect( \
            Qt.LeftToRight, Qt.AlignCenter, self.size(), desktop.availableGeometry()))

    def initUI(self):
        ''' FIRST LINE '''
        vbox = QVBoxLayout()
        self.setLayout(vbox)
        vbox.addItem()


        self.centerWindow()
        self.setWindowTitle("Email Screening")
        self.show()


