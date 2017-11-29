from PyQt5.QtWidgets import QApplication
import sys
from Ui import Ui

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Ui()
    sys.exit(app.exec_())