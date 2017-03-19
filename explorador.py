
from PyQt4 import QtGui
import os, sys

class PrettyWidget(QtGui.QWidget):
    def __init__(self):
        super(PrettyWidget, self).__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(600, 300, 400, 200)
        self.setWindowTitle('Single Browse')

        btn = QtGui.QPushButton('Browse\n(SINGLE)', self)
        btn.resize(btn.sizeHint())
        btn.clicked.connect(self.SingleBrowse)
        btn.move(150, 100)

        self.show()

    def SingleBrowse(self):
        filePath = QtGui.QFileDialog.getOpenFileName(self,
                                                     'Single File',
                                                     "~/Desktop/",
                                                     '*.xlsx')

        self.arch=filePath
        print(str(self.arch))



def main():
    app = QtGui.QApplication(sys.argv)
    w = PrettyWidget()
    app.exec_()


if __name__ == '__main__':
    main()