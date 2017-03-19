import sys
from PyQt4 import QtGui, uic

 

form_class = uic.loadUiType("acercade.ui")[0]

class acercade(QtGui.QMainWindow, form_class ):
    def __init__(self, parent):
        QtGui.QMainWindow.__init__(self, parent)
        self.setupUi(self)
        