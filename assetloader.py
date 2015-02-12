import sys, csv, os, json, subprocess, ctypes
from os.path import basename
from PyQt4 import QtGui, QtCore, QtCore
from PyQt4.QtCore import QSettings

##########################################################
################### Application GROUP AREA ###############
##########################################################

class ApplicationGroup(QtGui.QGroupBox):
    def __init__(self, orientation, title, parent=None):
        super(ApplicationGroup, self).__init__(title, parent)
        self.buttonLaunchapplication = QtGui.QPushButton('Launch Application', self)
        self.buttonLaunchapplication.clicked.connect(self.handle_launch_application)
        applicationLayout = QtGui.QBoxLayout(QtGui.QBoxLayout.TopToBottom)
        applicationLayout.addWidget(self.buttonLaunchapplication)
        self.setLayout(applicationLayout)
  
    def handle_launch_application(self):
        application_path = 'C:\\Program Files\\application.exe'

        subprocess.Popen({application_path})

###################################################
################### MAIN WINDOW ###################
###################################################

class Window(QtGui.QWidget):

    ##########################################################
    ################### INIT AND UI LAYOUT ###################
    ##########################################################

    def __init__(self, rows, columns):
        QtGui.QWidget.__init__(self)

        self.table = QtGui.QTableWidget(rows, columns, self)
        self.table.setHorizontalHeaderLabels(['Asset Name', 'Diffuse Color', 'Radius', 'Path'])
        self.header = self.table.horizontalHeader()
        self.header.setStretchLastSection(True)

        self.buttonHelp = QtGui.QPushButton('Help', self)
        self.buttonAdd = QtGui.QPushButton('Add Assets', self)
        self.buttonClear = QtGui.QPushButton('Clear All', self)
        self.buttonOpen = QtGui.QPushButton('Open Asset Configuration', self)
        self.buttonSave = QtGui.QPushButton('Save Asset Configuration', self)
        self.buttonRemoveRow = QtGui.QPushButton('Remove Selected Asset', self)

        self.buttonHelp.clicked.connect(self.handle_help)
        self.buttonAdd.clicked.connect(self.handle_add)
        self.buttonClear.clicked.connect(self.handle_clear)
        self.buttonOpen.clicked.connect(self.handle_open)
        self.buttonSave.clicked.connect(self.handle_save)
        self.table.cellClicked.connect(self.handle_selected_row)
        self.buttonRemoveRow.clicked.connect(self.handle_remove_row)

        self.applicationGroup = applicationGroup(QtCore.Qt.Horizontal, "application")

        layout = QtGui.QVBoxLayout(self)
        layout.addWidget(self.buttonAdd)
        layout.addWidget(self.buttonRemoveRow)
        layout.addWidget(self.buttonClear)
        layout.addWidget(self.table)
        layout.addWidget(self.buttonSave)
        layout.addWidget(self.buttonOpen)
        layout.addWidget(self.applicationGroup)
        layout.addWidget(self.buttonHelp)


    ########################################################
    ################### HELPER FUNCTIONS ###################
    ########################################################
    def convert_color(self, color):
        scale = 255.0
        c = "%.2f, %.2f, %.2f" %(round((color[0]/scale),2), round((color[1]/scale),2), round((color[2]/scale),2))
        return c

    def resize_columns(self):
        self.table.resizeColumnToContents(0)
        self.table.resizeColumnToContents(1)
        self.table.resizeColumnToContents(2)
        self.table.resizeColumnToContents(3)

    def convert_to_relative_path(self, path):
        relativePath = basename(unicode(path))
        return relativePath

    #######################################################
    ################### BUTTON HANDLERS ###################
    #######################################################
    def handle_help(self):
        os.startfile('Help.pdf')

    def handle_selected_row(self):
        col = self.table.currentItem().column()
        if col == 1:
            color = QtGui.QColorDialog.getColor()
            if color.isValid():
                self.table.currentItem().setBackgroundColor(color)
                normalizedColor = self.convert_color(color.getRgb())
                self.table.currentItem().setText(str(normalizedColor))
                self.resize_columns()

    def handle_add(self):
        for path in QtGui.QFileDialog.getOpenFileNames(self):
            relPath = self.convert_to_relative_path(path)
            assetName = basename(unicode(path)).split("_")[0]
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QtGui.QTableWidgetItem(assetName))
            self.table.setItem(row, 1, QtGui.QTableWidgetItem('1.0, 1.0, 1.0'))
            self.table.setItem(row, 2, QtGui.QTableWidgetItem('-1'))
            self.table.setItem(row, 3, QtGui.QTableWidgetItem(assetName + '_*.ext'))
        self.resize_columns()

    def handle_clear(self):
        while self.table.rowCount() > 0:
            for row in range(self.table.rowCount()):
                self.table.removeRow(row)

    def handle_remove_row(self):
        self.table.removeRow(self.table.currentRow())

    def handle_save(self):
        path = QtGui.QFileDialog.getSaveFileName(
            self, 'Save File', '', 'JSON(*.json)')
        if not path.isEmpty():
            with open(unicode(path), 'wb') as outfile:
                data = {}
                for row in range(self.table.rowCount()):
                    rowdata = {}
                    for column in range(self.table.columnCount()):
                        item = self.table.item(row, column)
                        if item is not None:
                            rowdata[unicode(self.table.horizontalHeaderItem(column).text())] = unicode(item.text()).encode('utf8')
                            data[rowdata[unicode(self.table.horizontalHeaderItem(0).text())]] = rowdata   
                json.dump(data, outfile, sort_keys = True, indent = 4)

    def handle_open(self):
        path = QtGui.QFileDialog.getOpenFileName(
            self, 'Open File', '', 'JSON(*.json)')
        if not path.isEmpty():
            with open(unicode(path), 'rb') as stream:
                self.table.setRowCount(0)
                json_data = json.load(stream)
                for item in json_data:
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    self.table.setItem(row, 0, QtGui.QTableWidgetItem(json_data[item]['Asset Name']))
                    self.table.setItem(row, 1, QtGui.QTableWidgetItem(json_data[item]['Diffuse Color']))
                    self.table.setItem(row, 2, QtGui.QTableWidgetItem(json_data[item]['Radius']))
                    self.table.setItem(row, 3, QtGui.QTableWidgetItem(json_data[item]['Path']))
        self.resize_columns()
        self.table.setHorizontalHeaderLabels(['Asset Name', 'Diffuse Color', 'Radius', 'Path'])

############################################
################### MAIN ###################
############################################
if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    window = Window(0, 4)
    window.resize(600, 400)
    window.setWindowTitle('VMRL Asset Loader')
    window.setWindowIcon(QtGui.QIcon('app_icon.ico'))
    window.show()
    app.setWindowIcon(QtGui.QIcon('app_icon.ico'))

    appid = 'yourname.assetloader'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appid)
    sys.exit(app.exec_())