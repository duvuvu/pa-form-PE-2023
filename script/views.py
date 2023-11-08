import os # import OS
import sys # import System

from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QFileDialog, QApplication, QLineEdit

# Own modules
from script.utils import utils_json, utils_outputstream

class Ui(QtWidgets.QMainWindow):
    def __init__(self, path_ui):
        super(Ui, self).__init__() # Call the inherited classes __init__ method
        uic.loadUi(path_ui, self) # Load the .ui file
        sys.stdout = utils_outputstream(self.plainTextEdit_ouput_stream)

        self.lineEdit_smartsheet_API_token.setEchoMode(QLineEdit.Password)
        self.lineEdit_workbook_password.setEchoMode(QLineEdit.Password)
        self.lineEdit_smartsheet_PE_member_ID.setEchoMode(QLineEdit.Password)

        self.pushButton_extension_layout.clicked.connect(lambda: self.toggleWidget(self.plainTextEdit_ouput_stream))

        self.load_default_config([self.lineEdit_smartsheet_API_token, 
                                    self.lineEdit_smartsheet_PE_member_ID, 
                                    self.lineEdit_browse_directory,
                                    self.lineEdit_workbook_password], [self.checkBox_lock_workbook, self.checkBox_show_reports])

        self.toolButton_browse_directory.clicked.connect(lambda: self.browse_and_fill_linetext_folder(self.lineEdit_browse_directory))

        self.pushButton_set_as_default.clicked.connect(lambda: self.set_default_config([self.lineEdit_smartsheet_API_token, 
                                                                                          self.lineEdit_smartsheet_PE_member_ID, 
                                                                                          self.lineEdit_browse_directory,
                                                                                          self.lineEdit_workbook_password], [self.checkBox_lock_workbook, self.checkBox_show_reports]))

        self.progressBar.setHidden(True)
        self.plainTextEdit_ouput_stream.setHidden(True)
        self.show() # Show the GUI

    # Set as Default
    def set_default_config(self, line_edits, check_boxes):
        current_dir = os.getcwd()
        default_dir = os.path.join(current_dir, 'script', 'config', 'Config-Default.json')
        utils_json().save_data_from_json(default_dir, line_edits, check_boxes)
        print('<<<<Saved Default Config>>>>')
        print()

    def load_default_config(self, line_edits, check_boxes):
        current_dir = os.getcwd()
        default_dir = os.path.join(current_dir, 'script', 'config', 'Config-Default.json')
        utils_json().load_data_from_json(default_dir, line_edits, check_boxes)
        print('<<<<Loaded Default Config>>>>')
        print()

    # 1.2. Browse folder=======================================================================
    def browse_folder(self):
        dir = QFileDialog.getExistingDirectory(self, "Browse Folder")
        return dir
        
    def browse_and_fill_linetext_folder(self, linetext):
        path = self.browse_folder()
        if path != '':
            linetext.setText(path)
            print('<<<<Changed Download Directory>>>>')
            print()
        return


    def update_status(self, status, progressBar, numerator, denominator):
        progressBar.setHidden(False)
        progressBar.setFormat(status)
        progressBar.setValue(int(numerator/denominator*100))
        QApplication.processEvents()


    # Toggle extension layout====================================================
    def toggleWidget(self, widget):
        # Toggle the visibility of the widget
        widget.setVisible(not widget.isVisible())

        # Adjust the window height based on the widget's visibility
        height = self.height()
        if widget.isVisible():
            height += widget.height()+6
        else:
            height -= widget.height()+6
        self.resize(self.width(), height)


#============================================================
if __name__ == "__main__":
    pass