import sys # import System
from PyQt5 import QtWidgets

from script import models_da, views

# Own modules
from script import data_process
from script.utils import utils_ui

class main_process:
    def generate_PA_form_data(self, ui_instance):
        try:
            ui_smartsheet_token = ui_instance.lineEdit_smartsheet_API_token.text()
            ui_smartsheet_PE_member_id = ui_instance.lineEdit_smartsheet_PE_member_ID.text()
            ui_workbook_password = ui_instance.lineEdit_workbook_password.text()
            bool_lock_pa_form = ui_instance.checkBox_lock_workbook.isChecked()
            bool_unhide_reports = ui_instance.checkBox_show_reports.isChecked()
            ui_tablewidget = ui_instance.tableWidget
            ui_progressbar = ui_instance.progressBar

            ui_instance.update_status('', ui_progressbar, 0, 4)
            print('<<<<Generate PA Form>>>>')

            # Smartsheet API
            smartsheet_api = models_da.Smartsheet_apiDA().activate_client(ui_smartsheet_token)

            # Employee data
            ui_instance.update_status('<<Loading team data from Smartsheet>>', ui_progressbar, 1, 4)
            teams = models_da.TeamDA().create_instance(smartsheet_api=smartsheet_api, sheet_ids=[ui_smartsheet_PE_member_id])
            utils_ui().df_to_qtablewidget(teams[int(ui_smartsheet_PE_member_id)].df_display, ui_tablewidget)

            # Employee Function data
            ui_instance.update_status('<<Loading employee function data from Smartsheet>>', ui_progressbar, 2, 4)
            for key, team in teams.items():
                models_da.Employee_FunctionDA().create_instance(smartsheet_api=smartsheet_api, team_instance=team)

            # Excel PA Form
            ui_instance.update_status('<<Exporting to PA Form>>', ui_progressbar, 3, 4)
            for key, team in teams.items():
                data_process.process_PA_form(team_instance=team, password_wb=ui_workbook_password, bool_lock_pa_form=bool_lock_pa_form, bool_unhide_reports=bool_unhide_reports)
            
            ui_instance.update_status('<<Completed>>', ui_progressbar, 4, 4)
            print('<<Completed>>')
            print()
        except Exception as e:
            print("****Error****: General Process: {}".format(str(e)))

    def download_employee_function_data(self, ui_instance):
        try:
            ui_smartsheet_token = ui_instance.lineEdit_smartsheet_API_token.text()
            ui_smartsheet_PE_member_id = ui_instance.lineEdit_smartsheet_PE_member_ID.text()
            ui_download_dir = ui_instance.lineEdit_browse_directory.text()
            ui_tablewidget = ui_instance.tableWidget
            ui_progressbar = ui_instance.progressBar

            ui_instance.update_status('', ui_progressbar, 0, 4)
            print('<<<<Download Function Data>>>>')

            # Smartsheet API
            smartsheet_api = models_da.Smartsheet_apiDA().activate_client(ui_smartsheet_token)

            # Employee data
            ui_instance.update_status('<<Loading team data from Smartsheet>>', ui_progressbar, 1, 4)
            teams = models_da.TeamDA().create_instance(smartsheet_api=smartsheet_api, sheet_ids=[ui_smartsheet_PE_member_id])
            utils_ui().df_to_qtablewidget(teams[int(ui_smartsheet_PE_member_id)].df_display, ui_tablewidget)

            # Employee Function data
            ui_instance.update_status('<<Loading employee function data from Smartsheet>>', ui_progressbar, 2, 4)
            for key, team in teams.items():
                models_da.Employee_FunctionDA().create_instance(smartsheet_api=smartsheet_api, team_instance=team)

            # Download employee function
            ui_instance.update_status('<<Downloading employee function data>>', ui_progressbar, 3, 4)
            for key, team in teams.items():
                data_process.download_employee_function_data(smartsheet_api=smartsheet_api, team_instance=team, download_dir=ui_download_dir)
            
            ui_instance.update_status('<<Completed>>', ui_progressbar, 4, 4)
            print('<<Completed>>')
            print()
        except Exception as e:
            print("****Error****: General Process: {}".format(str(e)))

def main():
    # initialize variables
    path_ui='script//UI//UI_main.ui'

    # Open UI
    app = QtWidgets.QApplication(sys.argv)
    ui = views.Ui(path_ui=path_ui)

    ui.pushButton_generate_data.clicked.connect(lambda: main_process().generate_PA_form_data(ui_instance=ui))
    ui.pushButton_download_function_data.clicked.connect(lambda: main_process().download_employee_function_data(ui_instance=ui))

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()