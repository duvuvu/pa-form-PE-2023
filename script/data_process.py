import xlwings as xw
import os
from datetime import datetime # datetime library

class process_PA_form:
    def __init__(self, team_instance, password_wb, bool_lock_pa_form):
        for index, row in team_instance.df_data.iterrows():
            try:
                print('<<Exporting to PA Form: ', row['employee_name'], '>>')
                path_wb = row['employee_pa_form_path']
                
                #==Initial variables
                id_employee = row['id_employee']
                employee_name = row['employee_name']
                employee_level = row['employee_level']
                employee_job_title = row['employee_job_title']
                employee_date_join = row['employee_date_join']
                employee_date_review_start = row['employee_date_review_start']
                employee_date_review_end = row['employee_date_review_end']

                employee_function = row['employee_function']
                employee_function_df_classified = employee_function.df_classified
                number_of_task_type = employee_function_df_classified.shape[0]

                col_task_type= 'A'
                dict_col = {'L0': 'F',
                            'L1': 'I',
                            'L2': 'L',
                            'L3': 'O',
                            'L4': 'R',
                            'L5': 'U'}
                
                #--Get workbook path
                path_wb = row['employee_pa_form_path']

                #--Get the Excel application object
                app = xw.App()

                #--Open workbook
                wb = xw.Book(path_wb)

                #--Declair and initialize worksheets
                ws_cover = wb.sheets['PA Form']
                ws_self_recognition = wb.sheets['Self-recognition']
                ws_competency = wb.sheets['Competency Data']
                ws_matrix = wb.sheets['Matrix']

                #--Unprotect workbook and worksheets
                wb.api.Unprotect(Password=password_wb)
                ws_cover.api.Unprotect(Password=password_wb)
                ws_self_recognition.api.Unprotect(Password=password_wb)
                ws_competency.api.Unprotect(Password=password_wb)
                ws_matrix.api.Unprotect(Password=password_wb)

                #--Update personal data
                ws_cover.range('N7').value = int(id_employee)
                ws_cover.range('I6').value = employee_name
                ws_cover.range('R7').value = employee_level
                ws_cover.range('H7').value = employee_job_title
                ws_cover.range('N6').value = datetime.strptime(employee_date_join, '%Y-%m-%d').strftime('%m/%d/%Y')
                ws_cover.range('C10').value = datetime.strptime(employee_date_review_start, '%Y-%m-%d').strftime('%m/%d/%Y')
                ws_cover.range('E10').value = datetime.strptime(employee_date_review_end, '%Y-%m-%d').strftime('%m/%d/%Y')

                # Get and set textbox
                ws_self_recognition.shapes("TextBox 1A").text = employee_name
                ws_self_recognition.shapes("TextBox 3A").text = datetime.strptime(employee_date_review_start, '%Y-%m-%d').strftime('%Y')


                #----Clear old data
                last_row = ws_competency.range('J500').end('up').row # column J
                number_of_OLD_task_type = int((last_row + 1 - 5) / 2)
                ws_cover.range(f"{'X2:AB{}'.format(str(2+number_of_OLD_task_type-1))}").clear_contents()
                if number_of_OLD_task_type-2 > 0:
                    ws_cover.range(f"{'X3:AB{}'.format(3+number_of_OLD_task_type-2-1)}").delete(shift='up')

                ws_competency.range('7:500').clear()

                #----Import new data
                if number_of_task_type-2 > 0:
                    
                    competency_data_FR = 5
                    competency_data_LR = competency_data_FR + number_of_task_type*2 - 1
                    ws_competency.range('5:6').copy()
                    ws_competency.range(f"{'{}:{}'.format(competency_data_FR+2, competency_data_LR)}").paste()

                    rng = ws_cover.range(f"{'X3:AB{}'.format(3+number_of_task_type-2-1)}")
                    rng.insert(shift='down')

                    i = 0
                    for index, row in employee_function_df_classified.iterrows():
                        ws_competency.range(f"{'{}{}'.format(col_task_type, str(competency_data_FR + 2*i))}").value = index
  
                        for key, value in dict_col.items():
                            cell_upper = ws_competency.range(f"{'{}{}'.format(value, str(competency_data_FR + 2*i))}")
                            cell_upper_data = employee_function_df_classified.at[index,f"{'{}_T1'.format(key)}"]
                            cell_upper.value = cell_upper_data
                            if cell_upper_data == '-':
                                cell_upper.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                            else:
                                cell_upper.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
                            
                            cell_lower = ws_competency.range(f"{'{}{}'.format(value, str(competency_data_FR + 1 + 2*i))}")
                            cell_lower_data = employee_function_df_classified.at[index,f"{'{}_T2'.format(key)}"]
                            cell_lower.value = cell_lower_data
                            if cell_lower_data == '-':
                                cell_lower.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                            else:
                                cell_lower.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

                        ws_cover.range(f"{'X{}'.format(2+i)}").formula = "='Competency Data'!A"+str(5+2*i)
                        ws_cover.range(f"{'Y{}'.format(2+i)}").value = 2
                        ws_cover.range(f"{'Z{}'.format(2+i)}").value = 3
                        ws_cover.range(f"{'AA{}'.format(2+i)}").value = 3.5
                        ws_cover.range(f"{'AB{}'.format(2+i)}").formula = "='Competency Data'!B"+str(5+2*i)
                        i += 1

                #--Protect worksheets
                if bool_lock_pa_form:
                    ws_cover.api.Protect(Password=password_wb)
                    ws_self_recognition.api.Protect(Password=password_wb)
                    ws_competency.api.Protect(Password=password_wb)
                    ws_matrix.api.Protect(Password=password_wb)
                    wb.api.Protect(Password=password_wb)

                #--Save and close workbook
                wb.save()
                wb.close()

                #--Close the Excel application
                app.quit()
                
            except Exception as e:
                print("****Error****: {}".format(str(e)))


class download_employee_function_data:
    def __init__(self, smartsheet_api, team_instance, download_dir):
        for index, row in team_instance.df_data.iterrows():
            try:
                print('<<Downloading Employee Function Data: ', row['employee_name'], '>>')
                id_sheet = int(row['employee_function'].id)
                smartsheet_api.client.Sheets.get_sheet_as_csv(id_sheet, os.path.join(download_dir)) # Get sheet as CSV
                
            except Exception as e:
                print("****Error****: Downloading: {}".format(str(e)))

