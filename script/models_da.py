from singleton_decorator import singleton

import pandas as pd
import re # regular expression library

# Own modules
from script import models
from script.utils import utils_smartsheet, utils_dataframe, utils_xlwings

@singleton
class Smartsheet_apiDA:
    def activate_client(self, smartsheet_token):
        smartsheet_client = utils_smartsheet().get_smartsheet_client(smartsheet_token)
        smartsheet_api = models.Smartsheet_api(smartsheet_token, smartsheet_client)
        return smartsheet_api
        
@singleton
class TeamDA:
    def __init__(self):
        self.__teams = {}

    def create_instance(self, smartsheet_api, sheet_ids):
        print('<<Loading team data from Smartsheet: Product Engineering>>')
        for sheet_id in sheet_ids:
            smartsheet_team_sheet_api = utils_smartsheet().get_smartsheet_sheet(smartsheet_api.client, int(sheet_id))
            team = models.Team(id=sheet_id, smartsheet_sheet_api=smartsheet_team_sheet_api)
            self.__teams[int(sheet_id)] = team
        return self.__teams

    def create_df_display(self, smartsheet_team_sheet_api):
        team_df = utils_smartsheet().smartsheet_sheet_to_dataframe(sheet_api=smartsheet_team_sheet_api)
        criteria = {'PA Update': True}
        team_df = utils_dataframe().filter_df(team_df, criteria)
        team_df.drop(['PA Update', 'Function Smartsheet ID', 'Link to PA Form'], axis=1, inplace=True)
        return team_df

    def import_n_clean_data(self, smartsheet_team_sheet_api):
        team_df = utils_smartsheet().smartsheet_sheet_to_dataframe(sheet_api=smartsheet_team_sheet_api)

        dict_rename = {'PA Update': 'employee_pa_update_bool','Employee Name': 'employee_name', 'Service': 'employee_service', 'Level': 'employee_level',
                     'Job Title': 'employee_job_title', 'Joining Date': 'employee_date_join',
                     'PMA Data Start Date': 'employee_date_pma_start', 'PMA Data End Date': 'employee_date_pma_end',
                     'Review Period Start Date': 'employee_date_review_start', 'Review Period End Date': 'employee_date_review_end',
                     'Link to PA Form': 'employee_pa_form_path', 'Notes': 'employee_notes',
                     'Code': 'id_employee', 'Function Smartsheet ID': 'id_sheet'}
        team_df = team_df.rename(columns=dict_rename)
        
        criteria = {'employee_pa_update_bool': True}
        team_df = utils_dataframe().filter_df(team_df, criteria)
        team_df['id_employee'] = team_df['id_employee'].astype(int)
        team_df.drop(['employee_pa_update_bool', 'employee_date_pma_start', 'employee_date_pma_end', 'employee_notes'], axis=1, inplace=True)
        
        team_df = team_df.assign(employee_function=None)
        
        return team_df

@singleton
class Employee_FunctionDA:
    def create_instance(self, smartsheet_api, team_instance):
        for index, row in team_instance.df_data.iterrows():
            employee_name = row['employee_name']
            id_sheet = row['id_sheet']
            print('<<Loading employee function data from Smartsheet: ', employee_name, '>>')
            smartsheet_employee_function_sheet_api = utils_smartsheet().get_smartsheet_sheet(smartsheet_api.client, int(id_sheet))
            employee_function = models.Employee_Function(id=id_sheet, smartsheet_sheet_api=smartsheet_employee_function_sheet_api, team_instance=team_instance)

    def import_n_clean_data(self, smartsheet_employee_function_sheet_api):
        employee_function_df = utils_smartsheet().smartsheet_sheet_to_dataframe(sheet_api=smartsheet_employee_function_sheet_api)

        dict_rename = {'ES Number': 'task_number', 'Project Name': 'task_name', 'Project Type': 'task_type',
                     'Due Date': 'task_date_due', 'Sent to US Date': 'task_date_sent',  'US Review Date': 'task_date_review',
                     "Customer's Response Email Capture": 'task_email_review',
                     'Proficiency':'task_proficiency', 'Completed CEI# (If any)':'task_CEI_number',
                     'Improvement Idea (If any)':'task_improvement_idea', 'Compliment from customer(s) for Improvement Idea': 'task_improvement_idea_compliment',
                     'Ready to Review': 'task_ready_to_reivew_bool', 'Approved by': 'task_approved_bool',
                     "Reviewer's Comments": 'task_reviewer_comments', 'Notes': 'task_notes',
                     'Meet Deadline': 'task_meet_deadline', 'Modified by': 'task_person_modify', 'Modified Date': 'task_date_modify'}
        employee_function_df.rename(columns = dict_rename, inplace = True)

        employee_function_df = utils_dataframe().filter_df_not_null(employee_function_df, 'task_approved_bool')
        employee_function_df['id_task'] = employee_function_df['task_number'] + ' ' + employee_function_df['task_name'] + ' (' + employee_function_df['task_date_sent'] + ')'
        employee_function_df = employee_function_df[['id_task'] + list(employee_function_df.columns[:-1])]
        employee_function_df.drop(['task_number', 'task_name', 'task_date_sent'], axis=1, inplace=True)
        employee_function_df.drop(['task_date_due', 'task_date_review', "task_email_review", 'task_ready_to_reivew_bool', 'task_approved_bool', 
                                   "task_reviewer_comments", 'task_notes', 'task_meet_deadline', 'task_person_modify', 'task_date_modify'], axis=1, inplace=True)
        # employee_function_df.insert(0, 'id_sheet', sheet_id)

        return employee_function_df

    
    def classify_data(self, df_data):
        index_values = utils_dataframe().get_unique_values(df_data, 'task_type')
        column_headers = ['L0_T1', 'L0_T2', 'L1_T1', 'L1_T2', 'L2_T1', 'L2_T2', 'L3_T1', 'L3_T2', 'L4_T1','L4_T2', 'L5_T1', 'L5_T2']
        df_classfied = pd.DataFrame(data='-', index=index_values, columns=column_headers)

        # Proficiency
        df_temp = utils_dataframe().filter_df_not_null(df_data[['id_task', 'task_type', 'task_proficiency']], 'task_proficiency')
        for index, row in df_temp.iterrows():
            task_proficiency = row['task_proficiency']
            if re.match(r'^L1\.', task_proficiency):
                col = 'L1_T1'
            elif re.match(r'^L2\.', task_proficiency):
                col = 'L2_T1'
            elif re.match(r'^L3\.', task_proficiency):
                col = 'L3_T1'
            elif re.match(r'^L4\.', task_proficiency):
                col = 'L4_T1'
            elif re.match(r'^L5\.1\.', task_proficiency):
                col = 'L5_T1'
            elif re.match(r'^L5\.2\.', task_proficiency):
                col = 'L5_T2'
            if df_classfied.at[row['task_type'], col] == '-':
                df_classfied.at[row['task_type'], col] = '● ' + row['id_task'] + '\n'
            else:
                df_classfied.at[row['task_type'], col] += '● ' + row['id_task'] + '\n'
            
        # CEI
        df_temp = utils_dataframe().filter_df_not_null(df_data[['id_task', 'task_type', 'task_CEI_number']], 'task_CEI_number')
        col = 'L4_T2'
        for index, row in df_temp.iterrows():
            if df_classfied.at[row['task_type'], col] == '-':
                df_classfied.at[row['task_type'], col] = '● ' + str(row['id_task']) + ': ' + str(row['task_CEI_number']) + '\n'
            else:
                df_classfied.at[row['task_type'], col] += '● ' + str(row['id_task']) + ': ' + str(row['task_CEI_number']) + '\n'

        #Improvement Idea
        df_temp = utils_dataframe().filter_df_not_null(df_data[['id_task', 'task_type', 'task_improvement_idea', 'task_improvement_idea_compliment']], 'task_improvement_idea')
        col = 'L3_T2'
        for index, row in df_temp.iterrows():
            if df_classfied.at[row['task_type'], col] == '-':
                if row['task_improvement_idea_compliment'] != None:
                    df_classfied.at[row['task_type'], col] = '● ' + str(row['id_task']) + ': ' + str(row['task_improvement_idea']) + '\n'
                else:
                    df_classfied.at[row['task_type'], col] = '●● ' + str(row['id_task']) + ': ' + str(row['task_improvement_idea']) + '\n'
            else:
                if row['task_improvement_idea_compliment'] != None:
                    df_classfied.at[row['task_type'], col] += '● ' + str(row['id_task']) + ': ' + str(row['task_improvement_idea']) + '\n'
                else:
                    df_classfied.at[row['task_type'], col] += '●● ' + str(row['id_task']) + ': ' + str(row['task_improvement_idea']) + '\n'
        
        df_temp = None

        return df_classfied