from script import models_da

class Smartsheet_api:
    def __init__(self, smartsheet_token, smartsheet_client):
        self.__token = smartsheet_token
        self.__client = smartsheet_client

    @property
    def token(self):
        return self.__token

    @property
    def client(self):
        return self.__client

class Team:
    def __init__(self, id, smartsheet_sheet_api):
        self.__id = id
        self.__SS_sheet_api = smartsheet_sheet_api
        self.__df_display = models_da.TeamDA().create_df_display(smartsheet_team_sheet_api= self.__SS_sheet_api)
        self.__df_data = models_da.TeamDA().import_n_clean_data(smartsheet_team_sheet_api= self.__SS_sheet_api)

    @property
    def id(self):
        return self.__id

    @property
    def SS_sheet_api(self):
        return self.__SS_sheet_api

    @property
    def df_display(self):
        return self.__df_display

    @property
    def df_data(self):
        return self.__df_data
    
    @df_data.setter
    def df_data(self, df_data):
        self.__df_data = df_data


class Employee_Function:
    def __init__(self, id, smartsheet_sheet_api, team_instance):
        self.__id = id
        self.__SS_sheet_api = smartsheet_sheet_api
        self.__df_data = models_da.Employee_FunctionDA().import_n_clean_data(smartsheet_employee_function_sheet_api= self.__SS_sheet_api)
        self.__df_classified = models_da.Employee_FunctionDA().classify_data(df_data=self.__df_data)

        team_instance.df_data.loc[team_instance.df_data['id_sheet'].isin([self.id]), 'employee_function'] = self

    @property
    def id(self):
        return self.__id

    @property
    def SS_sheet_api(self):
        return self.__SS_sheet_api

    @property
    def df_data(self):
        return self.__df_data
    
    @property
    def df_classified(self):
        return self.__df_classified