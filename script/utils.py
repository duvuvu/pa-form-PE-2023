import os
import math
import json
import smartsheet
import pandas as pd
import xlwings as xw

from PyQt5.QtGui import QTextCursor
from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtWidgets import QTableView, QHeaderView, QTableWidgetItem, QCheckBox

class utils_smartsheet:
    def get_smartsheet_client(self, smartsheet_token):
        smartsheet_client = smartsheet.Smartsheet(smartsheet_token)
        return smartsheet_client

    # Get SmartSheet sheet
    def get_smartsheet_sheet(self, smartsheet_client, sheet_id):
        smartsheet_sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
        return smartsheet_sheet

    # Convert SmartSheet sheet to DataFrame
    def smartsheet_sheet_to_dataframe(self, sheet_api):
        col_names = [col.title for col in sheet_api.columns]
        rows = []
        for row in sheet_api.rows:
            cells = []
            for cell in row.cells:
                    cells.append(cell.value)
            rows.append(cells)
        df = pd.DataFrame(rows, columns = col_names)
        return df
    
    # Download SmartSheet sheet to csv
    def download_smartsheet_to_csv(self, smartsheet_client, sheet_id, directory):
        sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
        sheet_name = sheet.name
        sheet_csv = smartsheet_client.Sheets.get_sheet_as_csv(sheet_id, os.path.join(directory)) # Get sheet as CSV

class utils_dataframe:
    # Filtered dataframe
    def filter_df(self, df, criteria_dict):
        filters = [] # create a list of filters based on the criteria dictionary
        for column, values in criteria_dict.items():
            if isinstance(values, bool):
                # if the value is a boolean, use it directly as a filter
                filters.append(df[column] == values)
            elif isinstance(values, str):
                # if the value is a string, treat it as a regular expression
                filters.append(df[column].str.contains(values, case=False, na=False))
            else:
                # if the value is a list, use the isin method
                filters.append(df[column].isin(values))
        filtered_df = df.loc[pd.concat(filters, axis=1).all(axis=1)] # apply the filters to the DataFrame
        return filtered_df

    # Filtered dataframe not null
    def filter_df_not_null(self, df, column_name):
        df_filtered = df[df[column_name].notnull()]
        return df_filtered

    # Get unique values of columns
    def get_unique_values(self, df, column_name):
        unique_values = df[column_name].unique().tolist()
        return unique_values

class utils_json:
    def save_data_from_json(self, file_path, line_edits=[], check_boxes=[]):
        ui_data = {}
        for le in line_edits:
            ui_data[le.objectName()] = le.text()
        for check_box in check_boxes:
            ui_data[check_box.objectName()] = check_box.isChecked()
        with open(file_path, 'w') as f:
            json.dump(ui_data, f, indent=4)


    def load_data_from_json(self, file_path, line_edits=[], check_boxes=[]):
        with open(file_path, 'r') as f:
            ui_data = json.load(f)
        for le in line_edits:
            if le.objectName() in ui_data:
                le.setText(ui_data[le.objectName()])
        for check_box in check_boxes:
            check_box.setChecked(ui_data.get(check_box.objectName(), False))

class utils_outputstream:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        cursor = self.text_widget.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(text)
        self.text_widget.setTextCursor(cursor)
        self.text_widget.ensureCursorVisible()
        QCoreApplication.processEvents()

    def flush(self):
        pass

class utils_xlwings:
    def open_wb(self, path_workbook):
        workbook = xw.Book(path_workbook)
        return workbook

    def close_wb(self, workbook, save_bool=True):
        workbook.save()
        workbook.close()

    def protect_ws(self, worksheets, password):
        for worksheet in worksheets:
            try:
                worksheet.api.Protect(Password=password)
            finally:
                pass

    def unprotect_ws(self, worksheets, password):
        for worksheet in worksheets:
            try:
                worksheet.api.Unprotect(Password=password)
            finally:
                pass

class utils_ui:
    # Add checkbox to qtablewiget item
    def add_checkbox(self, table_widget, row, col, state=False):
        checkbox_widget = QCheckBox()
        if not state:
            checkbox_widget.setCheckState(Qt.Unchecked)
        else:
            checkbox_widget.setCheckState(Qt.Checked)
        table_widget.setCellWidget(row, col, checkbox_widget)

    # Dataframe to qtablewidget
    def df_to_qtablewidget(self, df, table_widget, round_digits=None):        
        table_widget.setRowCount(df.shape[0])  # set the number of rows
        table_widget.setColumnCount(df.shape[1])  # set the number of columns
        table_widget.setHorizontalHeaderLabels(df.columns)  # set the horizontal header labels
        table_widget.setVerticalHeaderLabels(df.index.astype(str))  # set the vertical header labels
        # Set the cell values
        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                value_df = df.iloc[row, col]
                if value_df == '-' or value_df is None:
                    item = None
                    table_widget.setItem(row, col, item)
                elif value_df is True:
                    self.add_checkbox(table_widget, row, col, state=True)
                elif value_df is False:
                    self.add_checkbox(table_widget, row, col, state=False)
                else:
                    # Remove decimal part from numbers
                    if isinstance(value_df, (int, float)):
                        if math.isnan(value_df):
                            value_df = ""
                        else:
                            value_df = int(value_df)
                    item = QTableWidgetItem(str(value_df))
                    table_widget.setItem(row, col, item)


        # resize the columns to fit the contents
        header = table_widget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        
        # set some additional properties
        table_widget.setAlternatingRowColors(True)
        table_widget.setSelectionBehavior(QTableView.SelectRows)
        table_widget.setVerticalScrollMode(QTableView.ScrollPerPixel)
        table_widget.setHorizontalScrollMode(QTableView.ScrollPerPixel)
        # table.setSortingEnabled(True)
        # return table
        QCoreApplication.processEvents()