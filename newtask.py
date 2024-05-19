import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from googleapiclient.discovery import build
from google.oauth2 import service_account

# Define the scopes for Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Path to the service account key file
XLFolder = "xlFile.txt"
SERVICE_FILE = 'service_account.json'

# Function to select a folder
def select_folder():
    if os.path.exists(XLFolder):
        with open(XLFolder, 'r') as f:
            folder_path = f.read().strip()
    else:
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        folder_path = filedialog.askdirectory(title="Select Folder Containing Excel Files")
        with open(XLFolder, 'w') as f:
            f.write(folder_path)
    return folder_path if folder_path else None

# Function to authenticate and build the Google Sheets API service
def authenticate_gsheets():
    credentials = service_account.Credentials.from_service_account_file(SERVICE_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=credentials)
    return service

# Function to load all Excel files as DataFrames based on Task file
def load_excel_files(task_file, folder_path):
    task_df = pd.read_excel(task_file, header=1)
    
    # Print column names for debugging
    print("Column names in the task file:", task_df.columns.tolist())
    
    source_files = task_df['Source Excell Spreadsheet'].dropna().unique()
    dataframes = {}

    critical_columns = ['Source Excell Spreadsheet', 'Google Sheet ID', 'Tab', 'Starting CELL', 'Ending CELL * means all']
    task_df.dropna(subset=critical_columns, inplace=True)
    
    for file_name in source_files:
        full_path = os.path.join(folder_path, file_name)
        if os.path.exists(full_path):
            df_name = os.path.splitext(file_name)[0]  # Use the file name without extension as the DataFrame name
            dataframes[df_name] = pd.read_excel(full_path, header=1, sheet_name=None).popitem()[1]
        else:
            print(f"File {full_path} does not exist.")
    
    return task_df, dataframes

# Helper function to convert Excel cell notation to DataFrame indices
def excel_cell_to_indices(cell):
    import re
    match = re.match(r"([A-Z]+)([0-9]+)", cell.upper())
    column = match.group(1)
    row = int(match.group(2)) - 3  # Excel rows are 1-based, convert to 0-based
    column_number = 0
    for char in column:
        column_number = column_number * 26 + (ord(char) - ord('A')) + 1
    return row, column_number - 1

# Helper function to convert DataFrame indices back to Excel cell notation
def indices_to_excel_cell(row, col):
    excel_col = ""
    while col >= 0:
        excel_col = chr(col % 26 + ord('A')) + excel_col
        col = col // 26 - 1
    return f"{excel_col}{row + 1}"

# Function to extract data from a specified cell range in a DataFrame
def extract_data_from_range(df, start_cell, end_cell):
    start_row, start_col = excel_cell_to_indices(start_cell)
    if end_cell == '*':
        end_row, end_col = df.shape[0] - 1, df.shape[1] - 1
    else:
        end_row, end_col = excel_cell_to_indices(end_cell)
    # Clean the data to replace NaN with an empty string
    data = df.iloc[start_row:end_row+1, start_col:end_col+1].replace({np.nan: ''}).values.tolist()
    return data

# Function to duplicate a sheet
def duplicate_sheet(service, spreadsheet_id, source_sheet_name, new_sheet_name):
    # Get the sheet ID of the source sheet
    sheets = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get('sheets', [])
    source_sheet_id = None
    for sheet in sheets:
        if sheet['properties']['title'] == source_sheet_name:
            source_sheet_id = sheet['properties']['sheetId']
            break
    
    if source_sheet_id is None:
        print(f"Source sheet {source_sheet_name} not found.")
        return None
    
    # Duplicate the sheet
    duplicate_request = {
        "requests": [
            {
                "duplicateSheet": {
                    "sourceSheetId": source_sheet_id,
                    "newSheetName": new_sheet_name
                }
            }
        ]
    }
    
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=duplicate_request
    ).execute()
    
    print(f"Duplicated sheet {source_sheet_name} to {new_sheet_name}.")
    return new_sheet_name

# Function to update Google Sheets with data from DataFrames
def update_google_sheets(service, task_df, dataframes):
    for _, row in task_df.iterrows():
        file_name = row['Source Excell Spreadsheet']
        gsheet_id = row['Google Sheet ID']
        tab_name = row['Tab']
        xl_start_cell = row['Starting CELL']
        xl_end_cell = row['Ending CELL * means all']

        df_name = os.path.splitext(file_name)[0]
        
        if df_name in dataframes:
            df = dataframes[df_name]
            data_values = extract_data_from_range(df, xl_start_cell, xl_end_cell)
            
            gl_start_row, gl_start_col = excel_cell_to_indices(row['Starting Cell'])
            gl_end_row, gl_end_col = gl_start_row + len(data_values) - 1, gl_start_col + len(data_values[0]) - 1
            
            gl_start_cell = indices_to_excel_cell(gl_start_row, gl_start_col)
            gl_end_cell = indices_to_excel_cell(gl_end_row, gl_end_col)

            if tab_name == 'Create NEW':
                duplicate_tab = row['Duplicate Tab']
                new_sheet_name = f"{duplicate_tab}_copy"
                tab_name = duplicate_sheet(service, gsheet_id, duplicate_tab, new_sheet_name)
                if not tab_name:
                    continue

            # Determine the range to update
            data_range = f'{tab_name}!{gl_start_cell}:{gl_end_cell}'
            
            body = {
                'values': data_values
            }
            
            result = service.spreadsheets().values().update(
                spreadsheetId=gsheet_id, 
                range=data_range,
                valueInputOption='RAW', 
                body=body).execute()
            
            print(f"{result.get('updatedCells')} cells updated in {gsheet_id}.")

# Main script
def main():
    folder_path = select_folder()
    if not folder_path:
        print("Folder selection cancelled.")
        return

    task_file_path = os.path.join(folder_path, 'Task File.xlsx')
    if not os.path.exists(task_file_path):
        print("Task file does not exist in the selected folder.")
        return
    
    service = authenticate_gsheets()
    task_df, dataframes = load_excel_files(task_file_path, folder_path)
    update_google_sheets(service, task_df, dataframes)

if __name__ == "__main__":
    main()