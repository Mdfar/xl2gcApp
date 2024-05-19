# Excel to Google Sheets Automation

This project provides a tool to automate the process of copying data from Excel spreadsheets to Google Sheets using Python. The application uses a GUI built with Tkinter for user interactions and handles Google Sheets operations with the help of the gspread library.

## Features

- Selects and loads an Excel file containing task specifications.
- Copies specified data ranges from Excel sheets to corresponding Google Sheets.
- Supports the creation of new Google Sheets tabs based on existing templates.
- Provides a GUI for easy interaction and status updates.

## Requirements

- Python 3.x
- Tkinter (comes pre-installed with Python)
- pandas
- gspread
- google-auth
- gspread-dataframe
- openpyxl

## Installation

1. Install the required Python packages:
    ```bash
    pip install -r requirements.txt
    ```

2. Place your Google service account credentials file (`service_account.json`) in the project directory.

## Usage
1. Create a folder that will contain all the Excell files in it.

1. Run the application:
    ```bash
    python main.py
    ```

2. The GUI will prompt you to select the folder containing the Excel files. Then again prompts to select Google service account credentials file if they are not found in the current directory. This will for the first time.

3. Once the setup is complete, click the "Run the Code" button to start the data transfer process.

4. The output and status updates will be displayed in the GUI, including a completion message once the process is finished.

## Detailed Description of Code

### Main Components

#### 1. `gc_creator()`
This function checks for the existence of the service account credentials file and prompts the user to select it if not found. It then authorizes the gspread client.

#### 2. `load_xl_folder()`
This function loads the path of the folder containing Excel files, prompting the user to select the folder if not already saved.

#### 3. `excel_cell_to_indices(cell)`
This function converts an Excel cell reference (e.g., 'B2') to zero-based row and column indices.

#### 4. `main()`
The core function that:
- Loads the task specifications from an Excel file.
- Reads data from specified ranges in the Excel sheets.
- Copies data to the corresponding Google Sheets.
- Handles the creation and deletion of Google Sheets tabs as specified.

#### 5. `run_main()`
Sets up the Tkinter GUI, including a Text widget to display output, a completion label, and buttons to run the code and exit the application.

### Example Usage

```python
if __name__ == "__main__":
    run_main()