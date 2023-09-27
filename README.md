# Excel_Input_App
This Git repository contains a Python application for creating and managing Excel sheets through a graphical user interface (GUI). The application is split into multiple pages, each with specific functionality.


## Pages:

1. [create_or_open.py](create_or_open.py): The entry point of the application, allowing the user to choose between creating a new Excel file or working with an existing one.

2. [create_new_file_page.py](create_new_file_page.py): The page for creating a new Excel file. It provides input fields for entering data (Name, Age, Address, NIC Number) and specifying an Excel file name. Data entered here is written to the new Excel sheet.

3. [open_existing_file_page.py](open_existing_file_page.py): The page for working with an existing Excel file. It lists all existing Excel files in the current directory, and the user can select one to add new data to.

4. [input_data_page.py](input_data_page.py): The page for entering data to add to an existing Excel file. This page is accessed from the "Open Existing Excel File" page and allows the user to input Name, Age, Address, and NIC Number. Data entered here is appended to the selected Excel file.


## Dependencies:
Python 3.x
pandas library
tkinter library (for GUI)
openpyxl or xlsxwriter (for Excel file handling)

## Usage:
1. Run `create_or_open.py` to start the application.
2. Choose between creating a new Excel file or working with an existing one.
3. If you choose to create a new file:
   - Enter data for Name, Age, Address, NIC Number, and specify an Excel file name.
   - Click "Create Excel Sheet" to create the Excel file.
4. If you choose to work with an existing file:
   - Select an Excel file from the list of existing files.
   - Click "Open Selected Excel File" to add new data to the selected file.


## Contributions:
Contributions and improvements to the code are welcome. Feel free to fork the repository, make changes, and submit pull requests.
