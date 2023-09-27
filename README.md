# Excel_Input_App
This Git repository contains a Python application for creating and managing Excel sheets through a graphical user interface (GUI). The application is split into multiple pages, each with specific functionality.

## Pages:

### Choose Action (create_or_open.py):
This page allows the user to choose between creating a new Excel file or working with an existing one.
Provides options to "Create New Excel File" and "Open Existing Excel File."

### Create New File (create_new_file_page.py):
Accessed when the user selects "Create New Excel File" on the first page.
Allows the user to input Name, Age, Address, NIC Number, and specify an Excel file name.
Creates a new Excel file with the provided data when the "Create Excel Sheet" button is clicked.

### Open Existing File (open_existing_file_page.py):
Accessed when the user selects "Open Existing Excel File" on the first page.
Lists existing Excel files in the current directory when the "List Existing Files" button is clicked.

## Dependencies:
Python 3.x
pandas library
tkinter library (for GUI)
openpyxl or xlsxwriter (for Excel file handling)

## Usage:
Run create_or_open.py to launch the application.
Choose to either create a new Excel file or work with an existing one.
Follow the prompts and input data as required.
For creating new files, click the "Create Excel Sheet" button.
For listing existing files, click the "List Existing Files" button.

## Contributions:
Contributions and improvements to the code are welcome. Feel free to fork the repository, make changes, and submit pull requests.
