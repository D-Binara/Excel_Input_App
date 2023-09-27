import pandas as pd
import tkinter as tk
from tkinter import messagebox

def create_excel():
    # Get the data from the input fields
    name = name_entry.get()
    age = age_entry.get()
    address = address_entry.get()
    nic = nic_entry.get()

    # Create or load the Excel file
    excel_filename = file_name_entry.get() + ".xlsx"
    writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

    # Create a pandas DataFrame with the user's data
    data = {
        'Name': [name],
        'Age': [age],
        'Address': [address],
        'NIC Number': [nic]
    }

    df = pd.DataFrame(data)

    # Write the DataFrame to the Excel sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Save the Excel file
    writer.close()

    messagebox.showinfo("Excel Sheet Created", f"The Excel sheet '{excel_filename}' has been created successfully.")

# Create the main application window for Page 2
app = tk.Tk()
app.title("Create New Excel File")

# Create input fields and labels for user data
name_label = tk.Label(app, text="Name:")
name_label.pack()
name_entry = tk.Entry(app)
name_entry.pack()

age_label = tk.Label(app, text="Age:")
age_label.pack()
age_entry = tk.Entry(app)
age_entry.pack()

address_label = tk.Label(app, text="Address:")
address_label.pack()
address_entry = tk.Entry(app)
address_entry.pack()

nic_label = tk.Label(app, text="NIC Number:")
nic_label.pack()
nic_entry = tk.Entry(app)
nic_entry.pack()

# Create input field and label for Excel file name
file_name_label = tk.Label(app, text="Excel File Name:")
file_name_label.pack()
file_name_entry = tk.Entry(app)
file_name_entry.pack()

# Create button for creating the Excel file
create_button = tk.Button(app, text="Create Excel Sheet", command=create_excel)
create_button.pack(pady=10)

app.mainloop()
