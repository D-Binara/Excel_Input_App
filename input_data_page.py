import pandas as pd
import tkinter as tk
from tkinter import Entry, Label, Button, messagebox
import sys

# Check if an argument (selected_file) is provided when running the script
if len(sys.argv) != 2:
    messagebox.showerror("Error", "Selected Excel file not provided.")
    sys.exit(1)

# Extract the selected file from the command line argument
selected_file = sys.argv[1]

# Function to create an Excel sheet and add data
def create_excel():
    # Get the data from the input fields
    name = name_entry.get()
    age = age_entry.get()
    address = address_entry.get()
    nic = nic_entry.get()

    # Create or load the Excel file
    excel_filename = selected_file
    writer = pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a')

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

    messagebox.showinfo("Data Added", "Data has been added to the Excel sheet.")

# Create the main application window for the input data page
app = tk.Tk()
app.title("Input Data")

# Create input fields and labels for user data
name_label = Label(app, text="Name:")
name_label.pack()
name_entry = Entry(app)
name_entry.pack()

age_label = Label(app, text="Age:")
age_label.pack()
age_entry = Entry(app)
age_entry.pack()

address_label = Label(app, text="Address:")
address_label.pack()
address_entry = Entry(app)
address_entry.pack()

nic_label = Label(app, text="NIC Number:")
nic_label.pack()
nic_entry = Entry(app)
nic_entry.pack()

# Create button for adding data to the Excel sheet
add_button = Button(app, text="Add Data to Excel Sheet", command=create_excel)
add_button.pack(pady=10)

app.mainloop()
