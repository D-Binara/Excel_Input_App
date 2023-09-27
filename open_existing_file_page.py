import os
import tkinter as tk
from tkinter import messagebox, Listbox, Scrollbar, Button

# Function to switch to the input data page with the selected file
def open_input_data_page():
    selected_index = listbox.curselection()
    if selected_index:
        selected_file = listbox.get(selected_index[0])
        app.destroy()  # Close the current page
        os.system(f'python input_data_page.py "{selected_file}"')  # Launch the input data page with the selected file
    else:
        messagebox.showinfo("No File Selected", "Please select an Excel file from the list.")

def populate_listbox():
    existing_files = [filename for filename in os.listdir() if filename.endswith('.xlsx')]
    if existing_files:
        for filename in existing_files:
            listbox.insert(tk.END, filename)
    else:
        listbox.insert(tk.END, "No existing Excel files found in the current directory.")

# Create the main application window for Page 3 (Open Existing File)
app = tk.Tk()
app.title("Open Existing Excel File")

# Create a listbox to display existing Excel files
listbox = Listbox(app, selectmode=tk.SINGLE, width=40, height=10)
listbox.pack(pady=10)

# Add scrollbar to the listbox
scrollbar = Scrollbar(app, orient=tk.VERTICAL)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=listbox.yview)

# Create a button to open the selected Excel file
open_button = Button(app, text="Open Selected Excel File", command=open_input_data_page)
open_button.pack(pady=10)

# Populate the listbox with existing Excel files
populate_listbox()

app.mainloop()
