import tkinter as tk


def create_new_file():
    # Close the current page
    root.destroy()

    # Navigate to the page for creating a new file
    import create_new_file_page


def open_existing_file():
    # Close the current page
    root.destroy()

    # Navigate to the page for working with an existing file
    import open_existing_file_page


# Create the main application window for Page 1
root = tk.Tk()
root.title("Excel Sheet Creator")

# Add buttons for choosing an action
create_button = tk.Button(root, text="Create New Excel File", command=create_new_file)
create_button.pack(pady=20)

open_button = tk.Button(root, text="Open Existing Excel File", command=open_existing_file)
open_button.pack(pady=20)

root.mainloop()
