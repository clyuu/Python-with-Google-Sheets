import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
from tkinter import ttk, messagebox
import logging
import re

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def authorize_client():
    try:
        # Define scope and credentials
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('sheet.json', scope)
        
        # Authorize the client
        client = gspread.authorize(credentials)
        logger.info('Client authorized successfully.')
        return client
    except Exception as e:
        logger.error(f'An error occurred during authorization: {e}')
        messagebox.showerror("Error", "An error occurred during authorization.")
        return None

def list_spreadsheets(client):
    try:
        spreadsheets = client.openall()
        return spreadsheets
    except Exception as e:
        logger.error(f'An error occurred while listing spreadsheets: {e}')
        messagebox.showerror("Error", "An error occurred while listing spreadsheets.")
        return []

def get_spreadsheet_data(worksheet):
    try:
        data = worksheet.get_all_records()
        logger.info(f'Retrieved data: {data}')
        return data
    except Exception as e:
        logger.error(f'An error occurred while retrieving data: {e}')
        messagebox.showerror("Error", "An error occurred while retrieving data.")
        return []

def append_row_to_worksheet(worksheet, row_to_add):
    try:
        worksheet.append_row(row_to_add)
        logger.info(f'Appended row: {row_to_add}')
    except Exception as e:
        logger.error(f'An error occurred while appending row: {e}')
        messagebox.showerror("Error", "An error occurred while appending row.")

def delete_row_from_worksheet(worksheet, row_id):
    try:
        cell = worksheet.find(str(row_id))
        worksheet.delete_rows(cell.row)
        logger.info(f'Deleted row with ID: {row_id}')
    except Exception as e:
        logger.error(f'An error occurred while deleting row: {e}')
        messagebox.showerror("Error", "An error occurred while deleting row.")

def read_data():
    client = authorize_client()
    if client:
        spreadsheet_title = 'MY SHEET'
        try:
            spreadsheet = client.open(spreadsheet_title)
            worksheet = spreadsheet.sheet1
            data = get_spreadsheet_data(worksheet)
            display_data(data)
        except gspread.exceptions.SpreadsheetNotFound:
            logger.error(f'Spreadsheet with title "{spreadsheet_title}" not found.')
            messagebox.showerror("Error", f'Spreadsheet with title "{spreadsheet_title}" not found.')
        except Exception as e:
            logger.error(f'An error occurred: {e}')
            messagebox.showerror("Error", "An error occurred while accessing the spreadsheet.")

def append_data():
    id = id_entry.get()
    name = name_entry.get()
    age = age_entry.get()

    if id and name and age:
        client = authorize_client()
        if client:
            spreadsheet_title = 'MY SHEET'
            try:
                spreadsheet = client.open(spreadsheet_title)
                worksheet = spreadsheet.sheet1
                existing_data = get_spreadsheet_data(worksheet)

                # Check for duplicate ID
                if any(str(record["ID"]) == id for record in existing_data):
                    messagebox.showwarning("Input Error", "ID already exists. Please enter a unique ID.")
                    return

                row_to_add = [id, name, age]
                append_row_to_worksheet(worksheet, row_to_add)
                messagebox.showinfo("Success", "Data appended successfully.")
            except gspread.exceptions.SpreadsheetNotFound:
                logger.error(f'Spreadsheet with title "{spreadsheet_title}" not found.')
                messagebox.showerror("Error", f'Spreadsheet with title "{spreadsheet_title}" not found.')
            except Exception as e:
                logger.error(f'An error occurred: {e}')
                messagebox.showerror("Error", "An error occurred while accessing the spreadsheet.")
    else:
        messagebox.showwarning("Input Error", "Please enter ID, name, and age.")

def load_data():
    id = id_entry.get()
    if not id:
        messagebox.showwarning("Input Error", "Please enter an ID.")
        return
    
    client = authorize_client()
    if client:
        spreadsheet_title = 'MY SHEET'
        try:
            spreadsheet = client.open(spreadsheet_title)
            worksheet = spreadsheet.sheet1
            data = get_spreadsheet_data(worksheet)
            for record in data:
                if str(record["ID"]) == id:
                    name_entry.delete(0, tk.END)
                    name_entry.insert(0, record["Name"])
                    age_entry.delete(0, tk.END)
                    age_entry.insert(0, record["Age"])
                    return
            messagebox.showinfo("Not Found", "No record found with the entered ID.")
        except gspread.exceptions.SpreadsheetNotFound:
            logger.error(f'Spreadsheet with title "{spreadsheet_title}" not found.')
            messagebox.showerror("Error", f'Spreadsheet with title "{spreadsheet_title}" not found.')
        except Exception as e:
            logger.error(f'An error occurred: {e}')
            messagebox.showerror("Error", "An error occurred while accessing the spreadsheet.")

def delete_data():
    id = id_entry.get()
    if not id:
        messagebox.showwarning("Input Error", "Please enter an ID.")
        return

    client = authorize_client()
    if client:
        spreadsheet_title = 'MY SHEET'
        try:
            spreadsheet = client.open(spreadsheet_title)
            worksheet = spreadsheet.sheet1
            delete_row_from_worksheet(worksheet, id)
            messagebox.showinfo("Success", "Data deleted successfully.")
            name_entry.delete(0, tk.END)
            age_entry.delete(0, tk.END)
        except gspread.exceptions.SpreadsheetNotFound:
            logger.error(f'Spreadsheet with title "{spreadsheet_title}" not found.')
            messagebox.showerror("Error", f'Spreadsheet with title "{spreadsheet_title}" not found.')
        except Exception as e:
            logger.error(f'An error occurred: {e}')
            messagebox.showerror("Error", "An error occurred while accessing the spreadsheet.")

def clear_all_data():
    for item in tree.get_children():
        tree.delete(item)
    id_entry.delete(0, tk.END)
    name_entry.delete(0, tk.END)
    age_entry.delete(0, tk.END)
    messagebox.showinfo("Success", "All data cleared.")

def validate_integer(P):
    if P.isdigit() or P == "":
        return True
    else:
        return False

def validate_name(P):
    if re.match("^[A-Za-z]*$", P):
        return True
    else:
        return False

def validate_age(P):
    if re.match("^[0-9]{0,2}$", P):
        return True
    else:
        return False

# Set up the main application window
root = tk.Tk()
root.title("Google Sheets Interface")
root.configure(bg='#808080')

vcmd_int = (root.register(validate_integer), '%P')
vcmd_name = (root.register(validate_name), '%P')
vcmd_age = (root.register(validate_age), '%P')

# Create a frame for the input fields
input_frame = tk.Frame(root, bg='#808080')
input_frame.grid(row=0, column=0, padx=10, pady=5, sticky=tk.NW)

# Create input fields for ID, name, and age
tk.Label(input_frame, text="ID", bg='#808080', fg='white').grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
id_entry = tk.Entry(input_frame, validate="key", validatecommand=vcmd_int)
id_entry.grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)

tk.Label(input_frame, text="Name", bg='#808080', fg='white').grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
name_entry = tk.Entry(input_frame, validate="key", validatecommand=vcmd_name)
name_entry.grid(row=1, column=1, padx=10, pady=5, sticky=tk.W)

tk.Label(input_frame, text="Age", bg='#808080', fg='white').grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
age_entry = tk.Entry(input_frame, validate="key", validatecommand=vcmd_age)
age_entry.grid(row=2, column=1, padx=10, pady=5, sticky=tk.W)

# Create a frame for the buttons
button_frame = tk.Frame(root, bg='#808080')
button_frame.grid(row=0, column=1, padx=10, pady=5, sticky=tk.NE)

# Update buttons to match the provided image
button_style = {
    'bg': 'white',
    'fg': 'black',
    'font': ('Helvetica', 10),
    'relief': 'raised',
    'bd': 2,
    'padx': 5,
    'pady': 5
}

append_button = tk.Button(button_frame, text="Add", command=append_data, **button_style)
append_button.grid(row=0, column=0, padx=5, pady=5)

clear_button = tk.Button(button_frame, text="Clear", command=clear_all_data, **button_style)
clear_button.grid(row=0, column=1, padx=5, pady=5)

read_button = tk.Button(button_frame, text="Read", command=read_data, **button_style)
read_button.grid(row=1, column=0, padx=5, pady=5)

load_button = tk.Button(button_frame, text="Load", command=load_data, **button_style)
load_button.grid(row=1, column=1, padx=5, pady=5)

delete_button = tk.Button(button_frame, text="Delete", command=delete_data, **button_style)
delete_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

# Create a Treeview widget to display the data
columns = ("ID", "Name", "Age")
tree = ttk.Treeview(root, columns=columns, show="headings")
tree.heading("ID", text="ID")
tree.heading("Name", text="Name")
tree.heading("Age", text="Age")
tree.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

# Add some padding
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(1, weight=1)

# Start the GUI event loop
root.mainloop()
