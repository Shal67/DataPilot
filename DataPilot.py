import tkinter as tk
from tkinter import filedialog, ttk
from datetime import datetime
import pandas as pd
 
# Functions for the first Excel file (expired.py)
def load_file_1():
    """Function to load the first Excel file."""
    global df1, filtered_devices, domains, clients
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        try:
            df1 = pd.read_excel(file_path)
            file_path_label_1.config(text=f"File 1 loaded: {file_path}")
 
            if "DOMAIN" in df1.columns and "Client" in df1.columns:
                domains = sorted(df1["DOMAIN"].dropna().unique())
                domains.insert(0, "All Domains")  # Add "All Domains" at the beginning
 
                domain_dropdown["values"] = domains
                domain_dropdown.set("All Domains")  # Default to "All Domains"
 
            # Consider all clients by default
            filtered_devices = df1[["Device Name", "Serial number"]].dropna().values.flatten()
            update_client_list()  # Update client list immediately
            update_device_suggestions()  # Enable suggestions immediately
 
        except Exception as e:
            file_path_label_1.config(text=f"Failed to load file: {e}")
 
def update_client_list(*args):
    """Update the client names based on the selected domain or show all clients."""
    global clients
    if df1 is None:
        return
 
    selected_domain = domain_dropdown.get()
    if selected_domain == "All Domains":
        clients = sorted(df1["Client"].dropna().unique())
    else:
        clients = sorted(df1[df1["DOMAIN"] == selected_domain]["Client"].dropna().unique())
 
    clients.insert(0, "All Clients")  # Add "All Clients" at the beginning
    client_dropdown["values"] = clients
    client_dropdown.set("All Clients")  # Default to "All Clients"
 
    update_device_list()
 
def update_device_list(*args):
    """Update the device names and serial numbers based on the selected client or show all devices."""
    global filtered_devices
    if df1 is None:
        return
 
    # Clear the device entry and output box
    device_entry.delete(0, tk.END)
    output_text.delete("1.0", tk.END)
 
    selected_client = client_dropdown.get()
    if selected_client == "All Clients":
        filtered_devices = df1[["Device Name", "Serial number"]].dropna().values.flatten()
    else:
        filtered_devices = df1[df1["Client"] == selected_client][["Device Name", "Serial number"]].dropna().values.flatten()
 
    update_device_suggestions()
    update_expired_count()
    display_sheet_data(selected_client)  # Display data for the selected client
 
def update_device_suggestions(*args):
    """Update device suggestions based on the search input."""
    if df1 is None:
        return
 
    search_value = device_entry.get().strip().lower()
    suggestion_list.delete(0, tk.END)
 
    if not search_value:
        return
 
    matching_devices = [device for device in filtered_devices if search_value in str(device).lower()]
    for device in matching_devices:
        suggestion_list.insert(tk.END, device)
 
def fill_device_entry(event):
    """Fill the device entry with the selected suggestion."""
    selected_device = suggestion_list.get(tk.ANCHOR)
    device_entry.delete(0, tk.END)
    device_entry.insert(0, selected_device)
    suggestion_list.delete(0, tk.END)
 
def fetch_row():
    """Fetch row data for the selected device or serial number."""
    if df1 is None:
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, "No file loaded.")
        return
 
    selected_input = device_entry.get().strip().lower()
    if not selected_input:
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, "Please enter a Device Name or Serial Number.")
        return
 
    # Clean up column names to remove leading/trailing spaces
    df1.columns = df1.columns.str.strip()
 
    # Check if necessary columns are present
    if "Device Name" not in df1.columns or "Serial number" not in df1.columns:
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, "No column named 'Device Name' or 'Serial number' found.")
        return
 
    df_normalized = df1.copy()
    df_normalized["Device Name"] = df_normalized["Device Name"].astype(str).str.lower()
    df_normalized["Serial number"] = df_normalized["Serial number"].astype(str).str.lower()
 
    # Search for either device name or serial number
    row = df1[(df_normalized["Device Name"] == selected_input) | (df_normalized["Serial number"] == selected_input)]
 
    if not row.empty:
        output_text.delete("1.0", tk.END)
        formatted_output = "\n".join([f"{col}: {row.iloc[0][col]}" for col in row.columns])
        output_text.insert(tk.END, formatted_output)
    else:
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, "No matching device name or serial number found.")
 
def update_expired_count():
    """Update the count of expired devices based on the selected client."""
    if df1 is None or "Maintenance Status" not in df1.columns:
        expired_box.delete("1.0", tk.END)
        expired_box.insert(tk.END, "0")
        return
 
    selected_client = client_dropdown.get()
    if selected_client == "All Clients":
        expired_count = df1[df1["Maintenance Status"].str.lower() == "expired"].shape[0]
    else:
        expired_count = df1[(df1["Client"] == selected_client) & (df1["Maintenance Status"].str.lower() == "expired")].shape[0]
 
    expired_box.delete("1.0", tk.END)
    expired_box.insert(tk.END, str(expired_count))
 
def export_to_excel():
    """Export the row data to an Excel file."""
    if df1 is None:
        return
 
    selected_input = device_entry.get().strip().lower()
    if not selected_input:
        return
 
    df_normalized = df1.copy()
    df_normalized["Device Name"] = df_normalized["Device Name"].astype(str).str.lower()
    df_normalized["Serial number"] = df_normalized["Serial number"].astype(str).str.lower()
 
    row = df1[(df_normalized["Device Name"] == selected_input) | (df_normalized["Serial number"] == selected_input)]
 
    if not row.empty:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            row.to_excel(file_path, index=False, engine="openpyxl")
 
# Functions for the second Excel file (capacity.py)
def load_file_2():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if file_path:
        file_path_label_2.config(text=f"File 2 loaded: {file_path}")
        process_excel(file_path)
 
def process_excel(file_path):
    global excel_data, current_month_index, months
    excel_data = pd.read_excel(file_path, sheet_name=None)
    sheet_list = list(excel_data.keys())
    if "All Clients" not in sheet_list:
        sheet_list.insert(0, "All Clients")  # Add "All Clients" at the beginning if not already present
    client_dropdown["values"] = sheet_list
    client_dropdown.set("All Clients")  # Default to "All Clients"
   
    # Extract month names from the first column of the first sheet
    first_sheet = list(excel_data.values())[0]
    months = list(first_sheet.iloc[:, 0].astype(str))
   
    current_month_name = datetime.now().strftime("%B")
    if current_month_name in months:
        current_month_index = months.index(current_month_name)
    else:
        current_month_index = 0  # Default to first month if not found
 
    # Display All Clients sheet data by default
    display_sheet_data("All Clients")
 
def display_sheet_data(sheet_name):
    # Clear the device entry and output box
    device_entry.delete(0, tk.END)
    output_text.delete("1.0", tk.END)
   
    output_box_left.delete("1.0", tk.END)
    output_box_right.delete("1.0", tk.END)
 
    if sheet_name == "Select a sheet":
        return
 
    if sheet_name not in excel_data:
        output_box_right.insert(tk.END, f"Sheet '{sheet_name}' not found.\n")
        return
 
    if sheet_name == "All Clients":
        sheet_data = excel_data[sheet_name]
        display_data(sheet_data, output_box_right)
    else:
        summary_data = excel_data["All Clients"]
        model_data = excel_data[sheet_name]
 
        # Display All Clients sheet data in the left output box
        display_data(summary_data, output_box_left)
 
        # Display model data in the right output box
        display_data(model_data, output_box_right)
 
def display_data(sheet_data, output_box):
    sheet_df = pd.DataFrame(sheet_data)
 
    # Set proper headers from the first row, and ensure the first row remains as part of the data
    sheet_df.columns = sheet_df.iloc[0].astype(str)
    sheet_df = sheet_df.drop(0)  # Drop the first row after it has been set as the header
 
    if 0 <= current_month_index < len(sheet_df):
        month_row = sheet_df.iloc[current_month_index]
        for col, value in zip(sheet_df.columns, month_row.values):
            output_box.insert(tk.END, f"{col}: {value}\n")
 
def show_previous_month():
    global current_month_index
    if current_month_index > 0:
        current_month_index -= 1
        display_sheet_data(client_dropdown.get())
 
def show_next_month():
    global current_month_index
    if current_month_index < len(months) - 1:
        current_month_index += 1
        display_sheet_data(client_dropdown.get())
  
# GUI Setup
root = tk.Tk()
root.title("Excel Data Viewer")
 
frame = tk.Frame(root, bg="#F5F5F5", relief="solid", borderwidth=1)
frame.pack(pady=15, padx=15, fill="both", expand=True)

 
# File Load Buttons
load_button_1 = tk.Button(frame, text="Load Excel File 1", command=load_file_1, bg="#4CAF50", fg="white", font=("Helvetica", 10, "bold"), width=18)
load_button_1.grid(row=0, column=0, padx=5, pady=5)

load_button_2 = tk.Button(frame, text="Load Excel File 2", command=load_file_2, bg="#2196F3", fg="white", font=("Helvetica", 10, "bold"), width=18)
load_button_2.grid(row=0, column=1, padx=5, pady=5)

 
# Labels to show the file paths of the loaded files
file_path_label_1 = tk.Label(frame, text="No file 1 loaded", anchor="w", width=50, font=("Helvetica", 10), bg="#FFFFFF", relief="solid", borderwidth=1)
file_path_label_1.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

file_path_label_2 = tk.Label(frame, text="No file 2 loaded", anchor="w", width=50, font=("Helvetica", 10), bg="#FFFFFF", relief="solid", borderwidth=1)
file_path_label_2.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

 
# Buttons for navigating months
prev_month_button = tk.Button(
    frame,
    text="Previous Month",
    command=show_previous_month,
    width=20,
    font=("Helvetica", 10),
    bg="#4CAF50",
    fg="white",
    relief="solid",
    borderwidth=1
)
prev_month_button.grid(row=3, column=1, padx=5, pady=5)

next_month_button = tk.Button(
    frame,
    text="Next Month",
    command=show_next_month,
    width=20,
    font=("Helvetica", 10),
    bg="#4CAF50",
    fg="white",
    relief="solid",
    borderwidth=1
)
next_month_button.grid(row=3, column=2, padx=5, pady=5)

# Output Box for Data Display
output_frame = tk.Frame(root)
output_frame.pack(pady=10, padx=10)

output_box_left = tk.Text(
    output_frame,
    height=8,
    width=39,
    wrap=tk.NONE,
    bg="lightyellow",
    font=("Helvetica", 10),
    relief="solid",
    borderwidth=1
)
output_box_left.grid(row=0, column=0, padx=5, pady=5)

output_box_right = tk.Text(
    output_frame,
    height=8,
    width=39,
    wrap=tk.NONE,
    bg="lightblue",
    font=("Helvetica", 10),
    relief="solid",
    borderwidth=1
)
output_box_right.grid(row=0, column=1, padx=2, pady=2)

 
# UI Components from expired.py# Domain label and dropdown
domain_label = tk.Label(
    frame,
    text="Domain:",
    bg="#D0E9FF",
    font=("Helvetica", 9),
    relief="solid",
    borderwidth=1
)
domain_label.grid(row=4, column=0, padx=3, pady=3, sticky="w")

domain_dropdown = ttk.Combobox(
    frame,
    state="readonly",
    font=("Helvetica", 9),
    width=20
)
domain_dropdown.grid(row=4, column=1, padx=3, pady=3)
domain_dropdown.bind("<<ComboboxSelected>>", update_client_list)

# Client label and dropdown
client_label = tk.Label(
    frame,
    text="Client:",
    bg="#D0E9FF",
    font=("Helvetica", 9),
    relief="solid",
    borderwidth=1
)
client_label.grid(row=5, column=0, padx=3, pady=3, sticky="w")

client_dropdown = ttk.Combobox(
    frame,
    state="readonly",
    font=("Helvetica", 9),
    width=20
)
client_dropdown.grid(row=5, column=1, padx=3, pady=3)
client_dropdown.bind("<<ComboboxSelected>>", update_device_list)

# Device label, entry, suggestion list
device_label = tk.Label(
    frame,
    text="Device:",
    bg="#D0E9FF",
    font=("Helvetica", 9),
    relief="solid",
    borderwidth=1
)
device_label.grid(row=6, column=0, padx=3, pady=3, sticky="w")

device_entry = tk.Entry(
    frame,
    font=("Helvetica", 9),
    width=20
)
device_entry.grid(row=6, column=1, padx=3, pady=3)
device_entry.bind("<KeyRelease>", update_device_suggestions)

suggestion_list = tk.Listbox(
    frame,
    height=4,
    font=("Helvetica", 9),
    width=25
)
suggestion_list.grid(row=6, column=2, padx=3, pady=3)
suggestion_list.bind("<Double-Button-1>", fill_device_entry)

# Fetch and Export buttons
fetch_button = tk.Button(
    frame,
    text="Fetch",
    command=fetch_row,
    bg="#4CAF50",
    fg="white",
    font=("Helvetica", 9),
    height=1,
    width=10,
    relief="solid",
    borderwidth=1
)
fetch_button.grid(row=7, column=0, padx=3, pady=3)

export_button = tk.Button(
    frame,
    text="Export to Excel",
    command=export_to_excel,
    bg="#FF9800",
    fg="white",
    font=("Helvetica", 9),
    height=1,
    width=15,
    relief="solid",
    borderwidth=1
)
export_button.grid(row=7, column=1, padx=3, pady=3)

# Expired label and box
expired_label = tk.Label(
    frame,
    text="Expired:",
    bg="#D0E9FF",
    font=("Helvetica", 9),
    relief="solid",
    borderwidth=1
)
expired_label.grid(row=7, column=2, padx=3, pady=3, sticky="w")

expired_box = tk.Text(
    frame,
    height=1,
    width=8,
    font=("Helvetica", 9),
    bg="#FFE0E0",
    fg="black",
    relief="solid",
    borderwidth=1
)
expired_box.grid(row=7, column=3, padx=3, pady=3)

# Output text box
output_text = tk.Text(
    root,
    height=15,
    width=80,
    font=("Helvetica", 9),
    bg="#E0F7FA",
    relief="solid",
    borderwidth=1
)
output_text.pack(pady=2, padx=3)

 
# Global variables
df1 = None
filtered_devices = []
domains = []
clients = []
excel_data = {}
months = []
current_month_index = 0

root.mainloop()