import pandas as pd
import subprocess
import platform
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to let user pick an Excel file
def select_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        title="Select an Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    return file_path


# Function to ping a host
def ping_host(host):
    param = "-n" if platform.system().lower() == "windows" else "-c"
    command = ["ping", param, "1", host]

# Function to ping a host
def ping_host(host):
    param = "-n" if platform.system().lower() == "windows" else "-c"
    command = ["ping", param, "1", host]
    
    try:
        output = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        return "Reachable" if output.returncode == 0 else "Unreachable"
    except Exception as e:
        return f"Error: {e}"
    
# Let user choose the file
file_path = select_file()

if not file_path:
    messagebox.showerror("Error", "No file selected. Exiting.")
    exit()

# Load the selected Excel file
df = pd.read_excel(file_path)

# Check if 'Host' column exists
if "Host" not in df.columns:
    messagebox.showerror("Error", "Excel file must contain a 'Host' column.")
    exit()

# Ping each host and save results
df["Status"] = df["Host"].apply(ping_host)

# Save results to a new file
output_file = file_path.replace(".xlsx", "_ping_results.xlsx")
df.to_excel(output_file, index=False)

messagebox.showinfo("Success", f"Results saved to {output_file}")