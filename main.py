import pandas as pd
import subprocess
import platform
import socket
import tkinter as tk
from tkinter import filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor

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
def ping_host(row):
    # Find the first column that contains "IP" in its name (case-insensitive)
    ip_column = next((col.strip() for col in row.keys() if "ip" in col.lower().strip()), None)
    host = str(row["Host"]).strip()
    ip_from_file = str(row[ip_column]).strip() if ip_column and row[ip_column] else None  # Get the IP from detected column
    
    try:
        # Step 1: Try to resolve DNS
        ip_address = socket.gethostbyname(host)
    except socket.gaierror:
        # Step 2: If DNS resolution fails, try using the IP from the file
        if pd.notna(ip_from_file):
            ip_address = ip_from_file
        else: 
            return host, "DNS Failure and no IP available"  # Hostname does not resolve
    
    # Step 2: If DNS resolves, attempt to ping
    param = "-n" if platform.system().lower() == "windows" else "-c"
    command = ["ping", param, "1", host]
    
    try:
        output = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        # If the host is reachable
        if output.returncode == 0:
            return host, "Reachable"
        else:
            # Step 2: Host unreachable, try to ping the IP address
            command = ["ping", param, "1", ip_address]
            output = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            # Return result for IP ping
            return host, "Only IP reachable" if output.returncode == 0 else "Unreachable"
        
    except Exception as e:

        return host, ip_address, f"Error: {e}"
    
# Let user choose the file
file_path = select_file()

# If no file was selected, exit the program
if not file_path:
    messagebox.showerror("Error", "No file selected. Exiting.")
    exit()

# Load the selected Excel file
df = pd.read_excel(file_path)

# Check if 'Host' column exists
if "Host" not in df.columns:
    messagebox.showerror("Error", "Excel file must contain a 'Host' column.")
    exit()

df = df[df["Host"].notna()]  # Removes NaN values

# Multi-threaded execution of ping_host function
with ThreadPoolExecutor(max_workers=200) as executor:  # Adjust the number of threads as needed
    results = list(executor.map(ping_host, df.to_dict(orient="records")))

print (results)

for result in results:
    print(result)

# Store results in DataFrame
df["Status"] = [status for _, status in results]

# Save results to a new file
output_file = file_path.replace(".xlsx", "_ping_results.xlsx")
df.to_excel(output_file, index=False)

messagebox.showinfo("Success", f"Results saved to {output_file}")