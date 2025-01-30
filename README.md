# Ping Script

This is a Python script that reads a list of hosts from an Excel file and pings each host to check if they are reachable. It provides the status of each host, including whether the DNS resolved successfully and whether the host is reachable.

## Features

- Read host addresses from an Excel file.
- Ping each host and check its availability.
- Display results for DNS resolution and host reachability.
- User-friendly GUI to select the Excel file.

## Requirements

- Python 3.x
- Required Python packages:
  - pandas
  - openpyxl
  - socket
  - tkinter (for GUI)

## Usage
- Place your Excel file with a list of hosts in a folder.
- Run the script and select the Excel file using the GUI.
- The script will attempt to ping each host and show the results in the console.
