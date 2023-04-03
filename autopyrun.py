import os
import subprocess
from pathlib import Path
from datetime import datetime, timedelta

# Change the working directory
os.chdir("C:\\Users\\KISS Admin\\Desktop\\IVYENT_DH\\P4. stock check automation code")

# Run get_stockchecklist_IVY.py
subprocess.run(["python", "get_stockchecklist_IVY.py"], check=True)

# Run dailyDM_simulator.py
subprocess.run(["python", "dailyDM_simulator.py"], check=True)

# Open the folder in File Explorer
folder_path = "C:\\Users\\KISS Admin\\Desktop\\stock check practice"
subprocess.run(["explorer", folder_path])

# Define time threshold: 10 minutes ago
time_threshold = datetime.now() - timedelta(minutes=10)

# Find and open recent Excel files created within the last 10 minutes
for entry in os.scandir(folder_path):
    if entry.is_file() and entry.name.endswith(".xlsx"):
        file_creation_time = datetime.fromtimestamp(entry.stat().st_ctime)
        if file_creation_time > time_threshold:
            subprocess.run(["start", "excel.exe", entry.path], shell=True)
