import os
import shutil
import pandas as pd

# Read the Excel sheet and extract the file paths
excel_file = r"C:\Users\Aslam\Desktop\file_paths.xlsx"
df = pd.read_excel(excel_file)
column_name = 'path'  # Replace with the correct column name from your Excel file
file_paths = df[column_name].tolist()

# Source and destination folders
source_folder = r"C:\Users\Aslam\Desktop\FORM 16A Q1 FY 2022-23"
destination_folder = r"C:\Users\Aslam\Desktop\TDS CRTIFICATE"

# Iterate over the file paths and copy the files
for file_path in file_paths:
    # Extract the file name from the file path
    file_name = os.path.basename(file_path)
    
    # Construct the source and destination paths
    source_path = os.path.join(source_folder, file_name)
    destination_path = os.path.join(destination_folder, file_name)
    
    # Copy the file from the source to the destination folder
    shutil.copyfile(source_path, destination_path)
    print(f"File '{file_name}' copied successfully!")
