# Import modules
import os
import glob
import pandas as pd
import xlsxwriter

# Get a list of file paths that match '*.xlsx' pattern
file_paths = glob.glob(r'C:\Users\Aslam\Desktop\Reports for Business Central_Aslam\Vishambhar\*.pdf')

# Create a new excel workbook and a worksheet
workbook = xlsxwriter.Workbook(os.path.join(os.path.dirname(os.path.abspath(__file__)),"file_list.xlsx"))
worksheet = workbook.add_worksheet()

# Set the row and column indices
row = 0
col = 0

# Loop through the file paths and write the file names and folder names to the worksheet
for file_path in file_paths:
    # Extract the file name and folder name
    file_name = os.path.basename(file_path)
    folder_name = os.path.dirname(file_path)

    # Write the file name and folder name to the worksheet
    worksheet.write(row, col, file_name)
    worksheet.write(row, col + 1, folder_name)

    # Increment the row index
    row += 1

# Save and close the workbook
workbook.close()
