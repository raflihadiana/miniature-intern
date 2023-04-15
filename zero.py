from openpyxl import load_workbook
import pandas as pd
import os

# Set the path to the directory containing the .xlsx files
xlsx_path = 'C:\\Users\\rafli\\XL\\Database RAN Capacity\\test'

# Loop over the .xlsx files in the directory
for filename in os.listdir(xlsx_path):
    if filename.endswith('.xlsx'):
        # Load the .xlsx file into a pandas dataframe
        file_path = os.path.join(xlsx_path, filename)
        print(f"Loading file {file_path}...")
        wb = load_workbook(file_path)
        sheet = wb.active
        data = []
        for row in sheet.rows:
            data.append([cell.value for cell in row])
        excel_file = pd.DataFrame(data[1:], columns=data[0])
        
        # Modify the data in the dataframe as needed
        print(f"Modifying data in file {file_path}...")
        excel_file = excel_file.replace(['ENULL', '#N/A'], ' ')
        
        # Convert the dataframe to CSV and save to a file with a new name
        new_filename = filename[:-5] + '_modified.csv'
        csv_file_path = os.path.join(xlsx_path, new_filename)
        print(f"Saving file {csv_file_path}...")
        excel_file.to_csv(csv_file_path, index=False)
        
print("All files processed.")
