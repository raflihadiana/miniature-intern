from openpyxl import load_workbook
import pandas as pd
import os

# Set the path to the directory containing the .xlsb files
xlsb_path = '/home/ubuntu/binary'

# Loop over the .xlsb files in the directory
for filename in os.listdir(xlsb_path):
    if filename.endswith('.xlsb'):
        # Load the .xlsb file into a pandas dataframe
        file_path = os.path.join(xlsb_path, filename)
        print(f"Loading file {file_path}...")
        with load_workbook(file_path, read_only=True) as wb:
            sheet = wb.active
            rows_generator = sheet.rows
            headers = [cell.value for cell in next(rows_generator)]
            excel_file = pd.DataFrame([[cell.value for cell in row] for row in rows_generator], columns=headers)

            # Modify the data in the dataframe as needed
            print(f"Modifying data in file {file_path}...")
            excel_file = excel_file.replace(['ENULL', '#N/A'], ' ')

            # Convert the dataframe to CSV and save to a file with a new name
            new_filename = filename[:-5] + '_modified.csv'
            csv_file_path = os.path.join(xlsb_path, new_filename)
            print(f"Saving file {csv_file_path}...")
            excel_file.to_csv(csv_file_path, index=False)

print("All files processed.")
