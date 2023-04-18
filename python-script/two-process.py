from openpyxl import Workbook, load_workbook
import pandas as pd
import os
import time

# Set the path to the directory containing the .xlsb files
xlsb_path = input("Enter the path to the directory containing the .xlsb files: ")

# Loop over the .xlsb files in the directory
for filename in os.listdir(xlsb_path):
    if filename.endswith('.xlsb'):
        # Get user input for week and year
        week = input(f"Enter week value for {filename}: ")
        year = int(input(f"Enter year value for {filename}: "))

        # Convert the .xlsb file to .xlsx format
        xlsb_file_path = os.path.join(xlsb_path, filename)
        xlsx_file_path = os.path.splitext(xlsb_file_path)[0] + '.xlsx'
        print(f"Converting {xlsb_file_path} to {xlsx_file_path}...")
        start_time = time.time()
        try:
            wb = load_workbook(xlsb_file_path)
            with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
                writer.book = wb
                writer.save()
        except Exception as e:
            print(f"Error converting {xlsb_file_path}: {str(e)}")
            continue
        finally:
            if 'wb' in locals():
                wb.close()

        end_time = time.time()
        duration = end_time - start_time
        print(f"File {filename} converted in {duration:.2f} seconds.")

        # Load the .xlsx file into a pandas dataframe
        print(f"Loading file {xlsx_file_path}...")
        start_time = time.time()
        try:
            excel_file = pd.read_excel(xlsx_file_path)

            # Delete the first row
            excel_file = excel_file.iloc[1:]

            # Add week and year columns with user input values
            excel_file.insert(0, 'year', year)
            excel_file.insert(1, 'week', week)

            # Modify the data in the dataframe as needed
            print(f"Modifying data in file {xlsx_file_path}...")
            excel_file = excel_file.replace(to_replace=r'^\d*\.\d*ENULL$', value=' ', regex=True)
            excel_file = excel_file.replace(['#N/A'], ' ')

            # Convert the dataframe to CSV and save to a file with a new name
            new_filename = os.path.splitext(filename)[0] + '_modified.csv'
            csv_file_path = os.path.join(xlsb_path, new_filename)
            print(f"Saving file {csv_file_path}...")
            excel_file.to_csv(csv_file_path, index=False)

        except Exception as e:
            print(f"Error processing {xlsx_file_path}: {str(e)}")

        end_time = time.time()
        duration = end_time - start_time
        print(f"File {filename} processed in {duration:.2f} seconds.")

print("All files processed.")