from openpyxl import load_workbook
import pandas as pd
import os
import time
import mysql.connector

# Set the path to the directory containing the .xlsx files
xlsx_path = input("Enter the path to the directory containing the .xlsx files: ")

# Get user input for MySQL database details
host = input("Enter the MySQL host: ")
port = input("Enter the MySQL port: ")
database = input("Enter the MySQL database name: ")
user = input("Enter the MySQL user: ")
password = input("Enter the MySQL password: ")

# Create a connection to the MySQL database
cnx = mysql.connector.connect(host=host, port=port, database=database, user=user, password=password)
cursor = cnx.cursor()

# Loop over the .xlsx files in the directory
for filename in os.listdir(xlsx_path):
    if filename.endswith('.xlsx'):
        # Get user input for week and year
        week = input(f"Enter week value for {filename}: ")
        year = int(input(f"Enter year value for {filename}: "))

        # Load the .xlsx file into a pandas dataframe
        file_path = os.path.join(xlsx_path, filename)
        print(f"Loading file {file_path}...")
        start_time = time.time()
        try:
            wb = load_workbook(file_path, read_only=True)
            sheet = wb.active
            data = []
            for row in sheet.rows:
                data.append([cell.value for cell in row])
            excel_file = pd.DataFrame(data[1:], columns=data[0])

            # Delete the first row
            excel_file = excel_file.iloc[1:]

            # Add week and year columns with user input values
            excel_file.insert(0, 'year', year)
            excel_file.insert(1, 'week', week)

            # Modify the data in the dataframe as needed
            print(f"Modifying data in file {file_path}...")
            excel_file = excel_file.replace(to_replace=r'^\d*\.\d*ENULL$', value=' ', regex=True)
            excel_file = excel_file.replace(['#N/A'], ' ')

            # Convert the dataframe to CSV and save to a file with a new name
            new_filename = filename[:-5] + '_modified.csv'
            csv_file_path = os.path.join(xlsx_path, new_filename)
            print(f"Saving file {csv_file_path}...")
            excel_file.to_csv(csv_file_path, index=False)

            # Read the CSV file and create a DataFrame object
            df = pd.read_csv(csv_file_path)

            # Insert the data into the MySQL database
            table_name = filename[:-5] + '_table'
            df.to_sql(name=table_name, con=cnx, if_exists='replace', index=False)

        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")

        finally:
            if 'wb' in locals():
                wb.close()

        end_time = time.time()
        duration = end_time - start_time
        print(f"File {filename} processed in {duration:.2f} seconds.")

print("All files processed.")

# Close the connection to the MySQL database
cursor.close()
cnx.close()
