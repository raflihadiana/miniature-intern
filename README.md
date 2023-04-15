# miniature-intern

## Apa Tujuan dari Project Ini?

Project ini diadakan untuk melakukan query data cell capacity 4G XL Axiata. Hal ini digunakan untuk mempermudah administrasi data. Namun ada beberapa hal yang perlu diperhatikan:

1. Membuat python script
2. Membuat sql script

### Python Script

Script python dibuat guna untuk melakukan otomatisasi dalam memanipulasi banyak data sekaligus. Yakni mengubah isi dari cell excel binary dan menkonversikan hasil akhir tersebut kedalam format csv. Berikut contohnya:

```python
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
```

### SQL Script

SQL script digunakan untuk melakukan pembuatan table pada data Capacity 4G Cell. Hal ini dikarenakan inputan manual dengan data 300.000 per-minggu hanya akan memakan banyak waktu, maka dari itu diperlukan sql script untuk membuat table dengan nama sesuai format.


Selain itu bagaimana cara melakukan load file (di sini baru satu file saja), dimana akan di update tentang bagaimana cara melakukan bulk load file, me-load data dalam banyak file sekaligus.

```sql
LOAD DATA INFILE 'C:\Users\rafli\XL\Database RAN Capacity\CSV\W01 - Cell 4G NQI Capacity 30000 Data.csv'
INTO TABLE cell_capacity 
FIELDS TERMINATED BY ','
ENCLOSED BY '"'
LINES TERMINATED BY '\n'
IGNORE 1 ROWS;
```