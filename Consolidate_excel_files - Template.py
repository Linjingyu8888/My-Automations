from pathlib import Path
from openpyxl import load_workbook
import xlrd
import csv
import traceback
import datetime

# Get the current directory
source_directory = Path(r"C:\Users\[Your path]")
target_csv_file = "001_consolidation.csv"
files_to_ignore = {target_csv_file, "[Your file to ignore]"}

# number of rows to skip if the excel files do not have headers on the first row
number_of_rows_to_skip = 2



def main(source_directory, target_csv_file, files_to_ignore):
    # List all files in the source_directory 
    all_file_names = [f for f in source_directory.iterdir() if f.is_file() and f.name not in files_to_ignore]

    # Create a list of all the rows
    all_rows = []

    # Create Headers
    headers_row = []


    print("Files in the source_directory:")
    for file_name in all_file_names:
        print(file_name.name)

        if not (file_name.name.endswith('.xls') or file_name.name.endswith('.xlsx')):
            print("skipping file, because only xls or xlsx files are supported")
            continue

        if file_name.name.endswith('.xls'):
            # Open the workbook
            workbook = xlrd.open_workbook(file_name)

            # Access the first sheet by index
            first_sheet = workbook.sheet_by_index(0)


            # Iterate over rows in the first sheet
            for row_index in range(first_sheet.nrows):
                if row_index < number_of_rows_to_skip:
                    continue

                row = first_sheet.row_values(row_index)
            
                if row_index == number_of_rows_to_skip:
                    headers_row = ["file_name"] + row
                else:
                    all_rows.append([file_name.name] + row)
        if file_name.name.endswith('.xlsx'):
            # Open the workbook
            workbook = load_workbook(file_name)
            first_sheet = workbook.worksheets[0]
        
            # Iterate over rows in the first sheet
            row_index = 0
            for row in first_sheet.rows:
                if row_index < number_of_rows_to_skip:
                    continue
                values = []
                for cell in row:
                    value = cell.value
                    if isinstance(value, datetime.datetime):
                        value = value.date().isoformat()
                    values.append(value)
                if row_index == number_of_rows_to_skip:
                    headers_row = ["file_name"] + values
                else:
                    all_rows.append([file_name.name] + values)
                row_index = row_index + 1

    # Save all consolidated rows in 1 single excel
    with open(source_directory / target_csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, dialect='excel', quotechar='"', quoting=csv.QUOTE_STRINGS,)
        writer.writerow(headers_row)
        for row in all_rows:
            print(row)
            writer.writerow(row)

    print("It's done! ")
    
if __name__ == "__main__":
    try:
        main(source_directory, target_csv_file, files_to_ignore)
    except Exception as e:
        print("Error!", e, traceback.format_exc())
        input("Error: fix it and try again after closing this terminal")