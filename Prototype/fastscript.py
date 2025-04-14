import pandas as pd
from openpyxl import load_workbook

# Global variables
file_name = ""  # Will be set in setup()
cols = []       # Which cols you want to use
sheet_name = '' # Will be set in setup()
df = None       # Will hold the DataFrame

def setup():
    global file_name, sheet_name, cols
    
    file_name = input("Enter File Name: ")
    if not file_name.lower().endswith(('.xlsx', '.xls')):
        file_name = file_name + '.xlsx'
        print(f"Using file name: {file_name}")
    
    # Get columns to use
    cols_input = input("Enter column names or indices separated by commas: ")
    cols = [col.strip() for col in cols_input.split(',')]
    
    # Get sheet name
    sheet_name = input("Enter name for the new sheet: ")

def load_data():
    global df
    df = pd.read_excel(file_name, header=0, usecols=cols)
    print("Data loaded:")
    print(df)

def check_name():
    global book
    book = load_workbook(file_name)
    if sheet_name in book.sheetnames:
        response = input(f"Sheet '{sheet_name}' already exists. Replace it? (yes/no): ")
        if response.lower() == 'yes':
            book.remove(book[sheet_name])
            print(f"Removed existing sheet '{sheet_name}'")
        else:
            print("Operation cancelled.")
            exit()

def write_sheet():
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Data successfully written to sheet '{sheet_name}'")

# Execute the functions
if __name__ == "__main__":
    setup()
    load_data()
    check_name()
    write_sheet()