import pandas as pd
from openpyxl import load_workbook
# it probably makes sense to remove most of the inputs here and just make people edit the script. especially the columns 
# Global variables
file_name = ""  # you should need the file to be in the same folder as the script, you don't need the file extension if you dont want it. the script should handle it
cols = []       # Which cols you want to use, You can use names or index, I could also allow you to use lettered cols if you really wanted
sheet_name = '' # This likely could just be "Clean_data" or something if the input is annoying
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

if __name__ == "__main__":
    setup()
    load_data()
    check_name()
    # should be able to slap in your own functions here if you want to do some fancy things to the sheet
    write_sheet()


