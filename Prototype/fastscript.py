from dependencies import *
from config import *

#You need to install both pandas and openpyxl > type into your cmd: pip install pandas openpyxl


def setup():
    global file_name
    file_name = input("Enter File Name: ")#asks user for the target file
    if not file_name.lower().endswith(('.xlsx', '.xls')):#if the user forgot to add the file extension this should handle it
        file_name = file_name + '.xlsx'
    print(f"Using file name: {file_name}")

def load_data():
    global df
    df = pd.read_excel(file_name, header=0, usecols=cols)
    print("Data loaded:")
    print(df)#Printing the DF here so the user at a glance can see if it's run properly

def add_client_column():
    df.insert(0, "Client", "")


def write_sheet():
    global writer
    writer = pd.ExcelWriter(output, engine="openpyxl") #Initialising the writer object which will turn our DF into an excel doc
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.close()

def format_sheet():
    # Load the saved workbook
    workbook = load_workbook(output)
    worksheet = workbook[sheet_name]
    max_length = 0
  
    for column_cells in worksheet.columns:
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            cell.alignment = Alignment(wrap_text=True)
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
                new_width = max_length + 1
                worksheet.column_dimensions[column_letter].width = new_width
    # Save the formatted workbook
    workbook.save(output)

def find_client():
    for index, message in df['Message'].items():
        message = message.lower()
        for list_name, clients in client_list.items():
            if any(re.search(r'\b' + re.escape(client.lower()) + r'\b', message) for client in clients):
                df.at[index, 'Client'] = list_name
                break



    
   
if __name__ == "__main__":
    #Initialise
    setup() 
    load_data()
    #Modify
    add_client_column()
    find_client()
    print(df)
    #Finalise
    write_sheet()
    format_sheet()
