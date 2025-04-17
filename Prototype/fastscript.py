import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment

#You need to install both pandas and openpyxl > type into your cmd: pip install pandas openpyxl

# Global variables
file_name = ""  
cols = ["ID","TeamManager","Message","Complexity","SkillSet","RequestTopic","ResolutionDetails","RequestedBy","RequestResolvedBy","Resolution","Rating","RequestorKnowledgeBase","RequestorTicketNumber"]  # Columns to use
sheet_name = "Data"  # Worksheet name
output = r"C:\Users\rory.wilcox\Desktop\Fast\output.xlsx"     #You need to input your desired destination where the file will be saved, it may need to be local
client_list = ["OIR","DPC","PIM","QLDRA","OIC"]
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
   
if __name__ == "__main__":
    setup() 
    load_data()
    write_sheet()
    format_sheet()