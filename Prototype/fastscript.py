import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import re

#You need to install both pandas and openpyxl > type into your cmd: pip install pandas openpyxl

# Global variables
file_name = ""  
cols = ["ID","TeamManager","Message","Complexity","SkillSet","RequestTopic","ResolutionDetails","RequestedBy","RequestResolvedBy","Resolution","Rating","RequestorKnowledgeBase","RequestorTicketNumber"]  # Columns to use
sheet_name = "Data"  # Worksheet name
output = r"C:\Users\rory.wilcox\Desktop\Fast\output.xlsx"     #You need to input your desired destination where the file will be saved, it may need to be local

client_list = {
    "oir":["oir", "office of indsutrial relations", "the office of industrial relations", "industrial relations", "wrc", "workers comp", "workers compensation","qirc"],
    "pinnacle":["pim", "pinnacle", "palisade", "hyperion", "maplebrown", "maple-brown", "maplebrown abbott","maple-brown abott", "maple-brown abbott", "anitpodes", "pinnacle investments","palisade asset management","firetrail","hyperion","plato","spheria","two trees","solaris","res cap","rez cap","resolution capital"],
    "oic":["oic", "office of the information commissioner", "office of information comissioner", "office information comissioner"],
    "qldra":["qldra", "qra", "queensland reconstruction authority", "reconstruction authority", "queensland reconstruction"],
    "nsu":["nsu", "neurosensory","neuro","neurosense"],
    "dpc":["dpc", "department premier cabinet", "departmentofthepremiercabinet", "psc","premiers","oog","oqpc","tcis","integrity commissioner"],
    "ozcare":["ozcare", "oz care", "oscare","ozc"],
    "qtco":["qtco","queensland treasury corporation","qtc"],
    "opq":["opq"],
    "olgr":["olgr"]
    }

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
    setup() 
    load_data()
    add_client_column()
    find_client()
    print(df)
    write_sheet()
    format_sheet()
