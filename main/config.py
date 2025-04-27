#File Paths
file_name = ""  #currently the script asks for this so you can leave it blank. you can edit the input function out and set a designated file path here instead.
output = r"C:\Users\output.xlsx"  #put your own desired file path here

# Script configuration
cols = ["ID","TeamManager","Message","Complexity","SkillSet","RequestTopic","ResolutionDetails","RequestedBy","RequestResolvedBy","Resolution","Rating","RequestorKnowledgeBase","RequestorTicketNumber"]  # Columns to use
sheet_name = "Data"  # Worksheet name

# Other Data

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

