#Script to extract data out of all the tables in the /data dirctory

import pandas as pd
from openpyxl import Workbook
import openpyxl
import json
#instatntiate workbook object
#testing opnpyxl
# wb = openpyxl.load_workbook("data/WesternCape_LS.xlsx")
# print(wb.sheetnames)
#end of testing ====
def get_Sheet_Data(sheet):
    print(sheet)
    return
#Get excel sheet based on excel filename and sheetname
def get_Sheet_From_File(filename, sheetname):
    workbook =  openpyxl.load_workbook(filename,data_only=True)
    try:
        print(workbook.sheetnames)
        worksheet = workbook[sheetname]
    except:
        print("ERROR%******")
        return
    return worksheet

#Get loadshedding data 
def get_loadshedding_data(ws):
    list_of_schemas = []
    data_schema = {}
    for row in range(16,112):
        print(ws.cell(row,3).value)
        if(ws.cell(row,3).value == 1):
            # refresh data schema 
            data_schema = {
        "start":str(ws.cell(row,1).value),
        "end":str(ws.cell(row,2).value),
        "loadshedding":{
            1:{},
            2:{},
            3:{},
            4:{},
            5:{},
            6:{},
            7:{},
            8:{},
        }
            }
            #start of if indent
            print(f"{ws.cell(row,1).value}-{ws.cell(row,2).value}")
        rowData = ""
        #add start time 
        #add start time for specefic group
        row_list = []
        for col in range (4,35):
            # print(f"=======getting data {row}==========")
            row_list.append(ws.cell(row,col).value)
        #add stage no and groups affected from day 1-31
        data_schema["loadshedding"][ws.cell(row,3).value] = row_list
        print(row_list)
        #when the 8th group is reached start from 1 again
        if ws.cell(row,3).value ==8:
            print(data_schema)
            list_of_schemas.append(data_schema)

        # print(ws.cell(111,3).value)
    return list_of_schemas

if __name__ == "__main__":
    filename ="data/WesternCape_LS.xlsx"
    sheet = get_Sheet_From_File(filename,sheetname="Schedule")
    get_Sheet_Data(sheet)
    dict_table = get_loadshedding_data(sheet)
    tuple_obj = (*dict_table,)
    # dict_data = json.loads(str(dict_table))
    jsonObj = json.dumps(dict_table)
    with open("ls_data.json", "w") as outfile:
        outfile.write(jsonObj)