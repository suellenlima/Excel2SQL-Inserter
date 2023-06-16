# -*- coding: utf-8 -*-
import openpyxl

def create_insert_script(filename):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    script = ""
    id = 14
    for row in sheet.iter_rows(values_only=True):
        category = str(row[0])
        question = str(row[2])
        answer = str(row[3])
        teacher = str(row[4])
        coordinator = str(row[5])
        
        isTeacher = True if teacher == 'X' else False
        isCoordinator = True if coordinator == 'X' else False
        
        if isTeacher:
            id = id + 1
            insert_command = f"INSERT INTO "+'"'+"FAQ"+'"'+" ("+'"'+"ID"+'"'+", "+'"'+"QUESTION"+'"'+", "+'"'+"ANSWER"+'"'+", "+'"'+"ID_CATEGORY"+'"'+", "+'"'+"IS_ACTIVE"+'"'+") VALUES ("+str(id)+", "+"'"+question+"'"+", "+"'"+answer+"'"+", (select "+'"'+"ID"+'"'+" from "+'"'+"FAQ_CATEGORY"+'"'+"where "+'"'+"PROFILE"+'"'+" = 3 and "+'"'+"NAME"+'"'+" = "+"'"+str(category)+"'"+"), true);\n"
            script += insert_command 
        if isCoordinator:
            id = id + 1
            insert_command = f"INSERT INTO "+'"'+"FAQ"+'"'+" ("+'"'+"ID"+'"'+", "+'"'+"QUESTION"+'"'+", "+'"'+"ANSWER"+'"'+", "'"'+"ID_CATEGORY"+'"'+", "+'"'+"IS_ACTIVE"+'"'+") VALUES ("+str(id)+", "+"'"+question+"'"+", "+"'"+answer+"'"+", (select "+'"'+"ID"+'"'+" from "+'"'+"FAQ_CATEGORY"+'"'+"where "+'"'+"PROFILE"+'"'+" = 2 and "+'"'+"NAME"+'"'+" = "+"'"+str(category)+"'"+"), true);\n"
            script += insert_command 

    with open("insert.sql", "w", encoding='utf-8') as f:
        f.write(script)

create_insert_script('C:\learn\Excel2SQL-Inserter\Excel2SQL-Inserter\FAQNeuroinfinity.xlsx')
