import os
from pyautocad import Autocad
import win32com.client
import pandas as pd 
import argparse



def find_and_replace(file):
    #Read to excel file
    df = pd.read_excel(file, sheet_name=0)
    files = df['files'].tolist()
    find = df['find'].tolist()
    replace_with = df['replace'].tolist()

    dico = {}

    for index in range(len(files)):
        dico[str(files[index])] = [str(find[index]), str(replace_with[index])]
    #get files frm current dir and current dir path
    dwgfiles = filter(os.path.isfile, os.listdir( os.curdir ) )
    cwd = os.path.abspath(os.path.curdir) #current working dir

    
    for f in dwgfiles: #loop on files
        if f.endswith(".dwg"): #Choose Dwg files
            #Initiate AutoCad APP
            acad = win32com.client.Dispatch("AutoCAD.Application") 
            acad.Visible = True
            acad = Autocad() 

            #Open document
            acad.app.Documents.open(cwd + "/" + f)

            print(acad.doc.Name) #Curent file name

            doc = acad.ActiveDocument   # Document object 
            #Traverse the cad image object,Modifying Object Properties
            for entity in acad.ActiveDocument.ModelSpace:
                name = entity.EntityName
                if name == 'AcDbMText':
                    if dico[f][0] in entity.TextString:
                        #Modify object properties
                        text1 = str(entity.TextString)
                        text1 = text1.replace (dico[f][0], dico[f][1])
                        entity.TextString = text1
                elif name == 'AcDbText':
                    if dico[f][0] in entity.TextString:
                        #Modify object properties
                        text1 = str(entity.TextString)
                        text1 = text1.replace (dico[f][0], dico[f][1])
                        entity.TextString = text1
                
            doc.Close(True)
            acad = None
            print(f, ", Done")


def Main():
    parser = argparse.ArgumentParser()
    #Arguments 
    parser.add_argument("Fichier_excel", help="Excel file's name with 3 columns named: 'Files' DWG files' names, 'find' Text entities to find, 'replace' text to replace the associated entity", type= str)

    args = parser.parse_args()

    find_and_replace(args.Fichier_excel)

if __name__ == '__main__':
    Main()