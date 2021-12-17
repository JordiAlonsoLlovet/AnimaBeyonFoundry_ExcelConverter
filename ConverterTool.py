# importing the module
import json
import openpyxl
import string
import os
import re
import sys, getopt
from pathlib import Path
import time

# Setting the path to the xlsx file:
try:
    opt, args = getopt.gnu_getopt(sys.argv[1:], "hot:r:", longopts=['help', 'old', 'template=', 'resultPath='])
    for o, a in opt:
        if o in ["-h", "--help"]:
            print("usage: python ConverterTool.py [excel path] [nombre PJ]")
            print("Opciones:")
            print("-r (--resultPath) [path] : Directorio donde guardar el archivo JSON")
            print("-t (--template) [path]   : Directorio de la plantilla JSON. Por defecto es ./actorSmart.json")
            print("-o (--old)               : Excel de version antigua")
            exit()

    xlsx_file = os.path.join(args[0]) #os.path.join('C:\\x\\tools', 'Elienai.xlsm') 
    name = args[1]
    file_name = name.replace(" ", "_")+".json"
    if(not os.path.exists(".\\results")):
                os.makedirs(".\\results")
    targetFile = os.path.join(".\\results", file_name)
    templateFile = 'actorSmart.json'
    actor = {
        "name": name,
        "type": "character"
    }
    genericKiCell = "F24"
    for o, a in opt:
        if o in ["-o", "--old"]:
            genericKiCell = "E24"
        elif o in ["-t", "--template"]:
            templateFile = a
        elif o in ["-r", "--resultPath"]:
            if(not os.path.exists(a)):
                os.makedirs(a)
            targetFile = os.path.join(a, file_name)
       

    wb = openpyxl.load_workbook(xlsx_file, data_only=True)
except Exception as e:
    print(e)
    time.sleep(15)

def getCell(cell):
    sheet = wb.active
    return sheet[cell].value

def getValue(name):
    if (name[0] == '$'):
        return getCell(name[1:])
    # get DefinedNameList instance
    defined_name = wb.defined_names[name]

    # get destinations which is a generator
    destinations = defined_name.destinations  
    values = {}
    for sheet_title, coordinate in destinations:
        values.update({sheet_title: wb[sheet_title][coordinate].value})
    my_value = values.get(sheet_title)
    return my_value
    
try:
    for s in range(len(wb.sheetnames)):
        if wb.sheetnames[s] == 'Principal':
            break
    wb.active = s
    principal = wb.active
    with open(templateFile) as json_file:
        data = json.load(json_file)

        for key in data["characteristics"]["primaries"]:
            data["characteristics"]["primaries"][key]["value"] = getValue(data["characteristics"]["primaries"][key]["value"])

        #data["characteristics"]["primaries"]["constitution"]["value"] = getValue('CON')
        #data["characteristics"]["primaries"]["dexterity"]["value"] = getValue('DES')
        #data["characteristics"]["primaries"]["strength"]["value"] = getValue('FUE')
        #data["characteristics"]["primaries"]["intelligence"]["value"] = getValue('INT')
        #data["characteristics"]["primaries"]["perception"]["value"] = getValue('PER')
        #data["characteristics"]["primaries"]["agility"]["value"] = getValue('AGI')
        #data["characteristics"]["primaries"]["power"]["value"] = getValue('POD')
        #data["characteristics"]["primaries"]["willPower"]["value"] = getValue('VOL')

        data["characteristics"]["secondaries"]["lifePoints"]["value"] = principal["N11"].value
        data["characteristics"]["secondaries"]["lifePoints"]["max"] = principal["N11"].value
        data["characteristics"]["secondaries"]["initiative"]["base"]["value"] = getValue('Turno_Nat_final')
        data["characteristics"]["secondaries"]["fatigue"]["value"] = principal["N16"].value
        data["characteristics"]["secondaries"]["fatigue"]["max"] = principal["N16"].value
        data["characteristics"]["secondaries"]["movementType"]["final"]["value"] = getValue('TipodeMovimiento')
        data["characteristics"]["secondaries"]["resistances"]["physical"]["base"]["value"] = principal["J58"].value
        data["characteristics"]["secondaries"]["resistances"]["disease"]["base"]["value"] = principal["J59"].value
        data["characteristics"]["secondaries"]["resistances"]["poison"]["base"]["value"] = principal["J60"].value
        data["characteristics"]["secondaries"]["resistances"]["magic"]["base"]["value"] = principal["J61"].value
        data["characteristics"]["secondaries"]["resistances"]["psychic"]["base"]["value"] = principal["J62"].value

        for field in data["secondaries"]:
            for key in data["secondaries"][field]:
                data["secondaries"][field][key]["base"]["value"] = getValue(data["secondaries"][field][key]["base"]["value"])

        for i in range(73, 77):
            valor = principal.cell(row=i, column=17).value
            if valor != "-":
                secondary = {"base":{"value":valor}, "final":{"value": 0}}
                data["secondaries"]["secondarySpecialSkills"].update({principal.cell(row=i, column=13).value: secondary})

        data["combat"]["attack"]["base"]["value"] = getValue('HA_final')
        data["combat"]["block"]["base"]["value"] = getValue('HP_final')
        data["combat"]["dodge"]["base"]["value"] = getValue('HE_final')
        data["combat"]["wearArmor"]["value"] = getValue('LlevarArmadura_final')

        for s in range(len(wb.sheetnames)):
            if wb.sheetnames[s] == 'Místicos':
                break
        wb.active = s
        magic = wb.active

        #data["mystic"]["zeonRegeneration"]["base"]["value"] = getValue('ACT_final')
        data["mystic"]["act"]["main"]["base"]["value"] = getValue('ACT_final')
        data["mystic"]["zeon"]["value"] = getValue('Zeón_final')
        data["mystic"]["zeon"]["max"] = getValue('Zeón_final')
        data["mystic"]["magicProjection"]["base"]["value"] = getValue('ProyecciónMágica_final')
        #TODO Imbalance

        for path in data["mystic"]["magicLevel"]["spheres"]:
            via = data["mystic"]["magicLevel"]["spheres"][path]["value"]
            for i in range(15, 25):
                if magic.cell(row=i, column=3).value == via:
                    data["mystic"]["magicLevel"]["spheres"][path]["value"] = magic.cell(row=i, column=8).value
            if data["mystic"]["magicLevel"]["spheres"][path]["value"] == via:
                data["mystic"]["magicLevel"]["spheres"][path]["value"] = 0

        data["mystic"]["magicLevel"]["total"]["value"] = getValue("NiveldeMagia_final")
        #data["mystic"]["magicLevel"]["used"]["value"] = 

        for skill in data["mystic"]["summoning"]:
            data["mystic"]["summoning"][skill]["base"]["value"] = getValue(data["mystic"]["summoning"][skill]["base"]["value"])

        for s in range(len(wb.sheetnames)):
            if wb.sheetnames[s] == 'Ki':
                break
        wb.active = s
        ki = wb.active

        data["domine"]["kiAccumulation"]["generic"]["max"] = ki[genericKiCell].value
        data["domine"]["kiAccumulation"]["generic"]["value"] = ki[genericKiCell].value
        data["domine"]["martialKnowledge"]["max"]["max"] = getValue("CM_final")
        for acu in  data["domine"]["kiAccumulation"]:
            if acu != "generic":
                data["domine"]["kiAccumulation"][acu]["base"]["value"] = getValue(data["domine"]["kiAccumulation"][acu]["base"]["value"])

        for s in range(len(wb.sheetnames)):
            if wb.sheetnames[s] == 'Psíquico':
                break
        wb.active = s
        psi = wb.active

        data["psychic"]["psychicPotential"]["base"]["value"] = psi["H11"].value
        data["psychic"]["psychicProjection"]["base"]["value"] = getValue('Proyecciónpsíquica_final')
        data["psychic"]["psychicPoints"]["value"] = psi["I16"].value
        data["psychic"]["psychicPoints"]["max"] = psi["I16"].value
        data["psychic"]["innatePsychicPower"]["amount"]["value"] = psi["M13"].value

        actor.update({"data": data})


    with open(targetFile, 'w') as outfile:
        json.dump(actor, outfile)
except Exception as e:
    print(e)
    time.sleep(15)



    

