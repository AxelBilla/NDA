import pandas
import re
import pyxlsb
import openpyxl

from fuzzysearch import find_near_matches
from transformers import pipeline
from unidecode import unidecode
from enum import IntEnum

class Column(IntEnum):
    Code = 0
    Label = 1
    Colors = 10
    Dimensions = 11

# Identifies palette colors in label
def GetColors(label, palette = {}):
    clean_label = unidecode(str(label).upper())
    colors = []

    for word in clean_label.split():
        if(word in palette):
            colors.append(palette[word])

    return colors

# Identifies dimensions in label
def GetDimensions(label = ""):
    res = re.search(r"((\d+\s*((\.|,|\*|x)\s*\d+\s*)+)|(\d+))(?=\s*cm)", label, re.IGNORECASE)
    if(res):
        return res[0].replace(" ", "").replace("x","*")
    else:
        return "n/a"


# Get Colors & Dimensions from Label
def Extra(sheet):
    colors_path = "colors.json"
    colors = pandas.read_json(colors_path)

    color_names = {}

    for i in colors["name"]:
        if(i not in color_names):
            color_names[unidecode(str(i).upper())] = i

    columns = sheet.columns

    extraneous_information = {
        columns[Column.Code]: [],
        columns[Column.Colors]: [],
        columns[Column.Dimensions]: []
    }

    for i, row in sheet.iterrows():
        code = row[columns[Column.Code]]
        label = row[columns[Column.Label]]

        extraneous_information[columns[Column.Code]].append(GetColors(label, color_names))
        extraneous_information[columns[Column.Colors]].append(GetColors(label, color_names))
        extraneous_information[columns[Column.Dimensions]].append(GetDimensions(label))

    return extraneous_information




def Start():
    # [TO-DO]: Replace input with direct file access
    sheet_path = "20210614 Ecommerce sales.xlsb" #input("Sheet Path: ")
    sheet = pandas.read_excel(sheet_path)

    sheet.insert(Column.Colors, "Couleurs", [], True)
    sheet.insert(Column.Dimensions, "Dimensions", [], True)

    extras = Extra(sheet)

    pandas.DataFrame(extras).to_excel("extras.xlsx")


Start()

