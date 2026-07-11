import pandas
import re
import pyxlsb

from unidecode import unidecode
from enum import IntEnum

class Column(IntEnum):
    Code = 0
    Label = 1
    Nature = 4

# Returns the correct Nature of an item
def GetNature(label, expected, categories = {}, include_partial = False):

    # Sanitization (removal of spaces & stripping of accents)
    clean_label = unidecode(str(label)).replace(" ", "")
    clean_expected = unidecode(str(expected)).replace(" ", "")
    
    # Only 100% trusted source of truth
    match = re.search(clean_expected, clean_label, flags=re.IGNORECASE)
    if(match):
        return expected


    for category in categories:
        match = re.search(unidecode(str(category)), clean_label, flags=re.IGNORECASE)
        if(match):
            return category

    
    if (include_partial):
        for expected_word in clean_expected.split():
            match = re.search(expected_word, clean_label, flags=re.IGNORECASE)
            if(match):
                return expected

    return "n/a"



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
    





def Start():
    # [TO-DO]: Replace input with direct file access
    sheet_path = input("Sheet Path: ")
    sheet = pandas.read_excel(sheet_path)

    columns = sheet.columns
    categories = {}

    for i in sheet[columns[Column.Nature]]:
        if(i not in categories):
            categories[i]=i




    colors_path = "colors.json"
    colors = pandas.read_json(colors_path)

    color_names = {}

    for i in colors["name"]:
        if(i not in color_names):
            color_names[unidecode(str(i).upper())] = i





    for i, row in sheet.iterrows():
        label = row[columns[Column.Label]]
        nature = row[columns[Column.Nature]]

        print("COLORS: "+str(GetColors(label, color_names)))
        print("DIMENSIONS: "+str(GetDimensions(label)))

        res = GetNature(label, nature, categories, True)
        if(res!=nature):
            # [Replace]: Debug info
            print(i)
            print(label)
            print(nature)
            print("FOUND: "+str(res))


Start()
