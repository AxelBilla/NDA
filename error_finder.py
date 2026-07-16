import pandas
import re
import sys
import pyxlsb
import openpyxl

from fuzzysearch import find_near_matches
from unidecode import unidecode
from enum import IntEnum

class Column(IntEnum):
    Code = 0
    Label = 1
    Seller = 2
    Universe = 3
    Nature = 4
    Date = 5
    Quantity = 6
    Price = 7
    Duration = 8
    Description = 9
    Colors = 10
    Dimensions = 11

    Found = 12

# Returns the correct Nature of an item
def GetNature(label, expected, categories = {}, include_partial = False):

    # Sanitization (removal of spaces & stripping of accents)
    clean_label = unidecode(str(label)).replace(" ", "")
    clean_expected = unidecode(str(expected)).replace(" ", "")
    
    # Only 100% trusted source of truth
    match = find_near_matches(clean_expected, clean_label, max_l_dist=1)
    if(match):
        return expected

    if (include_partial):
        for expected_word in clean_expected.split():
            match = find_near_matches(expected_word, clean_label, max_l_dist=1)
            if(match):
                return expected

    for category in categories:
        match = find_near_matches(unidecode(str(category)), clean_label, max_l_dist=1)
        if(match):
            return category

    return "n/a"




def Start():
    sheet_path = str(sys.argv[1])
    sheet = pandas.read_excel(sheet_path)

    columns = sheet.columns
    categories = {}

    for i in sheet[columns[Column.Nature]]:
        if(i not in categories):
            categories[i]=i



    errors = {
        columns[Column.Code ]: [],
        columns[Column.Label]: [],
        columns[Column.Universe]: [],
        columns[Column.Nature]: [],
        "FOUND": []
    }

    for i, row in sheet.iterrows():
        label = row[columns[Column.Label]]
        nature = row[columns[Column.Nature]]
        universe = row[columns[Column.Universe]]

        res = GetNature(label, nature, categories, True)
        if(res!=nature):
            code = str(row[columns[Column.Code]])

            if(code):
                errors[columns[Column.Code]].append(code)
            else :
                errors[columns[Column.Code]].append("N/A")

            if(label):
                errors[columns[Column.Label]].append(label)
            else :
                errors[columns[Column.Label]].append("N/A")
                

            if(nature):
                errors[columns[Column.Nature]].append(nature)
            else :
                errors[columns[Column.Nature]].append("N/A")
                

            if(universe):
                errors[columns[Column.Universe]].append(universe)
            else :
                errors[columns[Column.Universe]].append("N/A")

            if(str(res)):
                errors["FOUND"].append(str(res))
            else :
                errors["FOUND"].append("N/A")



    pandas.DataFrame(errors).to_excel("test.xlsx", index=False)


Start()

