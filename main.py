import pandas
import re
import pyxlsb

from unidecode import unidecode
from enum import IntEnum

class Column(IntEnum):
    Code = 0
    Label = 1
    Nature = 4

# Returns the correct Nature
def CompareNature(label, expected, categories = {}, include_partial = False):

    # Sanitization (removal of spaces & stripping of accents)
    clean_label = unidecode(str(label)).replace(" ", "")
    clean_expected = unidecode(str(expected)).replace(" ", "")
    
    # Only 100% trusted source of truth
    match = re.search(clean_expected, clean_label, flags=re.IGNORECASE)
    if(match):
        return expected
    
    
    if (include_partial):
        for expected_word in clean_expected.split():
            match = re.search(expected_word, clean_label, flags=re.IGNORECASE)
            if(match):
                return expected

    for category in categories:
        match = re.search(unidecode(str(category)), clean_label, flags=re.IGNORECASE)
        if(match):
            return category

    return "Null"

def Start():
    # [TO-DO]: Replace input with direct file access
    path = input("Sheet Path: ")
    sheet = pandas.read_excel(path)

    columns = sheet.columns
    categories = {}

    for i in sheet[columns[Column.Nature]]:
        if(i not in categories):
            categories[i]=i


    for i, row in sheet.iterrows():
        label = row[columns[Column.Label]]
        nature = row[columns[Column.Nature]]

        res = CompareNature(label, nature, categories, True)
        if(res!=nature):
            # [Replace]: Debug info
            print(i)
            print(label)
            print(nature)
            print("ERR: "+res)


Start()
