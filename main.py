import pandas
import re
import sys
import pyxlsb
import openpyxl

from fuzzysearch import find_near_matches
from transformers import pipeline
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
    
    
def GetMatch(label, category):
    classifier = pipeline("zero-shot-classification", model="MoritzLaurer/mDeBERTa-v3-base-mnli-xnli")
    processed_results = classifier(label, category, multi_label=False)
    
    return processed_results

def IsMatch(match, acceptable_threshold = 0.02):
    return GetMatch(label, category)["scores"][0] > acceptable_threshold

def GetBestMatches(label, *possible_categories, acceptable_threshold = 0.02):
    classifier = pipeline("zero-shot-classification", model="MoritzLaurer/mDeBERTa-v3-base-mnli-xnli")
    processed_results = classifier(label, possible_categories[0], multi_label=False)

    best_matches = []

    for i in range(len(processed_results["labels"])):
        score = processed_results["scores"][i]
        category = processed_results["labels"][i]
        
        if(score>acceptable_threshold):
            best_matches.append({"category":category, "score":score})

    return best_matches

def GetBestMatch(label, universe, category, *possible_categories):
    matches = GetBestMatches(label, possible_categories[0])

    if(len(matches)==1):
        return matches[0]

    categories = []
    for match in matches:
        categories.append(match["category"])
    
    universe_matches = GetMatch(universe, categories)
    universe_matches.sort(key=lambda e : e["score"])

    print(universe_matches)
    return universe_matches[0]







# Find entries with potentially erroneous categories
def Find(sheet):
    columns = sheet.columns
    categories = {}

    for category in sheet[columns[Column.Nature]]:
        if(category not in categories):
            categories[category]=category


    errors = {
        columns[Column.Code]: [],
        columns[Column.Label]: [],
        columns[Column.Universe]: [],
        columns[Column.Nature]: [],
        Column.Found: []
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
                errors[Column.Found].append(str(res))
            else :
                errors[Column.Found].append("N/A")



    return errors



# Find an appropriate category for each entry marked as potentially erroneous
def Solve(sheet, errors):

    columns = sheet.columns
    categories = []

    for category in sheet[columns[Column.Nature]]:
        if(category not in categories):
            categories.append(category)


    edits = {}

    for i in range(len(errors[Column.Found])):
        code = str(errors[columns[Column.Code]])
        label = str(errors[columns[Column.Label]])
        nature = str(errors[columns[Column.Nature]])
        universe = str(errors[columns[Column.Universe]])
        found = str(errors[Column.Found])

        if(found!="n/a"):
            found_match = GetMatch(label, found)
            if(found_match):
                edits[code] = {Column.Nature: found}
                continue

        nature_match = IsMatch(label, nature)
        universe_match = IsMatch(label, universe)
        
        if(nature_match and universe_match):
            continue
        else:        
            edits[code] = {Column.Nature: GetBestMatch(label, universe, nature, categories)["category"]}

    
    return edits
    

# Get Colors & Dimensions from Label
def Extra(sheet):
    colors_path = "colors.json"
    colors = pandas.read_json(colors_path)

    color_names = {}

    for i in colors["name"]:
        if(i not in color_names):
            color_names[unidecode(str(i).upper())] = i

    columns = sheet.columns

    extraneous_information = {}

    for i, row in sheet.iterrows():
        code = str(row[columns[Column.Code]])
        label = str(row[columns[Column.Label]])

        extraneous_information[code] = {Column.Colors: GetColors(label, color_names), Column.Dimensions: GetDimensions(label)}

    return extraneous_information


# Merges edits to sheet
def Merge(sheet, edits):
    columns = sheet.columns

    for i, row in sheet.iterrows():
        code = row[columns[Column.Code]]
        
        if(code in edits):
            for column in edits[code].keys:                    
                row[columns[column]] = edits[code][column]

    return sheet






def Start():
    sheet_path = str(sys.argv[1])
    sheet = pandas.read_excel(sheet_path)

    sheet["Couleurs"] = ""
    sheet["Dimensions"] = ""
    

    extras = Extra(sheet)

    err = Find(sheet)
    edits = Solve(sheet, err)
    fixed = Merge(sheet, edits)

    output_file_name = "fixed_"+re.sub(r"(?!\.)[^.]+$", "xls", sheet_path)
    sheet.to_excel(output_file_name, index=False)


Start()

