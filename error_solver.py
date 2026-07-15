import pandas
import openpyxl
from transformers import pipeline

from enum import IntEnum

class Column(IntEnum):
    Code = 1
    Label = 2
    Universe = 3
    Nature = 4
    Found = 5
    
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




def Start():    
    # [TO-DO]: Replace input with direct file access
    sheet_path = "errors.xlsx" #input("Sheet Path: ")
    sheet = pandas.read_excel(sheet_path)

    columns = sheet.columns
    categories = []

    for i in sheet[columns[Column.Nature]]:
        if(i not in categories):
            categories.append(i)


    edits = {
        columns[Column.Code]: [],
        columns[Column.Nature]: [],
    }

    for i, row in sheet.iterrows():
        code = str(row[columns[Column.Code]])
        label = str(row[columns[Column.Label]])
        nature = str(row[columns[Column.Nature]])
        universe = str(row[columns[Column.Universe]])
        found = str(row[columns[Column.Found]])

        if(found!="n/a"):
            found_match = GetMatch(label, found)
            if(found_match):
                edits[columns[Column.Code]].append(code)
                edits[columns[Column.Nature]].append(found)
                continue

        nature_match = IsMatch(label, nature)
        universe_match = IsMatch(label, universe)
        
        if(nature_match and universe_match):
            continue
        else:    
            edits[columns[Column.Code]].append(code)
            edits[columns[Column.Nature]].append(GetBestMatch(label, universe, nature, categories)["category"])

    
    pandas.DataFrame(edits).to_excel("edits.xlsx")


Start()

