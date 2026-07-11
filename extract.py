# Extract 20 entries for each Nature, to be used as training data
import pandas
import pyxlsb
    
def Start():
    # [TO-DO]: Replace input with direct file access
    path = input("Sheet Path: ")
    sheet = pandas.read_excel(path)

    entries_by_nature = sheet.groupby("Nature")

    entries = {}
    for nature, row in entries_by_nature:
        entries[nature] = row.head(20)


    # [Replace]: Debug info
    print(entries)

Start()