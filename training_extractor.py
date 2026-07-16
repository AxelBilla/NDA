# Extract X entries for each Nature, to be used as training data
import pandas
import sys
import pyxlsb

sample_quantity = 20
    
def Start():
    sheet_path = str(sys.argv[1])
    sheet = pandas.read_excel(path)

    entries_by_nature = sheet.groupby("Nature")

    for nature, row in entries_by_nature:
        entries[nature] = row.head(sample_quantity)


    print(entries)

Start()
