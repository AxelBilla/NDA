# Extract X entries for each Nature, to be used as training data
import pandas
import pyxlsb

sample_quantity = 20
    
def Start():
    # [TO-DO]: Replace input with direct file access
    path = "20210614 Ecommerce sales.xlsb" #input("Sheet Path: ")
    sheet = pandas.read_excel(path)

    entries_by_nature = sheet.groupby("Nature")

    for nature, row in entries_by_nature:
        entries[nature] = row.head(sample_quantity)


    print(entries)

Start()
