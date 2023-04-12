import openpyxl as op
import json
import pandas as pd
from pandas import *

# Batch : {Degree : { Section : { Subject : { Room : Timings } } } }

# batch color dict
batchColor = {
    "FFFFB740": "2022",
    "FF7F6000": "2021",
    "FFF1C232": "2020",
    "FFFFE599": "2019",
    "FF7F4CFF": "2022",
    "FF351C75": "2021",
    "FFB17FD7": "2020",
    "FFB4A7D6": "2019",
    "FF00F600": "2022",
    "FF274E13": "2021",
    "FF6AA84F": "2020",
    "FFB6D7A8": "2019",
    "FFE62C06": "2022",
    "FF85200C": "2021",
    "FFDD7E6B": "2020",
    "FFF4CCCC": "2019",
    "FF0000FF": "2022",
    "FF0050EF": "2021",
    "FF599DDA": "2020",
    "FFABCCEB": "2019"
}

test = {
    "Batch": {"Degree": {"Section": {"Course": {"Day": {"Room": "Timings"}}}}},
}

# open the workbook
wb = op.load_workbook('FSCTimeTable.xlsx')

# get the sheet
sheetMonday = wb['Monday']
sheetTuesday = wb['Tuesday']
sheetWednesday = wb['Wednesday']
sheetThursday = wb['Thursday']
sheetFriday = wb['Friday']

# print(test["Batch"])

data = [[]]

with open("timetable_copy.json", "r") as jsonFile:
    data = json.load(jsonFile)

data[0]["Department"] = "NewPath"

with open("timetable_copy.json", "w") as jsonFile:
    json.dump(data, jsonFile)

print(data)