#!/usr/bin/python3
from openpyxl import load_workbook
import re
import pandas as pd

def readXLSX(ExcelFile,planet):
    df = pd.read_excel(ExcelFile)
    df.insert(loc=0, column='Planet',value = planet)

    return df
    


jupiterDF = readXLSX("jupiter.xlsx","Jupiter")
saturnDF = readXLSX("saturn.xlsx","Saturn")
uranusDF = readXLSX("uranus.xlsx","Uranus")
neptuneDF = readXLSX("neptune.xlsx","Neptune")

allMoons = pd.concat([jupiterDF,
saturnDF,
uranusDF,
neptuneDF])

print(allMoons)

allMoons.to_excel("allMoons.xlsx", index = False)
