#!/usr/bin/python3
"""
File: readAllfiles.py
Author: John Major
Date: 2025-10-28
Description:  Read all 4 files and place into allMoons.xlsx and insert 
the planet name of the moon
"""

from openpyxl import load_workbook
import re
import pandas as pd

def readXLSX(ExcelFile,planet):
    df = pd.read_excel(ExcelFile)

    # add the planet name
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
