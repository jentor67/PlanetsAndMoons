#!/usr/bin/python3
import pandas as pd
from urllib.request import Request, urlopen


def getMoons(webSite,tableNumber,excelFile):
    hdr = {'User-Agent': 'Mozilla/5.0'}
    req = Request(webSite,headers=hdr)
    page = urlopen(req)

    tables = pd.read_html(page)
    df = tables[tableNumber]  
    print(df)
    df.to_excel(excelFile, index = False)




jupiterMoons = "https://en.wikipedia.org/wiki/Moons_of_Jupiter"
#getMoons(jupiterMoons,1,"jupiter.xlsx")

saturnMoons = "https://en.wikipedia.org/wiki/Moons_of_Saturn"
#getMoons(saturnMoons, 2,"saturn.xlsx")

uranusMoons = "https://en.wikipedia.org/wiki/Moons_of_Uranus"
#getMoons(uranusMoons,1,"uranus.xlsx")

neptuneMoons = "https://en.wikipedia.org/wiki/Moons_of_Neptune"
#getMoons(neptuneMoons,1,"neptune.xlsx")
