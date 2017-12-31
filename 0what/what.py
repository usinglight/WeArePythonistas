#!/usr/bin/env python
# # -*- coding: utf-8 -*-

"""WeArePythonistas - Part 0: What do we want to see"""

from openpyxl import load_workbook
from operator import itemgetter
import matplotlib.pyplot as plt

# I put the file on my machine so I have a clear privacy separation
file='/Users/stefansteinbauer/Documents/onmachinedata/report.xlsx'
wsName='Attendees'
wb = load_workbook(file)
ws=wb[wsName]

# read first 700 values from column 20 (languages) into list
# not very pythonic, but here you also see the stripping and lowercasing to
# get through the masses of different ways to write Javascript
interests=list()
for i in range(2, 701):
    cellval = str(ws.cell(row=i, column=20).value)
    cellval = cellval.replace(' ','')
    cellval = cellval.upper()
# and here I look at people filling in their languageS
    interests.extend(cellval.split(","))

#counting the languages while removing the duplicates
lang2={x:interests.count(x) for x in interests}
# Getting the Top 10
lang = dict(sorted(lang2.items(), key=itemgetter(1), reverse=True)[:15])

# let's print that - xkcd style
plt.xkcd()

fig = plt.figure()
ax = fig.add_subplot(1, 1, 1)
ax.spines['right'].set_color('none')
ax.spines['top'].set_color('none')
plt.gcf().subplots_adjust(left=0.3)

plt.title("WHAT WE WANT")
plt.rcParams.update({'figure.autolayout': True})
plt.barh(range(len(lang)),list(lang.values()), align='center')
plt.yticks(range(len(lang)), list(lang.keys()))

plt.show()
