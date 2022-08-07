# -*- coding: utf-8 -*-
"""
Created on Tue Jul  5 12:23:58 2022

@author: efull

Name: Elijah Journey Fullerton
email: ejf2179@columbia.edu

Desc: This project is intended to convert aggregate waste values over <1000
(Could be modified for more) cost estimate templates with the same naming
scheme
"""
import pandas as pd
import os

"""
Enter Name Here. nameLeft should be the text before the number. nameRight 
should be the text following it.
"""
nameLeft='05WorkCostBreakdownforOMB ('
nameRight=')'
data=[]
for num in range(1000):
    try:
        df = pd.read_excel(nameLeft+str(num)+nameRight+'.xls')
        agg = []
        agg.append(nameLeft+str(num)+nameRight+'.xls')
        for row in range(23):
            agg.append(df.loc[row+1][14])
        data.append(agg)
    except:
        continue
df = pd.DataFrame(data, columns=['File','Steel','Masonry', 'Concrete', 'Gypsum','Metal', 'Wood',
                                 'Stone','Asphalt','Insulation','Plaster','Tile',
                                 'Linoleum','Carpet','Glass','Plastic','Fiberglass',
                                 'Vinyl','Aluminum','Copper','Rubber','Sand','Bituminous',
                                 'Total'])
df.to_excel('TOTAL'+nameLeft+nameRight+'.xlsx',index=False)