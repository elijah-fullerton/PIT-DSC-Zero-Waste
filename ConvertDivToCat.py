# -*- coding: utf-8 -*-
"""
Created on Tue Jul  5 12:23:58 2022

@author: efull

Name: Elijah Journey Fullerton
email: ejf2179@columbia.edu

Desc: This project is intended to convert JOC Catalogue Converted Excel Sheets
into Catalogues for C&D Waste Estimation
"""

import pandas as pd
import os

#Check Current Directory
#print(os.listdir(os.getcwd()))

#Open Current Division
df = pd.read_excel('CTC - NYC HA, GC - Owner_Part3.xlsx')
#print(df)

#Convert to CSV
df.to_csv('temp.csv',index=False)
df = pd.read_csv('temp.csv')
#print(df.loc[4][0])
#print(str(df.loc[0][0]))

#Initialization
idx = 3000000
unit = None
lineItem = None
data=[]
#Create Units set
units_set=set()
units_set.add('CF')
units_set.add('CY')
units_set.add('SF')
units_set.add('SY')
square_set=set()
square_set.add('SF')
square_set.add('SY')
#Creatunits_set=set()
num_set=set()
num_set.add('0')
num_set.add('1')
num_set.add('2')
num_set.add('3')
num_set.add('4')
num_set.add('5')
num_set.add('6')
num_set.add('7')
num_set.add('8')
num_set.add('9')
num_set.add('.')
#Create Materials Set
materials_df = pd.read_excel('JOC_Div_Materials.xlsx')
materials_set = set()
weights_df= pd.read_excel('Weights Table.xlsx')
for row in range(len(materials_df.index)):
    materials_set.add(materials_df.loc[row][0])
curMaterial='MATERIAL'
#Create Hazardous Materials Set
hazards_set = set()
for row in range(4):
    hazards_set.add(materials_df.loc[row][1])
cur_hazardous=False
#Break up cell by line
for row in range(len(df.index)):
    curCell=(str.splitlines(str(df.loc[row][0])))
    #For each line, break up substrings by double space
    for line in curCell:
        curSubStrSet=line.split('  ')
        #For each substring, if substring contains '...', do more
        lineIdx=-1
        for curSubStrIdx in range(len(curSubStrSet)):
            if curSubStrSet[curSubStrIdx]!='':
                lineIdx+=1
                if lineIdx==2:
                    #Set Unit
                    unitIdx=curSubStrIdx-1
                    while unit == None:
                        if curSubStrSet[unitIdx]!='':
                            unit=curSubStrSet[unitIdx].strip()
                        else:
                            unitIdx-=1
                    if unit!=False:
                        #Set Line Item
                        lineItem = curSubStrSet[curSubStrIdx].split('...')[0].strip()
                        #Set Material
                        for material in materials_set:
                            if material in lineItem.lower():
                                curMaterial = material
                        #Set Weight
                        weight='WEIGHT'
                        if curMaterial!='MATERIAL':
                            weightRow=None
                            weightColumn=None
                            for row in range(28):
                                if curMaterial==str(weights_df.loc[row][0]):
                                    weightRow=row
                                    for col in range(6):
                                        if unit==str(weights_df.loc[0][col]):
                                            weightColumn=col
                                            if unit in square_set:
                                                depth=0
                                                whole=0
                                                if '"' in lineItem:
#Set line items such as '24" x 48", Perforated, Painted Aluminum, Acoustical 
#Lay-In Metal Ceiling Panel (USG Panz™)' to a default depth of 1
                                                    if ' x ' in lineItem:
                                                        depth=1
                                                    elif '/' in lineItem:
 #What we are trying to do in the following lines is pick up depth values 
 #formatted as #-#/#, while also avoiding issues when words are written as
 #word-word.
                                                        if '-' in lineItem:
                                                            try:
                                                                whole=int(lineItem[lineItem.index('-')-1])
                                                            except:
                                                                continue
                                                        try:
                                                            num=int(lineItem[lineItem.index('/')-1])
                                                            dem=int(lineItem[lineItem.index('/')+1])
                                                            depth=whole+float(num/dem)
                                                        except:
                                                            left = lineItem.index('"')-1
                                                            right = left
                                                            while left>0 and lineItem[left-1] in num_set:
                                                                left-=1
                                                            try: 
                                                                depth=float(lineItem[left:right+1])
                                                            except:
                                                                continue
                                                    else:
                                                        left = lineItem.index('"')-1
                                                        right = left
                                                        while left>0 and lineItem[left-1] in num_set:
                                                            left-=1
                                                        try: 
                                                            depth=float(lineItem[left:right+1])
                                                        except:
                                                            continue
                                                        #Discarded
                                                        #try:
                                                        #    if int(lineItem[lineItem.index('"')-2]) in range(11):
                                                        #        depth=10*int(lineItem[lineItem.index('"')-2])+int(lineItem[lineItem.index('"')-1])
                                                        #except:
                                                        #    try:
                                                        #        depth=int(lineItem[lineItem.index('"')-1])
                                                        #    except:
                                                        #        continue
                            if weightColumn != None and weightRow != None:
                                weight = weights_df.loc[weightRow][weightColumn]
                                if unit in square_set:
                                    weight=weight*depth
                        #Check Hazardous
                        for hazard in hazards_set:
                            if hazard in lineItem.lower():
                                cur_hazardous=True
                        #Write to Data
                        if unit in units_set and curMaterial in materials_set:
                            #if unit=='SF' and curMaterial=='concrete' and weight==0:
                                #print(float('1.'+str(idx)),unit,lineItem,curMaterial,weight,depth)
                            data.append([idx,unit,lineItem,curMaterial,weight,cur_hazardous])
                        #Re-Initialize
                        curMaterial='MATERIAL'
                        cur_hazardous=False
                        unit=None
                        idx+=1
#Write to Output
df = pd.DataFrame(data, columns=['Index','Unit', 'Line Item', 'Material','Weight','Hazardous'])
df.to_excel('trueOut.xlsx',index=False)