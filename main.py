#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 14:40:58 2021

@author: jairosouza
"""


import pandas as pd
import datetime
import glob
import os



def getAntropometria(xls, ref):
    t = pd.read_excel(xls, sheet_name='ANTROPOMETRIA', header=None, index_col=None, na_values='n/a')
    #print(t.iloc[2,:22])
    df = pd.DataFrame(columns=t.iloc[2,:22])
    
    i=4
    while(1):
        if isinstance(t.iloc[i,0], datetime.date):
            d2 = pd.DataFrame(t.iloc[i,:22]).transpose()
            d2.columns = df.columns
            df = df.append(d2,ignore_index=False)
            i=i+1
        else:
            break
    df.reset_index(inplace=True, drop=True)
    df['id'] = ref
    return df

def getAnamnese(xls, ref):
    t = pd.read_excel(xls, sheet_name='ANAMNESE ALIMENTAR', header=None, index_col=None, na_values='n/a')
    indexFilter = [0, 1, 15]
    lista2 = [i for i in range(t.shape[0]) if i not in indexFilter]
    e = t.iloc[lista2,:]

    d2 = pd.DataFrame(e.iloc[:,1]).transpose()
    d2.columns = e.iloc[:,0]
    d2.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
    d2['id'] = ref
    return d2

def getBioquimica(xls, ref):
    t = pd.read_excel(xls, sheet_name='BIOQUIMICA', header=None, index_col=None, na_values='n/a')
    columns=t.iloc[2,:45]
    columns[1] = 'UreiaSangue'
    df = pd.DataFrame(columns=columns)
    
    i=6
    while(1):
        if i<t.shape[0] and isinstance(t.iloc[i,0], datetime.date):
            d2 = pd.DataFrame(t.iloc[i,:45]).transpose()
            d2.columns = df.columns
            df = df.append(d2,ignore_index=False)
            i=i+1
        else:
            break
    df.reset_index(inplace=True, drop=True)
    df['id'] = ref
    return df
    
def getProntuario(xls, ref):
    t = pd.read_excel(xls, sheet_name='PRONTUARIO', header=None, index_col=None, na_values='n/a')
    indexFilter = [0, 9, 10, 11, 12, 13, 14, 20, 21, 22, 32 ]
    lista2 = [i for i in range(t.shape[0]) if i not in indexFilter]
    e = t.iloc[lista2,:]

    d2 = pd.DataFrame(e.iloc[:,1]).transpose()
    d2.columns = e.iloc[:,0]
    d2['id'] = ref
    return d2

def getData(files):
    dfAntro = dfBioq = dfPront = pd.DataFrame()
    for i in range(len(files)):
        print('Reading file (' + str(i+1) + '/' + str(len(files))+'): ' + files[i])
        xls = pd.ExcelFile(files[i])
        dfAntroAux = getAntropometria(xls, i)
        dfBioqAux  = getBioquimica(xls, i)
        dfProntAux = getProntuario(xls, i)
        dfAnaAux   = getAnamnese(xls, i)
        
        if i == 0:
            dfAntro = dfAntroAux
            dfBioq  = dfBioqAux
            dfPront = dfProntAux
            dfAna   = dfAnaAux
        else:
            dfAntro = dfAntro.append(dfAntroAux, ignore_index=True)
            dfBioq  = dfBioq.append(dfBioqAux, ignore_index=True)
            dfPront = dfPront.append(dfProntAux, ignore_index=True)
            dfAna   = dfAna.append(dfAnaAux, ignore_index=True)
    return dfAntro, dfBioq, dfPront, dfAna
            

def createDataset():
    files = glob.glob("files/**/*.xls*", recursive=True)
    files.sort(key=os.path.abspath)
    d1, d2, d3, d4 = getData(files)
    writer = pd.ExcelWriter('dataset.xlsx')
    d3.to_excel(writer,'Prontuario')
    d1.to_excel(writer,'Antropometria')
    d2.to_excel(writer,'Bioquimica')
    d4.to_excel(writer,'Anamnese')
    # data.fillna() or similar.
    writer.save()
    print('Banco de dados criado com sucesso! ')




