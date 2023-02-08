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



def getAcompanhamentoRetorno(xls, ref):
    t = pd.read_excel(xls, sheet_name='ACOMPANHAMENTO', header=None, index_col=None, na_values='n/a')
    indexFilter = [i for i in range(25)]
    lista2 = [i for i in range(6) if i not in indexFilter]
    e = t.iloc[lista2,:]

    d2 = pd.DataFrame(e.iloc[:,1]).transpose()
    d2.columns = e.iloc[:,0]
    d2.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
    d2['id'] = ref
    return d2



def getAcompanhamentoPrimeiraConsulta(xls, ref):
    t = pd.read_excel(xls, sheet_name='ACOMPANHAMENTO', header=None, index_col=None, na_values='n/a')
    indexFilter = [0, 1]
    lista2 = [i for i in range(6) if i not in indexFilter]
    e = t.iloc[lista2,:]

    d2 = pd.DataFrame(e.iloc[:,1]).transpose()
    d2.columns = e.iloc[:,0]
    d2.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
    d2['id'] = ref
    return d2

def getAdesao(xls, ref):
    t = pd.read_excel(xls, sheet_name='ADESÃO', header=None, index_col=None, na_values='n/a')
    columns=t.iloc[0,:13]
    df = pd.DataFrame(columns=columns)
    
    i=1
    while(1):
        if i<t.shape[0] and isinstance(t.iloc[i,0], datetime.date):
            d2 = pd.DataFrame(t.iloc[i,:13]).transpose()
            d2.columns = df.columns
            df = df.append(d2,ignore_index=False)
            i=i+1
        else:
            break
    df.reset_index(inplace=True, drop=True)
    df['id'] = ref
    return df


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

def getAnaliseNutricional(xls, ref):
    t = pd.read_excel(xls, sheet_name='ANALISE NUTRICIONAL', header=None, index_col=None, na_values='n/a')

    dados = dict()
    dados['ProteinaQuantidade']     = t.iloc[109,2]
    dados['CarboidratoQuantidade']  = t.iloc[109,3]
    dados['LipidioQuantidade']      = t.iloc[109,4]
    dados['ProteinaGKG']            = t.iloc[110,2]
    dados['CarboidratoGKG']         = t.iloc[110,3]
    dados['LipidioGKG']             = t.iloc[110,4]
    dados['ProteinaCalorias']       = t.iloc[111,2]
    dados['CarboidratoCalorias']    = t.iloc[111,3]
    dados['LipidioCalorias']        = t.iloc[111,4]
    dados['ProteinaPercentagem']    = t.iloc[112,2]
    dados['CarboidratoPercentagem'] = t.iloc[112,3]
    dados['LipidioPercentagem']     = t.iloc[112,4]
    dados['CaloriasTotais']         = t.iloc[113,2]
    dados['KCalKG']                 = t.iloc[114,2]
    dados['IngestaoRecomendada']    = t.iloc[116,2]

    d2 = pd.DataFrame(data=dados, index=[0])
    d2.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
    d2['id'] = ref
    return d2

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
    
def getMedicamentos(xls, ref):
    t = pd.read_excel(xls, sheet_name='MEDICAMENTOS', header=None, index_col=None, na_values='n/a')
    columns=t.iloc[1,:6]
    columns=columns.append(pd.Series('Tipo'),ignore_index=True)
    columns=columns.append(pd.Series('ValorTipo'),ignore_index=True)
    df = pd.DataFrame(columns=columns)
    
    blocos = ['Hipertensao', 'Diabetes', 'Insulina', 'Dislipidemia', 'Outra', 'Bicabornato', 'Suplemento']
    posBlocos = [2, 14, 26, 38, 51, 63, 75]
    limite = t.shape[0]
    for p in range(len(blocos)):
        valortipo = t.iloc[posBlocos[p]-2,1]
        for cont in range(10):
            i = cont+posBlocos[p]
            if (i<t.shape[0] and t.iloc[i].notna().any()):
                d2 = pd.DataFrame(t.iloc[i,:6]).transpose()
                d2['Tipo'] = blocos[p]
                d2['ValorTipo'] = valortipo
                d2.columns = df.columns
                df = df.append(d2,ignore_index=False)
                
    df.reset_index(inplace=True, drop=True)
    df['id'] = ref
    return df
    
def getProntuario(xls, ref):
    t = pd.read_excel(xls, sheet_name='PRONTUARIO', header=None, index_col=None, na_values='n/a')
    indexFilter = [0, 9, 10, 11, 12, 13, 14, 32 ]
    lista2 = [i for i in range(38) if i not in indexFilter]
    e = t.iloc[lista2,:]

    d2 = pd.DataFrame(e.iloc[:,1]).transpose()
    d2.columns = e.iloc[:,0]
    d2['id'] = ref
    return d2

def getData(files, hasAdesao):
    dfAntro = dfBioq = dfPront = dfAna = dfMed = dfAnaN = pd.DataFrame()
    f = open("log.txt", "w")

    for i in range(len(files)):
        print('Reading file (' + str(i+1) + '/' + str(len(files))+'): ' + files[i])
        try:
            xls = pd.ExcelFile(files[i])
        except Exception as e:
            msg = ">>> [ERRO] " + str(e) + ": " + files[i] + "\n"
            print(msg)
            f.write(msg)
        
        try:
            dfAntroAux  = getAntropometria(xls, i)         
            dfBioqAux   = getBioquimica(xls, i)
            dfProntAux  = getProntuario(xls, i)
            dfAnaAux    = getAnamnese(xls, i)
            dfMedAux    = getMedicamentos(xls, i)
            dfAnaNAux   = getAnaliseNutricional(xls, i)
            dfAcom1CAux = getAcompanhamentoPrimeiraConsulta(xls, i)
            dfAcompaAux   = getAcompanhamentoRetorno(xls, i)
            if hasAdesao:
                dfAdesaoAux = getAdesao(xls, i)
        except:
            msg = '>>> [ERRO] Erro na coleta do arquivo' + files[i] + "\n"
            print(msg)
            f.write(msg)
            dfAntroAux  = dfAntroAux[0:0]
            dfBioqAux   = dfBioqAux[0:0]
            dfProntAux  = dfProntAux[0:0]
            dfAnaAux    = dfAnaAux[0:0]
            dfMedAux    = dfMedAux[0:0]
            dfAnaNAux   = dfAnaNAux[0:0]
            dfAcom1CAux = dfAcom1CAux[0:0]
            dfAcompaAux = dfAcompaAux[0:0]
            if hasAdesao:
                dfAdesaoAux = dfAdesaoAux[0:0]

        if i == 0:
            dfAntro  = dfAntroAux
            dfBioq   = dfBioqAux
            dfPront  = dfProntAux
            dfAna    = dfAnaAux
            dfMed    = dfMedAux
            dfAnaN   = dfAnaNAux
            dfAcom1C = dfAcom1CAux
            dfAcompR = dfAcompaAux
            if hasAdesao:
                dfAdesao = dfAdesaoAux
        else:
            dfAntro  = dfAntro.append(dfAntroAux, ignore_index=True)
            dfBioq   = dfBioq.append(dfBioqAux, ignore_index=True)
            dfPront  = dfPront.append(dfProntAux, ignore_index=True)
            dfAna    = dfAna.append(dfAnaAux, ignore_index=True)
            dfMed    = dfMed.append(dfMedAux, ignore_index=True)
            dfAnaN   = dfAnaN.append(dfAnaNAux, ignore_index=True)
            dfAcom1C = dfAcom1C.append(dfAcom1CAux, ignore_index=True)
            dfAcompR = dfAcompR.append(dfAcompaAux, ignore_index=True)
            if hasAdesao:
                dfAdesao = dfAdesao.append(dfAdesaoAux, ignore_index=True)
    
    f.close()
    if hasAdesao:
        return dfAntro, dfBioq, dfPront, dfAna, dfMed, dfAnaN, dfAcom1C, dfAcompR, dfAdesao
    else:
        return dfAntro, dfBioq, dfPront, dfAna, dfMed, dfAnaN, dfAcom1C, dfAcompR
            


'''
Análise nutricional foi implementado mas tem planilhas que não possuem essa aba.

Falta:
 - Acompanhamento
 - Acompanhamento (retornos)
'''
def createDataset(hasAdesao = True):
    files = glob.glob("files/**/*.xls*", recursive=True)
    files.sort(key=os.path.abspath)
    if hasAdesao:
        d1, d2, d3, d4, d5, d6, d7, d8, d9 = getData(files, hasAdesao)
    else:
        d1, d2, d3, d4, d5, d6, d7, d8 = getData(files, hasAdesao)
    writer = pd.ExcelWriter('dataset.xlsx')
    d3.to_excel(writer,'Prontuario')
    d1.to_excel(writer,'Antropometria')
    d2.to_excel(writer,'Bioquimica')
    d4.to_excel(writer,'Anamnese')
    d5.to_excel(writer,'Medicamento')
    d6.to_excel(writer,'AnaliseNutricional')
    d7.to_excel(writer,'AcompanhamentoPrimeiraConsulta')
    d8.to_excel(writer,"AcompanhamentoRetorno")
    if hasAdesao:
        d9.to_excel(writer,'Adesao')
    # data.fillna() or similar.
    writer.save()
    print('Banco de dados criado com sucesso! ')


createDataset(False)

