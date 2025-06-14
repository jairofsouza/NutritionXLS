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
import traceback
import sys

def getPressao(xls, ref):
    try: 
        t = pd.read_excel(xls, sheet_name='CONTROLE PA', header=None, index_col=None, na_values='n/a')
        df = pd.DataFrame(columns=t.iloc[0,:2])
        
        i=1
        while(1):
            if i < t.shape[0] and isinstance(t.iloc[i,0], datetime.date):
                d2 = pd.DataFrame(t.iloc[i,:2]).transpose()
                d2.columns = df.columns
                df = df.append(d2,ignore_index=False)
                i=i+1
            else:
                break
        df.reset_index(inplace=True, drop=True)
        df['id'] = ref

    except:
        raise Exception("CONTROLE PA")

    return df



def getAcompanhamentoRetorno(xls, ref):
    try:
        t = pd.read_excel(xls, sheet_name='ACOMPANHAMENTO', header=None, index_col=None, na_values='n/a')
        indexFilter = [i for i in range(25)]
    
        inicio=2
        salto=22
        fim=15
        qColunas=9

        box = [(inicio+(i*salto),inicio+fim+(i*salto)) for i in range(5)]
        locais = [i*3 for i in range(qColunas)]

        total = pd.DataFrame(columns=t.iloc[inicio+salto:inicio+salto+fim,0])
        for pos in locais:
            for (inicio,fim) in box:
                if pos+inicio != 2 and pos+1 < t.shape[1] and isinstance(t.iloc[inicio, pos+1], datetime.date):
                    d1 = pd.DataFrame(t.iloc[inicio:fim,pos+1]).transpose().rename(columns=t.iloc[inicio:fim,pos])
                    total = pd.concat([total, d1],ignore_index=True)

        if not total.empty:
            total['id'] = ref

    except:
        raise Exception("ACOMPANHAMENTO RETORNO")


    return total


def getAcompanhamentoPrimeiraConsulta(xls, ref):
    try:
        t = pd.read_excel(xls, sheet_name='ACOMPANHAMENTO', header=None, index_col=None, na_values='n/a')
        indexFilter = [0, 1]
        lista2 = [i for i in range(6) if i not in indexFilter]
        e = t.iloc[lista2,:]

        d2 = pd.DataFrame(e.iloc[:,1]).transpose()
        d2.columns = e.iloc[:,0]
        d2.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
        d2['id'] = ref
    except:
        raise Exception("ACOMPANHAMENTO")
    return d2

def getAdesao(xls, ref):
    try:
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
    except:
        raise Exception("ADESAO")
    return df


def getAntropometria(xls, ref):
    try:
        t = pd.read_excel(xls, sheet_name='ANTROPOMETRIA', header=None, index_col=None, na_values='n/a')
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
    except:
        raise Exception("ANTROPOMETRIA")

    return df

def getAnaliseNutricional(xls, ref):
    try:
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
    except:
        raise Exception("ANALISE NUTRICIONAL")
    return d2

def getAnamnese(xls, ref):
    try:
        t = pd.read_excel(xls, sheet_name='ANAMNESE ALIMENTAR', header=None, index_col=None, na_values='n/a')
        indexFilter = [0, 1, 15]
        lista2 = [i for i in range(t.shape[0]) if i not in indexFilter]
        e = t.iloc[lista2,:]

        d2 = pd.DataFrame(e.iloc[:,1]).transpose()
        d2.columns = e.iloc[:,0]
        d2.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
        d2['id'] = ref
    except:
        raise Exception("ANAMNESE ALIMENTAR")
    return d2

def getBioquimica(xls, ref):
    try:
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
    except:
        raise Exception("BIOQUIMICA")
    return df
    
def getMedicamentos(xls, ref):
    try:
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
    except:
        raise Exception("MEDICAMENTOS")
    return df
    
def getProntuario(xls, ref):
    try:
        t = pd.read_excel(xls, sheet_name='PRONTUARIO', header=None, index_col=None, na_values='n/a')
        indexFilter = [0, 9, 10, 11, 12, 13, 14, 32 ]
        lista2 = [i for i in range(38) if i not in indexFilter]
        e = t.iloc[lista2,:]

        d2 = pd.DataFrame(e.iloc[:,1]).transpose()
        d2.columns = e.iloc[:,0]
        d2['id'] = ref
    except:
        raise Exception("PRONTUARIO")
    return d2

def getData(files, hasAdesao):
    dfAntro = dfBioq = dfPront = dfAna = dfMed = dfAnaN = dfAcom1C = dfAcompR = dfAdesao = dfPressao = pd.DataFrame()
    lista = ["Antropometria", "Bioquimica", "Prontuario", "Anamnese", "Medicamentos", "AnaliseNutricional", "Acompanhamento (PrimeiraConsulta)", "Acompanhamento (Retornos)", "Adesao", "Controle PA"]

    f = open("log.txt", "w")

    for i in range(len(files)):
        print('Reading file (' + str(i+1) + '/' + str(len(files))+'): ' + files[i])
        try:
            xls = pd.ExcelFile(files[i])
        except Exception as e:
            msg = ">>> [ERRO] " + str(e) + ": " + files[i] + "\n"
            print(msg)
            f.write(msg)
            continue
        
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
                dfPressaoAux   = getPressao(xls, i)
        except Exception as e:
            print(traceback.format_exc())
            msg = '>>> [ERRO] Erro na coleta de dados: ' + files[i] + " (" + str(e) +  ")\n"
            print(msg)
            f.write(msg)
            continue

        try:
            line = 271 #guardar aqui esse número de linha!
            dfAntro  = dfAntroAux  if dfAntro.empty  else dfAntro.append(dfAntroAux, ignore_index=True)
            dfBioq   = dfBioqAux   if dfBioq.empty   else dfBioq.append(dfBioqAux, ignore_index=True)
            dfPront  = dfProntAux  if dfPront.empty  else dfPront.append(dfProntAux, ignore_index=True)
            dfAna    = dfAnaAux    if dfAna.empty    else dfAna.append(dfAnaAux, ignore_index=True)
            dfMed    = dfMedAux    if dfMed.empty    else dfMed.append(dfMedAux, ignore_index=True)
            dfAnaN   = dfAnaNAux   if dfAnaN.empty   else dfAnaN.append(dfAnaNAux, ignore_index=True)
            dfAcom1C = dfAcom1CAux if dfAcom1C.empty else dfAcom1C.append(dfAcom1CAux, ignore_index=True)
            dfAcompR = dfAcompaAux if dfAcompR.empty else dfAcompR.append(dfAcompaAux, ignore_index=True)
            if hasAdesao:
                dfAdesao = dfAdesaoAux   if dfAdesao.empty  else dfAdesao.append(dfAdesaoAux, ignore_index=True)
                dfPressao = dfPressaoAux if dfPressao.empty else dfPressao.append(dfPressaoAux, ignore_index=True)
        except Exception as e:
            print(traceback.format_exc())
            msg = '>>> [ERRO] Erro na coleta de dados (COLUNAS NÃO BATEM): ' + files[i] + " (" + str(lista[e.__traceback__.tb_lineno-line-1]) +  ")\n"
            print(msg)
            f.write(msg)
            continue
    
    f.close()
    if hasAdesao:
        return dfAntro, dfBioq, dfPront, dfAna, dfMed, dfAnaN, dfAcom1C, dfAcompR, dfAdesao, dfPressao
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
        d1, d2, d3, d4, d5, d6, d7, d8, d9, d10 = getData(files, hasAdesao)
    else:
        d1, d2, d3, d4, d5, d6, d7, d8 = getData(files, hasAdesao)
    writer = pd.ExcelWriter('dataset.xlsx')
    d3.to_excel(writer,sheet_name='Prontuario')
    d1.to_excel(writer,sheet_name='Antropometria')
    d2.to_excel(writer,sheet_name='Bioquimica')
    d4.to_excel(writer,sheet_name='Anamnese')
    d5.to_excel(writer,sheet_name='Medicamento')
    d6.to_excel(writer,sheet_name='AnaliseNutricional')
    d7.to_excel(writer,sheet_name='AcompanhamentoPrimeiraConsulta')
    d8.to_excel(writer,sheet_name='AcompanhamentoRetorno')
    if hasAdesao:
        d9.to_excel(writer,sheet_name='Adesao')
        d10.to_excel(writer, sheet_name='Pressao')
    # data.fillna() or similar.
    writer.close()
    print('Banco de dados criado com sucesso! ')


createDataset(True)

