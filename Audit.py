import pandas as pd
import numpy as np
import codecs
import time
import os
import sys
import string
import xlsxwriter
import io
import csv
import re
import urllib

from tkinter import filedialog
from tkinter import *
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from functools import reduce



import requests
from lxml import html
import urllib


def loadscript(a,b,c):
    print(a,b,c)

def loadscripts(step, trial, duration):
    print(step, end='', flush=True)
    for i in range(trial): 
        time.sleep(duration)
        print('.', end='', flush=True)
    time.sleep(.2)
    print("\n")

def add_zeros(col, intgr):
	dfObj[col] = dfObj[col].apply(lambda x: x.zfill(intgr))
	return dfObj[col]


files = os.listdir(os.curdir)

writer = pd.ExcelWriter('TrialReport.xlsx', engine='xlsxwriter')

start = time.time()



Tk().withdraw()

PrimeRX = ['primerx','Primerx','PrimeRX','PRIMERX','disp','Disp','DISP','report','Report','REPORT']
if any([i for i in files if any(x in i for x in PrimeRX)]):
    PrimeRXFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Dispense FILE')
    primeToExcel = time.time()
    print('primeToExcel:', primeToExcel - start)
    if "xls" in PrimeRXFile:
        rxRawDF = pd.read_excel(PrimeRXFile, sheet_name=0, header= None, index= False)
        rxRawDF.columns = rxRawDF.iloc[0]
        rxRawDF = rxRawDF.drop(rxRawDF.index[0])
    elif "csv" in PrimeRXFile:
        rxRawDF = pd.read_csv(PrimeRXFile,header= None)
        rxRawDF.columns = rxRawDF.iloc[0]
        rxRawDF = rxRawDF.drop(rxRawDF.index[0])
        rxRawDF = rxRawDF.replace({'=':'', '"':''}, regex=True)

rxRawDF.to_excel(writer, sheet_name='PrimeRx RAW', index=False)
excelwrite = time.time()
print('excel writing: ',  excelwrite - start)
rawColumns = ['NDC','DRGNAME','DRUGNAME','DRUG NAME','DRUG NAME ','DRUGSTRONG','Pack','PACK','PACKAGESIZE','QTY','Quantity','QUANT','Quant','STRENGTH']
rxRawDF = rxRawDF[np.intersect1d(rxRawDF.columns, rawColumns)]
rxRawDF.columns = rxRawDF.columns.str.replace('DRGNAME','DRUG NAME')
rxRawDF.columns = rxRawDF.columns.str.replace('DRUGNAME','DRUG NAME')
rxRawDF.columns = rxRawDF.columns.str.replace('DRUG NAME ','DRUG NAME')
rxRawDF.columns = rxRawDF.columns.str.replace('DRUGSTRONG','STRENGTH')
rxRawDF.columns = rxRawDF.columns.str.replace('PACKAGESIZE','PACK')
rxRawDF.columns = rxRawDF.columns.str.replace('QUANT','QUANTITY')

fixColumns = time.time()
print('Columns Fixed:', fixColumns - start)

rxRawDF[['PACK','QUANTITY']]= rxRawDF[['PACK','QUANTITY']].apply(pd.to_numeric, errors='coerce')
loadscript('Copying over PrimeRX', 6, .4)
reportDF = rxRawDF[['NDC','DRUG NAME','STRENGTH', 'PACK','QUANTITY']].copy()
reportDF['QUANTITY'] = reportDF['QUANTITY'].apply(float)
reportDF['PACK'] = reportDF['PACK'].apply(float)
reportQuantDF = reportDF.groupby(['NDC'], as_index=False).sum()
finishSum = time.time()
print(finishSum - start)
reportQuantDF = reportQuantDF.drop(['PACK'], axis=1)
reportDF = reportDF.drop(['QUANTITY'], axis=1)
reportDF = pd.merge(reportDF, reportQuantDF, on=['NDC'], how='left')
reportDF['DISP'] = reportDF['QUANTITY']/reportDF['PACK']
reportDF['DISP'] = reportDF['DISP'].apply(lambda x:round(x,1))


DrugID = [filename for filename in os.listdir('.') if re.search(r'drug', filename, re.IGNORECASE)] 
if not DrugID:
    pass
else:
    DrugIDFile = DrugID[0]
    DrugIDDF = pd.read_csv(DrugIDFile, encoding='ISO-8859-1')
    #DrugIDDF.columns = DrugIDDF.iloc[0]
    DrugIDDF = DrugIDDF.drop(DrugIDDF.index[0])
    DrugIDDF.to_excel(writer, sheet_name='DrugID', index=False)

if not DrugID:
    reportMultiDF = pd.DataFrame(columns=['NDC']) 
else:
    reportMultiDF0 = reportDF[['NDC','DRUG NAME','STRENGTH','DISP']]
    reportMultiDF = pd.merge(reportMultiDF0, DrugIDDF, on=['NDC'], how='left')
    reportMultiDF['DISP'] = reportMultiDF['DISP']*reportMultiDF['PACK']
    reportMultiDF = reportMultiDF[['NDC','DrugID','PACK','DRUG NAME','STRENGTH','DISP']]


KINRXFileX = [filename for filename in os.listdir('.') if re.search(r'kin*rx', filename, re.IGNORECASE)] 
if not KINRXFileX:
    kinrayRXDF = pd.DataFrame(columns=['NDC'])
    kinrayRXMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    KINRXFile = KINRXFileX[0] 
    if "xls" in KINRXFile:
            kinrayRXDF = pd.read_excel(KINRXFile,header= None)
    elif "csv" in KINRXFile:
            kinrayRXDF = pd.read_csv(KINRXFile,header= None)
    kinrayRXDF = kinrayRXDF[((kinrayRXDF.astype(str) == 'Universal NDC').cumsum()).any(1)]
    kinrayRXDF.columns = kinrayRXDF.iloc[0]
    kinrayRXDF = kinrayRXDF.drop(kinrayRXDF.index[0])
    kinrayRXDF.to_excel(writer, sheet_name='KIN_RX', index=False)
    kinrayRXDF['NDC']=kinrayRXDF['Universal NDC'].astype(str).str[:5]+'-'+kinrayRXDF['Universal NDC'].astype(str).str[5:9]+'-'+kinrayRXDF['Universal NDC'].astype(str).str[-2:]
    kinrayRXDF = kinrayRXDF.groupby(['NDC'], as_index=False).sum()
    kinrayRXDF = kinrayRXDF.rename(columns={'Qty': 'KIN RX'})
    kinrayRXDF = kinrayRXDF[['NDC','KIN RX']]
 
    if not DrugID:
        pass
    else:
        kinrayRXMultiDF = pd.merge(DrugIDDF, kinrayRXDF,  on=['NDC'], how='left')
        kinrayRXMultiDF['KINRX UNITS'] = kinrayRXMultiDF['KIN RX']*kinrayRXMultiDF['PACK']
        kinrayRXMultiDF = kinrayRXMultiDF[['DrugID','KINRX UNITS']]
        kinrayRXMultiDF = kinrayRXMultiDF.groupby(['DrugID'], as_index=False).sum()


KINOTCFileX = [filename for filename in os.listdir('.') if re.search(r'kin*otc', filename, re.IGNORECASE)] 
if not KINOTCFileX:
    kinrayOTCDF = pd.DataFrame(columns=['NDC'])
    kinrayOTCMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    KINOTCFile = KINOTCFileX[0]
    if "xls" in KINOTCFile:
            kinrayOTCDF = pd.read_excel(KINOTCFile,header= None)
    elif "csv" in KINOTCFile:
            kinrayOTCDF = pd.read_csv(KINOTCFile,header= None)
    kinrayOTCDF = kinrayOTCDF[((kinrayOTCDF.astype(str) == 'Universal NDC').cumsum()).any(1)]
    kinrayOTCDF.columns = kinrayOTCDF.iloc[0]
    kinrayOTCDF = kinrayOTCDF.drop(kinrayOTCDF.index[0])
    kinrayOTCDF.to_excel(writer, sheet_name='KIN_OTC', index=False)
    kinrayOTCDF['NDC']=kinrayOTCDF['Universal NDC'].astype(str).str[:5]+'-'+kinrayOTCDF['Universal NDC'].astype(str).str[5:9]+'-'+kinrayOTCDF['Universal NDC'].astype(str).str[-2:]
    kinrayOTCDF = kinrayOTCDF.groupby(['NDC'], as_index=False).sum()
    kinrayOTCDF = kinrayOTCDF.rename(columns={'Qty': 'KIN OTC'})
    kinrayOTCDF = kinrayOTCDF[['NDC','KIN OTC']]

    if not DrugID:
        pass
    else:
        kinrayOTCMultiDF = pd.merge(DrugIDDF, kinrayOTCDF, on=['NDC'], how='left')
        kinrayOTCMultiDF['KINOTC UNITS'] = kinrayOTCMultiDF['KIN OTC']*kinrayOTCMultiDF['PACK']
        kinrayOTCMultiDF = kinrayOTCMultiDF[['DrugID','KINOTC UNITS']]
        kinrayOTCMultiDF = kinrayOTCMultiDF.groupby(['DrugID'], as_index=False).sum()


MCKFileX = [filename for filename in os.listdir('.') if re.search(r'mck|mckesson', filename, re.IGNORECASE)]  
if not MCKFileX:
    MCKDF = pd.DataFrame(columns=['NDC'])
    MCKMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    MCKFile = MCKFileX[0]
    if "xls" in MCKFile:
            MCKDF = pd.read_excel(MCKFile,header= None, index= False)
    elif "csv" in MCKFile:
            MCKDF = pd.read_csv(MCKFile,header= None, encoding='ISO-8859-1')
    MCKDF = MCKDF[((MCKDF.astype(str) == 'NDC/UPC').cumsum()).any(1)]
    MCKDF.columns = MCKDF.iloc[0]
    MCKDF = MCKDF.drop(MCKDF.index[0])
    MCKDF.to_excel(writer, sheet_name='MCK', index=False)
    MCKDF['NDC/UPC']=MCKDF['NDC/UPC'].astype(str).str[:5]+'-'+MCKDF['NDC/UPC'].astype(str).str[5:9]+'-'+MCKDF['NDC/UPC'].astype(str).str[-2:]
    MCKDF = MCKDF[['NDC/UPC','Net']]
    MCKDF = MCKDF.rename(columns={'NDC/UPC': 'NDC'})
    MCKDF = MCKDF.rename(columns={'Net': 'MCK'})
    print(MCKDF.head())
    MCKDF['MCK']=MCKDF['MCK'].apply(float)

    MCKDF = MCKDF.groupby(['NDC'], as_index=False).sum()


    print(MCKDF.head())
    if not DrugID:
        pass
    else:
        MCKMultiDF = pd.merge(DrugIDDF, MCKDF, on=['NDC'], how='left')
        MCKMultiDF['MCK UNITS'] = MCKMultiDF['MCK']*MCKMultiDF['PACK']
        MCKMultiDF = MCKMultiDF[['DrugID','MCK UNITS']]
        MCKMultiDF = MCKMultiDF.groupby(['DrugID'], as_index=False).sum()

TopRXFileX = [filename for filename in os.listdir('.') if re.search(r'top', filename, re.IGNORECASE)]  
if not TopRXFileX:
    TopRXDF = pd.DataFrame(columns=['NDC'])
    TopRXMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    TopRXFile = TopRXFileX[0]
    if "xls" in TopRXFile:
            TopRXDF = pd.read_excel(TopRXFile,header= None, index= False)
    elif "csv" in TopRXFile:
            TopRXDF = pd.read_csv(TopRXFile,header= None, encoding='ISO-8859-1')
    ToPRXDF = ToPRXDF[((ToPRXDF.astype(str) == 'NDC#').cumsum()).any(1)]
    TopRXDF.columns = TopRXDF.iloc[0]
    TopRXDF = TopRXDF.drop(TopRXDF.index[0])
    TopRXDF.to_excel(writer, sheet_name='TOPRX', index=False)
    TopRXDF['NDC']=TopRXDF['NDC#'].astype(str).str[:5]+'-'+TopRXDF['NDC#'].astype(str).str[5:9]+'-'+TopRXDF['NDC#'].astype(str).str[-2:]
    TopRXDF = TopRXDF[['NDC','QUANTITY']]
    TopRXDF['QUANTITY']=TopRXDF['QUANTITY'].apply(float)
    TopRXDF = TopRXDF.groupby(['NDC'], as_index=False).sum()
    TopRXDF = TopRXDF.rename(columns={'QUANTITY': 'TopRX'})
   
    if not DrugID:
        pass
    else:
        TopRXMultiDF = pd.merge(DrugIDDF, TopRXMultiDF, on=['NDC'], how='left')
        TopRXMultiDF['TopRX UNITS'] = TopRXMultiDF['TopRX']*TopRXMultiDF['PACK']
        TopRXMultiDF = TopRXMultiDF[['DrugID','TopRX UNITS']]
        TopMultiDF = TopMultiDF.groupby(['DrugID'], as_index=False).sum()
   



ABCFileX = [filename for filename in os.listdir('.') if re.search(r'abc|amerisource', filename, re.IGNORECASE)] 
if not ABCFileX:
    AmerisourceDF = pd.DataFrame(columns=['NDC'])
    AmerisourceMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    ABCFile = ABCFileX[0]
    if "xls" in ABCFile:
        AmerisourceDF = pd.read_excel(ABCFile,header= None, index= False, sheet_name='Item Level Detail')
        AmerisourceDF = AmerisourceDF[((AmerisourceDF.astype(str) == 'NDC').cumsum()).any(1)]
        AmerisourceDF.columns = AmerisourceDF.iloc[0]
        AmerisourceDF = AmerisourceDF.drop(AmerisourceDF.index[0])
        AmerisourceDF.to_excel(writer, sheet_name='ABC', index=False)
        AmerisourceDF['NDC']=AmerisourceDF['NDC'].astype(str).str[:5]+'-'+AmerisourceDF['NDC'].astype(str).str[5:9]+'-'+AmerisourceDF['NDC'].astype(str).str[-2:]
        AmerisourceDF = AmerisourceDF[['NDC','Sales Less Credits Qty']]
        AmerisourceDF['Sales Less Credits Qty']=AmerisourceDF['Sales Less Credits Qty'].apply(float)
        AmerisourceDF = AmerisourceDF.groupby(['NDC'], as_index=False).sum()
        AmerisourceDF = AmerisourceDF.rename(columns={'Sales Less Credits Qty': 'ABC'})

        if not DrugID:
            pass
        else:
         
            AmerisourceMultiDF = pd.merge(DrugIDDF, AmerisourceMultiDF, on=['NDC'], how='left')
            AmerisourceMultiDF['ABC UNITS'] = AmerisourceMultiDF['Amerisource']*AmerisourceMultiDF['PACK']
            AmerisourceMultiDF = AmerisourceMultiDF[['DrugID','ABC UNITS']]
            AmerisourceMultiDF = AmerisourceMultiDF.groupby(['DrugID'], as_index=False).sum()

    elif "csv" in ABCFile:
        AmerisourceDF = pd.read_csv('ABC.csv', sep='\t')
        AmerisourceDF.to_excel(writer, sheet_name='ABC', index=False)
        AmerisourceDF = AmerisourceDF[((AmerisourceDF.astype(str) == 'NDC').cumsum()).any(1)]
        AmerisourceDF.columns = AmerisourceDF.iloc[0]
        AmerisourceDF = AmerisourceDF.drop(AmerisourceDF.index[0])
        AmerisourceDF['NDC']=AmerisourceDF['NDC'].astype(str).str[:5]+'-'+AmerisourceDF['NDC'].astype(str).str[5:9]+'-'+AmerisourceDF['NDC'].astype(str).str[-2:]
        AmerisourceDF = AmerisourceDF[['NDC','Shipped Qty']]
        AmerisourceDF['Shipped Qty']=AmerisourceDF['Shipped Qty'].apply(float)
        AmerisourceDF = AmerisourceDF.groupby(['NDC'], as_index=False).sum()
        AmerisourceDF = AmerisourceDF.rename(columns={'Shipped Qty': 'ABC'})
    
        if not DrugID:
            pass
        else:
         
            AmerisourceMultiDF = pd.merge(DrugIDDF, AmerisourceMultiDF, on=['NDC'], how='left')
            AmerisourceMultiDF['ABC UNITS'] = AmerisourceMultiDF['Amerisource']*AmerisourceMultiDF['PACK']
            AmerisourceMultiDF = AmerisourceMultiDF[['DrugID','ABC UNITS']]
            AmerisourceMultiDF = AmerisourceMultiDF.groupby(['DrugID'], as_index=False).sum()



OakFileX = [filename for filename in os.listdir('.') if re.search(r'oak', filename, re.IGNORECASE)] 
if not OakFileX:
    OakDF = pd.DataFrame(columns=['NDC'])
    OakMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    OakFile = OakFileX[0]
    if "xls" in OakFile:
            OakDF = pd.read_excel(OakFile,header= None, index= False)
    elif "csv" in OakFile:
            OakDF = pd.read_csv(OakFile, encoding='ISO-8859-1')
    OakDF = OakDF[((OakDF.astype(str) == 'NDC').cumsum()).any(1)]
    OakDF.columns = OakDF.iloc[0]
    OakDF = OakDF.drop(OakDF.index[0])
    OakDF.to_excel(writer, sheet_name='OAK', index=False)
    OakDF['NDC']=OakDF['NDC'].astype(str).str[:5]+'-'+OakDF['NDC'].astype(str).str[5:9]+'-'+OakDF['NDC'].astype(str).str[-2:]
    OakDF = OakDF[['NDC','Quantity']]
    OakDF['Quantity']=OakDF['Quantity'].apply(float)
    OakDF = OakDF.groupby(['NDC'], as_index=False).sum()
    OakDF = OakDF.rename(columns={'Quantity': 'OAK'})

    if not DrugID:
        pass
    else: 
        OakMultiDF0 = OakDF
        OakMultiDF = pd.merge(OakMultiDF0, reportMultiDF, on=['NDC'], how='left')
        OakMultiDF['OAK UNITS'] = OakMultiDF['OAK']*OakMultiDF['PACK']
        OakMultiDF = OakMultiDF[['DrugID','OAK UNITS']]
        OakMultiDF = OakMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)




MaksFileX = [filename for filename in os.listdir('.') if re.search(r'maks', filename, re.IGNORECASE)]  
if not MaksFileX:
    MaksDF = pd.DataFrame(columns=['NDC'])
    MaksMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    MaksFile = MaksFileX[0]
    if "xls" in MaksFile:
            MaksDF = pd.read_excel(MaksFile,header= None, index= False)
            MaksDF = MaksDF.rename(columns={'SOLD QTY':'Quantity'})
    elif "csv" in MaksFile:
            MaksDF = pd.read_csv(MaksFile, encoding='ISO-8859-1')
    MaksDF = MaksDF[((MaksDF.astype(str) == 'NDC/UPC').cumsum()).any(1)]
    MaksDF.columns = MaksDF.iloc[0]
    MaksDF = MaksDF.drop(MaksDF.index[0])
    MaksDF.to_excel(writer, sheet_name='MAKS', index=False)
    MaksDF['NDC']=MaksDF['NDC/UPC'].astype(str).str[:5]+'-'+MaksDF['NDC/UPC'].astype(str).str[5:9]+'-'+MaksDF['NDC/UPC'].astype(str).str[-2:]
    MaksDF = MaksDF[['NDC','Quantity']]
    MaksDF['Quantity']=MaksDF['Quantity'].apply(float)
    MaksDF = MaksDF.groupby(['NDC'], as_index=False).sum()
    MaksDF = MaksDF.rename(columns={'Quantity': 'MAKS'})
    
    if not DrugID:
        pass
    else: 
        MaksMultiDF0 = MaksDF
        MaksMultiDF = pd.merge(MaksMultiDF0, reportMultiDF, on=['NDC'], how='left')
        MaksMultiDF['Maks UNITS'] = MaksMultiDF['Maks']*MaksMultiDF['PACK']
        MaksMultiDF = MaksMultiDF[['DrugID','Maks UNITS']]
        MaksMultiDF = MaksMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)




AlpineFileX = [filename for filename in os.listdir('.') if re.search(r'alpine', filename, re.IGNORECASE)] 
if not AlpineFileX:
    AlpineDF = pd.DataFrame(columns=['NDC'])
    AlpineMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    AlpineFile = AlpineFileX[0]
    if "xls" in AlpineFile:
            AlpineDF = pd.read_excel(AlpineFile,header= None, index= False)
    elif "csv" in AlpineFile:
            AlpineDF = pd.read_csv(AlpineFile, encoding='ISO-8859-1')
    AlpineDF = AlpineDF[((AlpineDF.astype(str) == 'NDC').cumsum()).any(1)]
    AlpineDF.columns = AlpineDF.iloc[0]
    AlpineDF = AlpineDF.drop(AlpineDF.index[0])
    AlpineDF.to_excel(writer, sheet_name='Alpine', index=False)
    AlpineDF['NDC']=AlpineDF['NDC'].astype(str).str[:5]+'-'+AlpineDF['NDC'].astype(str).str[5:9]+'-'+AlpineDF['NDC'].astype(str).str[-2:]
    AlpineDF = AlpineDF[['NDC','Quantity']]
    AlpineDF['Quantity']=AlpineDF['Quantity'].apply(float)
    AlpineDF = AlpineDF.groupby(['NDC'], as_index=False).sum()
    AlpineDF = AlpineDF.rename(columns={'Quantity': 'ALPINE'})

    if not DrugID:
        pass
    else: 
        AlpineMultiDF0 = AlpineDF
        AlpineMultiDF = pd.merge(AlpineMultiDF0, reportMultiDF, on=['NDC'], how='left')
        AlpineMultiDF['ALPINE UNITS'] = AlpineMultiDF['ALPINE']*AlpineMultiDF['PACK']
        AlpineMultiDF = AlpineMultiDF[['DrugID','ALPINE UNITS']]
        AlpineMultiDF = AlpineMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)



HDSmithFileX = [filename for filename in os.listdir('.') if re.search(r'hds', filename, re.IGNORECASE)] 
if not HDSmithFileX:
    HDSmithDF = pd.DataFrame(columns=['NDC'])
    HDSmithMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    HDSmithFile = HDSmithFileX[0]
    if "xls" in HDSmithFile:
            HDSmithDF = pd.read_excel(HDSmithFile,header= None, index= False)
            
    elif "csv" in HDSmithFile:
            HDSmithDF = pd.read_csv(HDSmithFile, encoding='ISO-8859-1')
    HDSmithDF = HDSmithDF[((HDSmithDF.astype(str) == 'NDC').cumsum()).any(1)]
    HDSmithDF.columns = HDSmithDF.iloc[0]
    HDSmithDF = HDSmithDF.drop(HDSmithDF.index[0])
    HDSmithDF.to_excel(writer, sheet_name='HDSmith', index=False)
    HDSmithDF = HDSmithDF[['NDC','Units']]
    HDSmithDF['Units']=HDSmithDF['Units'].apply(float)
    HDSmithDF = HDSmithDF.groupby(['NDC'], as_index=False).sum()
    HDSmithDF = HDSmithDF.rename(columns={'Units': 'HDSMITH'})

    if not DrugID:
        pass
    else: 
        HDSmithMultiDF0 = HDSmithDF
        HDSmithMultiDF = pd.merge(HDSmithMultiDF0, reportMultiDF, on=['NDC'], how='left')
        HDSmithMultiDF['HDSMITH UNITS'] = HDSmithMultiDF['HDSMITH']*HDSmithMultiDF['PACK']
        HDSmithMultiDF = HDSmithMultiDF[['DrugID','HDSMITH UNITS']]
        HDSmithMultiDF = HDSmithMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)





AndaFileX = [filename for filename in os.listdir('.') if re.search(r'anda', filename, re.IGNORECASE)] 
if not AndaFileX:
    AndaDF = pd.DataFrame(columns=['NDC'])
    AndaMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    AndaFile = AndaFileX[0]
    if "xls" in AndaFile:
            #AndaDF = pd.read_excel(AndaFile,header= None,sheet_name=0, index= False)
            Andaxl = pd.ExcelFile(AndaFile)
            sheet_names = Andaxl.sheet_names
            Anda_sheet = sheet_names[-1]
            AndaDF = pd.read_excel(AndaFile, header=None, sheet_name = Anda_sheet, index= False)
            AndaDF = AndaDF[((AndaDF.astype(str) == 'NDC').cumsum()).any(1)]
            AndaDF.columns = AndaDF.iloc[0]
            AndaDF = AndaDF.drop(AndaDF.index[0])
            AndaDF = AndaDF.rename(columns={'UNITS_NET':'QTY SHIPPED'})
    elif "csv" in AndaFile:
            AndaDF = pd.read_csv(AndaFile, encoding='ISO-8859-1')
            AndaDF = AndaDF[((AndaDF.astype(str) == 'NDC').cumsum()).any(1)]
            AndaDF.columns = AndaDF.iloc[0]
            AndaDF = AndaDF.drop(AndaDF.index[0])
    AndaDF.to_excel(writer, sheet_name='ANDA', index=False)
    AndaDF['NDC']=AndaDF['NDC'].astype(str).str[:5]+'-'+AndaDF['NDC'].astype(str).str[5:9]+'-'+AndaDF['NDC'].astype(str).str[-2:]
    AndaDF = AndaDF[['NDC','QTY SHIPPED']]
    AndaDF['QTY SHIPPED']=AndaDF['QTY SHIPPED'].apply(float)
    AndaDF = AndaDF.groupby(['NDC'], as_index=False).sum()
    AndaDF = AndaDF.rename(columns={'QTY SHIPPED': 'ANDA'})
    AndaDF = AndaDF[['NDC','ANDA']]
 
    if not DrugID:
        pass
    else:  
        AndaMultiDF0 = AndaDF
        AndaMultiDF = pd.merge(AndaMultiDF0, reportMultiDF, on=['NDC'], how='left')
        AndaMultiDF['ANDA UNITS'] = AndaMultiDF['ANDA']*AndaMultiDF['PACK']
        AndaMultiDF = AndaMultiDF[['DrugID','ANDA UNITS']]
        AndaMultiDF = AndaMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)




CardinalFileX = [filename for filename in os.listdir('.') if re.search(r'cardinal', filename, re.IGNORECASE)] 
if not CardinalFileX:
    CardinalDF = pd.DataFrame(columns=['NDC'])
    CardinalMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    CardinalFile = CardinalFileX[0]
    if "xls" in CardinalFile:
            CardinalDF = pd.read_excel(CardinalFile,header= None,sheet_name=0, index= False)
    elif "csv" in CardinalFile:
            CardinalDF = pd.read_csv(CardinalFile, encoding='ISO-8859-1')
    CardinalDF = CardinalDF[((CardinalDF.astype(str) == 'NDC').cumsum()).any(1)]
    CardinalDF.columns = CardinalDF.iloc[0]
    CardinalDF = CardinalDF.drop(CardinalDF.index[0])
    CardinalDF.to_excel(writer, sheet_name='CARDINAL', index=False)
    CardinalDF['NDC']=CardinalDF['NDC'].astype(str).str[:5]+'-'+CardinalDF['NDC'].astype(str).str[5:9]+'-'+CardinalDF['NDC'].astype(str).str[-2:]
    CardinalDF = CardinalDF[['NDC','Quantity Shipped']]
    CardinalDF['Quantity Shipped']=CardinalDF['Quantity Shipped'].apply(float)
    CardinalDF = CardinalDF.groupby(['NDC'], as_index=False).sum()
    CardinalDF = CardinalDF.rename(columns={'Quantity Shipped': 'CARDINAL'})
    CardinalDF = CardinalDF[['NDC','CARDINAL']]
 
    if not DrugID:
        pass
    else:  
        CardinalMultiDF0 = CardinalDF
        CardinalMultiDF = pd.merge(CardinalMultiDF0, reportMultiDF, on=['NDC'], how='left')
        CardinalMultiDF['CARDINAL UNITS'] = CardinalMultiDF['CARDINAL']*CardinalMultiDF['PACK']
        CardinalMultiDF = CardinalMultiDF[['DrugID','CARDINAL UNITS']]
        CardinalMultiDF = CardinalMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)



HealthcareFileX = [filename for filename in os.listdir('.') if re.search(r'healthcare', filename, re.IGNORECASE)] 
if not HealthcareFileX:
    HealthcareDF = pd.DataFrame(columns=['NDC']) 
    HealthcareMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    HealthcareFile = HealthcareFileX[0]
    if "xls" in HealthcareFile:
            HealthcareDF = pd.read_excel(HealthcareFile,header= None,sheet_name=1, index= False)
    elif "csv" in HealthcareFile:
            HealthcareDF = pd.read_csv(HealthcareFile, encoding='ISO-8859-1')
    HealthcareDF = HealthcareDF.loc[:, HealthcareDF.columns.notnull()]
    HealthcareDF.columns = HealthcareDF.iloc[0]
    HealthcareDF = HealthcareDF.drop(HealthcareDF.index[0])
    HealthcareDF = HealthcareDF.dropna(axis=1, how='all')
    HealthcareDF.to_excel(writer, sheet_name='HEALTHCARE', index=False)
    HealthcareDF = HealthcareDF[['Item','Qty']]
    HealthcareDF.dropna(how='all', inplace=True)
    HealthcareDF['Item'].astype(str)
    HealthcareDF['NDC+'] = HealthcareDF['Item'].str.split('#').str[1:]
    HealthcareDF['NDC1'] = HealthcareDF['NDC+'].astype(str).str.split('-').str[0]
    HealthcareDF['NDC1.1'] = HealthcareDF['NDC1'].astype(str).str.split('\'').str[1]
    HealthcareDF['NDC2'] = HealthcareDF['NDC+'].astype(str).str.split('-').str[1]
    HealthcareDF['NDC3'] = HealthcareDF['NDC+'].astype(str).str.split('-').str[2]
    HealthcareDF['NDC3.1'] = HealthcareDF['NDC3'].astype(str).str.split(')').str[0]
    HealthcareDF['NDC']= HealthcareDF['NDC1.1']+'-'+ HealthcareDF['NDC2']+'-' + HealthcareDF['NDC3.1']
    HealthcareDF = HealthcareDF[['NDC','Qty']]
    HealthcareDF - HealthcareDF.groupby(['NDC'], as_index=False).sum()
    HealthcareDF = HealthcareDF.rename(columns={'Qty': 'HEALTHCARE'})
    HealthcareDF = HealthcareDF[['NDC','HEALTHCARE']]
  
    if not DrugID:
        pass
    else: 
        HealthcareMultiDF0 = HealthcareDF
        HealthcareMultiDF = pd.merge(HealthcareMultiDF0, reportMultiDF, on=['NDC'], how='left')
        HealthcareMultiDF['HEALTHCARE UNITS'] = HealthcareMultiDF['HEALTHCARE']*HealthcareMultiDF['PACK']
        HealthcareMultiDF = HealthcareMultiDF[['DrugID','HEALTHCARE UNITS']]
        HealthcareMultiDF = HealthcareMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)



HerculesFileX = [filename for filename in os.listdir('.') if re.search(r'hercules', filename, re.IGNORECASE)]  
if not HerculesFileX:
    HerculesDF = pd.DataFrame(columns=['NDC'])
    HerculesMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    HerculesFile = HerculesFile[0]
    if "xls" in HerculesFile:
            HerculesDF = pd.read_excel(HerculesFile,header= None, index= False)
    elif "csv" in HerculesFile:
            HerculesDF = pd.read_csv(HerculesFile, encoding='ISO-8859-1')
    HerculesDF.columns = HerculesDF.iloc[0]
    HerculesDF = HerculesDF.drop(HerculesDF.index[0])
    HerculesDF.to_excel(writer, sheet_name='HERCULES', index=False)
    HerculesDF['NDC'] = HerculesDF['NDC / Name']
    HerculesDF = HerculesDF[['NDC','Total Quantity']]
    HerculesDF['Total Quantity']=HerculesDF['Total Quantity'].apply(float)
    HerculesDF = HerculesDF.groupby(['NDC'], as_index=False).sum()
    HerculesDF = HerculesDF.rename(columns={'Total Quantity': 'HERCULES'})
    HerculesDF = HerculesDF[['NDC','HERCULES']]
  
    if not DrugID:
        pass
    else: 
        HerculesMultiDF0 = HerculesDF
        HerculesMultiDF = pd.merge(HerculesMultiDF0, reportMultiDF, on=['NDC'], how='left')
        HerculesMultiDF['HERCULES UNITS'] = HerculesMultiDF['HERCULES']*HerculesMultiDF['PACK']
        HerculesMultiDF = HerculesMultiDF[['DrugID','HERCULES UNITS']]
        HerculesMultiDF = HerculesMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)


IntegralRXFileX = [filename for filename in os.listdir('.') if re.search(r'integral', filename, re.IGNORECASE)] 
if not IntegralRXFileX:
    IntegralRXDF = pd.DataFrame(columns=['NDC'])
    IntegralRXMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    IntegralRXFile = IntegralRXFileX[0]
    if "xls" in IntegralRXFile:
            IntegralRXDF = pd.read_excel(IntegralRXFile,header= None,sheet_name=0, index= False)
    elif "csv" in IntegralRXFile:
            IntegralRXDF = pd.read_csv(IntegralRXFile, encoding='ISO-8859-1')
    IntegralRXDF.columns = IntegralRXDF.iloc[0]
    IntegralRXDF = IntegralRXDF.drop(IntegralRXDF.index[0])
    IntegralRXDF.to_excel(writer, sheet_name='INTEGRALRX', index=False)
    IntegralRXDF['NDC']=IntegralRXDF['NDC'].astype(str).str[:5]+'-'+IntegralRXDF['NDC'].astype(str).str[5:9]+'-'+IntegralRXDF['NDC'].astype(str).str[-2:]
    IntegralRXDF = IntegralRXDF[['NDC','QTY']]
    IntegralRXDF['QTY']=IntegralRXDF['QTY'].apply(float)
    IntegralRXDF = IntegralRXDF.groupby(['NDC'], as_index=False).sum()
    IntegralRXDF = IntegralRXDF.rename(columns={'QTY': 'INTEGRALRX'})
    IntegralRXDF = IntegralRXDF[['NDC','INTEGRALRX']]
 
    if not DrugID:
        pass
    else:  
        IntegralMultiDF0 = IntegralDF
        IntegralMultiDF = pd.merge(IntegralMultiDF0, reportMultiDF, on=['NDC'], how='left')
        IntegralMultiDF['INTEGRALRX UNITS'] = IntegralMultiDF['INTEGRALRX']*IntegralMultiDF['PACK']
        IntegralMultiDF = IntegralMultiDF[['DrugID','INTEGRALRX UNITS']]
        IntegralMultiDF = IntegralMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)


KyMEDSFileX = [filename for filename in os.listdir('.') if re.search(r'kymed', filename, re.IGNORECASE)] 
if not KyMEDSFileX:
    KyMEDSDF = pd.DataFrame(columns=['NDC'])
    KyMEDSMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    KyMEDSFile = KyMEDSFileX[0]
    if "xls" in KyMEDSFile:
            KyMEDSDF = pd.read_excel(KyMEDSFile,header= None,index= False)
    elif "csv" in KyMEDSFile:
            KyMEDSDF = pd.read_csv(KyMEDSFile, encoding='ISO-8859-1')
    KyMEDSDF = KyMEDSDF.loc[:, KyMEDSDF.columns.notnull()]
    KyMEDSDF.columns = KyMEDSDF.iloc[0]
    KyMEDSDF = KyMEDSDF.drop(KyMEDSDF.index[0])
    KyMEDSDF = KyMEDSDF.dropna(axis=1, how='all')
    KyMEDSDF.to_excel(writer, sheet_name='KYMEDS', index=False)
    KyMEDSDF = KyMEDSDF[['Item','Qty']]
    KyMEDSDF.dropna(how='all', inplace=True)
    KyMEDSDF['Item'].astype(str)
    KyMEDSDF['NDC'] = KyMEDSDF['Item'].astype(str).str.split(':').str[0]
    KyMEDSDF = KyMEDSDF.groupby(['NDC'], as_index=False).sum()
    KyMEDSDF = KyMEDSDF.rename(columns={'Qty': 'KYMEDS'})
    KyMEDSDF = KyMEDSDF[['NDC','KYMEDS']]
 
    if not DrugID:
        pass
    else:  
        KyMEDSMultiDF0 = KyMEDSDF
        KyMEDSMultiDF = pd.merge(KyMEDSMultiDF0, reportMultiDF, on=['NDC'], how='left')
        KyMEDSMultiDF['KYMEDS UNITS'] = KyMEDSMultiDF['KYMEDS']*KyMEDSMultiDF['PACK']
        KyMEDSMultiDF = KyMEDSMultiDF[['DrugID','KYMEDS UNITS']]
        KyMEDSMultiDF = KyMEDSMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)





MastersFileX = [filename for filename in os.listdir('.') if re.search(r'master', filename, re.IGNORECASE)] 
if not MastersFileX:
    MastersDF = pd.DataFrame(columns=['NDC'])
    MastersMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    MastersFile = MastersFileX[0]
    if "xls" in MastersFile:
            MastersDF = pd.read_excel(MastersFile,header= None, index= False)
            
    elif "csv" in MastersFile:
            MastersDF = pd.read_csv(MastersFile, encoding='ISO-8859-1')
    MastersDF = MastersDF[((MastersDF.astype(str) == 'NDC').cumsum()).any(1)]
    MastersDF.columns = MastersDF.iloc[0]
    MastersDF = MastersDF.drop(MastersDF.index[0])
    MastersDF.to_excel(writer, sheet_name='Masters', index=False)
    MastersDF = MastersDF[['NDC','Qty']]
    MastersDF['Qty']=MastersDF['Qty'].apply(float)
    MastersDF = MastersDF.groupby(['NDC'], as_index=False).sum()
    MastersDF = MastersDF.rename(columns={'Qty': 'MASTERS'})
 
    if not DrugID:
        pass
    else:    
        MastersMultiDF0 = MastersDF
        MastersMultiDF = pd.merge(MastersMultiDF0, reportMultiDF, on=['NDC'], how='left')
        MastersMultiDF['MASTERS UNITS'] = MastersMultiDF['MASTERS']*MastersMultiDF['PACK']
        MastersMultiDF = MastersMultiDF[['DrugID','MASTERS UNITS']]
        MastersMultiDF = MastersMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)




PaylessFileX = [filename for filename in os.listdir('.') if re.search(r'payless', filename, re.IGNORECASE)] 
if not PaylessFileX:
    PaylessDF = pd.DataFrame(columns=['NDC'])
    PaylessMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    PaylessFile = PaylessFileX[0]
    if "xls" in PaylessFile:
            PaylessDF = pd.read_excel(PaylessFile,header= None, index= False)
    elif "csv" in PaylessFile:
            PaylessDF = pd.read_csv(PaylessFile, encoding='ISO-8859-1')
    PaylessDF = PaylessDF[((PaylessDF.astype(str) == 'Item').cumsum()).any(1)]
    PaylessDF.columns = PaylessDF.iloc[0]
    PaylessDF = PaylessDF.drop(PaylessDF.index[0])
    PaylessDF.to_excel(writer, sheet_name='PAYLESS', index=False)
    PaylessDF = PaylessDF[['Item','Qty']]
    PaylessDF['Item'] = PaylessDF.loc[:, 'Item'].apply(lambda x: x[::-1])
    PaylessDF['NDC1'] = PaylessDF['Item'].astype(str).str.split('-').str[0]
    PaylessDF['NDC2'] = PaylessDF['Item'].astype(str).str.split('-').str[1]
    PaylessDF['NDC3'] = PaylessDF['Item'].astype(str).str.split('-').str[2]
    PaylessDF['NDC3.1'] = PaylessDF['NDC3'].astype(str).str.split(' ').str[0]
    PaylessDF['NDC3.1'] = PaylessDF['NDC3'].astype(str)
    PaylessDF['NDC+'] = PaylessDF['NDC3.1']+ '-' + PaylessDF['NDC2'] + '-' + PaylessDF['NDC1']
    PaylessDF['NDC+'].replace(regex=True,inplace=True,to_replace=r'\D',value=r'')
    PaylessDF['NDC']=PaylessDF['NDC+'].astype(str).str[:5]+'-'+PaylessDF['NDC+'].astype(str).str[5:9]+'-'+PaylessDF['NDC+'].astype(str).str[-2:]
    PaylessDF = PaylessDF[['NDC','Qty']]
    PaylessDF['Qty']=PaylessDF['Qty'].apply(float)
    PaylessDF = PaylessDF.groupby(['NDC'], as_index=False).sum()
    PaylessDF = PaylessDF.rename(columns={'NDC': 'NDC'})
    PaylessDF = PaylessDF.rename(columns={'Qty': 'PAYLESS'})
 
    if not DrugID:
        pass
    else:  
        PaylessMultiDF0 = PaylessDF
        PaylessMultiDF = pd.merge(PaylessMultiDF0, reportMultiDF, on=['NDC'], how='left')
        PaylessMultiDF['PAYLESS UNITS'] = PaylessMultiDF['PAYLESS']*PaylessMultiDF['PACK']
        PaylessMultiDF = PaylessMultiDF[['DrugID','PAYLESS UNITS']]
        PaylessMultiDF = PaylessMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)



PrimedFileX = [filename for filename in os.listdir('.') if re.search(r'primed', filename, re.IGNORECASE)] 
if not PrimedFileX:
    PrimedDF = pd.DataFrame(columns=['NDC'])
    PrimedMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    PrimedFile = PrimedFileX[0]
    if "xls" in PrimedFile:
            PrimedDF = pd.read_excel(PrimedFile,header= None, index= False)
    elif "csv" in PrimedFile:
            PrimedDF = pd.read_csv(PrimedFile, encoding='ISO-8859-1')
    PrimedDF.columns = PrimedDF.iloc[0]
    PrimedDF = PrimedDF.drop(PrimedDF.index[0])
    PrimedDF.to_excel(writer, sheet_name='Primed', index=False)
    PrimedDF = PrimedDF[['Product Code','Quantity']]
    PrimedDF['Quantity']=PrimedDF['Quantity'].apply(float)
    PrimedDF = PrimedDF.rename(columns={'Product Code':'NDC'})
    PrimedDF = PrimedDF.groupby(['NDC'], as_index=False).sum()
    PrimedDF['NDC']=PrimedDF['NDC'].astype(str).str[:5]+'-'+PrimedDF['NDC'].astype(str).str[5:9]+'-'+PrimedDF['NDC'].astype(str).str[-2:]
    PrimedDF = PrimedDF.rename(columns={'Quantity': 'PRIMED'})
 
    if not DrugID:
        pass
    else:  
        PrimedMultiDF0 = PrimedDF
        PrimedMultiDF = pd.merge(PrimedMultiDF0, reportMultiDF, on=['NDC'], how='left')
        PrimedMultiDF['PRIMED UNITS'] = PrimedMultiDF['PRIMED']*PrimedMultiDF['PACK']
        PrimedMultiDF = PrimedMultiDF[['DrugID','PRIMED UNITS']]
        PrimedMultiDF = PrimedMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)


RedmondFileX = [filename for filename in os.listdir('.') if re.search(r'redmond', filename, re.IGNORECASE)] 
if not RedmondFileX:
    RedmondDF = pd.DataFrame(columns=['NDC'])
    RedmondMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    RedmondFile = RedmondFileX[0]
    if "xls" in RedmondFile:
            RedmondDF = pd.read_excel(RedmondFile,header= None, index= False)
    elif "csv" in RedmondFile:
            RedmondDF = pd.read_csv(RedmondFile, encoding='ISO-8859-1')
    RedmondDF.columns = RedmondDF.iloc[0]
    RedmondDF = RedmondDF.drop(RedmondDF.index[0])
    RedmondDF.to_excel(writer, sheet_name='REDMOND', index=False)
    RedmondDF = RedmondDF[['Product Code','Quantity']]
    RedmondDF['Quantity']=RedmondDF['Quantity'].apply(float)
    RedmondDF = RedmondDF.rename(columns={'Product Code':'NDC'})
    RedmondDF = RedmondDF.groupby(['NDC'], as_index=False).sum()
    RedmondDF = RedmondDF.rename(columns={'Quantity': 'REDMOND'})
 
    if not DrugID:
        pass
    else:  
        RedmondMultiDF0 = RedmondDF
        RedmondMultiDF = pd.merge(RedmondMultiDF0, reportMultiDF, on=['NDC'], how='left')
        RedmondMultiDF['REDMOND UNITS'] = RedmondMultiDF['REDMOND']*RedmondMultiDF['PACK']
        RedmondMultiDF = RedmondMultiDF[['DrugID','REDMOND UNITS']]
        RedmondMultiDF = RedmondMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)


RXSupplyFileX = [filename for filename in os.listdir('.') if re.search(r'rxsupply', filename, re.IGNORECASE)] 
if not RXSupplyFileX:
    RXSupplyDF = pd.DataFrame(columns=['NDC'])
    RXSupplyMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    RXSupplyFile = RXSupplyFileX[0]
    if "xls" in RXSupplyFile:
            RXSupplyDF = pd.read_excel(RXSupplyFile,header= None, index= False)
            
    elif "csv" in RXSupplyFile:
            RXSupplyDF = pd.read_csv(RXSupplyFile, encoding='ISO-8859-1')
    RXSupplyDF.columns = RXSupplyDF.iloc[0]
    RXSupplyDF = RXSupplyDF.drop(RXSupplyDF.index[0])
    RXSupplyDF.to_excel(writer, sheet_name='RXSUPPLY', index=False)
    RXSupplyDF = RXSupplyDF[~RXSupplyDF['NDC'].str.contains('Total')]
    RXSupplyDF = RXSupplyDF[['NDC','QUANTITY']]
    RXSupplyDF['QUANTITY']=RXSupplyDF['QUANTITY'].apply(float)
    RXSupplyDF['NDC']=RXSupplyDF['NDC'].astype(str).str[:5]+'-'+RXSupplyDF['NDC'].astype(str).str[5:9]+'-'+RXSupplyDF['NDC'].astype(str).str[-2:]
    RXSupplyDF = RXSupplyDF.groupby(['NDC'], as_index=False).sum()
    RXSupplyDF = RXSupplyDF.rename(columns={'QUANTITY': 'RXSUPPLY'})
 
    if not DrugID:
        pass
    else:  
        RXSupplyMultiDF0 = RXSupplyDF
        RXSupplyMultiDF = pd.merge(RXSupplyMultiDF0, reportMultiDF, on=['NDC'], how='left')
        RXSupplyMultiDF['RXSUPPLY UNITS'] = RXSupplyMultiDF['RXSUPPLY']*RXSupplyMultiDF['PACK']
        RXSupplyMultiDF = RXSupplyMultiDF[['DrugID','RXSUPPLY UNITS']]
        RXSupplyMultiDF = RXSupplyMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)



TRXadeFileX = [filename for filename in os.listdir('.') if re.search(r'trxade', filename, re.IGNORECASE)] 
if not TRXadeFileX:
    TRXadeDF = pd.DataFrame(columns=['NDC'])
    TRXadeMultiDF = pd.DataFrame(columns=['DrugID'])
else:
    TRXadeFile = TRXadeFileX[0]
    if "xls" in TRXadeFile:
            TRXadeDF = pd.read_excel(TRXadeFile,header= None, index= False)
    elif "csv" in TRXadeFile:
            TRXadeDF = pd.read_csv(TRXadeFile, encoding='ISO-8859-1')
    TRXadeDF.columns = TRXadeDF.iloc[0]
    TRXadeDF = TRXadeDF.drop(TRXadeDF.index[0])
    TRXadeDF.to_excel(writer, sheet_name='TRXADE', index=False)
    TRXadeDF = TRXadeDF[['NDC','Qty Fulfilled']]
    TRXadeDF['Qty Fulfilled']=TRXadeDF['Qty Fulfilled'].apply(float)
    TRXadeDF['NDC'].replace(regex=True,inplace=True,to_replace=r'-',value=r'')
    TRXadeDF['NDC']=TRXadeDF['NDC'].astype(str).str[:5]+'-'+TRXadeDF['NDC'].astype(str).str[5:9]+'-'+TRXadeDF['NDC'].astype(str).str[-2:]
    TRXadeDF = TRXadeDF.groupby(['NDC'], as_index=False).sum()
    TRXadeDF = TRXadeDF.rename(columns={'Qty Fulfilled': 'TRXADE'})
 
    if not DrugID:
        pass
    else:  
        TRXadeMultiDF0 = TRXadeDF
        TRXadeMultiDF = pd.merge(TRXadeMultiDF0, reportMultiDF, on=['NDC'], how='left')
        TRXadeMultiDF['TRXADE UNITS'] = TRXadeMultiDF['TRXADE']*TRXadeMultiDF['PACK']
        TRXadeMultiDF = TRXadeMultiDF[['DrugID','TRXADE UNITS']]
        TRXadeMultiDF = TRXadeMultiDF.drop_duplicates(subset='DrugID', keep='first', inplace=False)


end  = time.time()
print( end - start)


def vendor(file_name, df_name):
    if any([i for i in files if any(x in i for x in Vendor)]):
        Tk().withdraw()
        file_name = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT VENDOR FILE') 
        if "xls" in file_name:
                df_name = pd.read_excel(file_name,header= None, index= False)
        elif "csv" in file_name:
                df_name = pd.read_csv(file_name, encoding='ISO-8859-1')
        df_name.columns = df_name.iloc[0]
        df_name = df_name.drop(df_name.index[0])
        df_name.to_excel(writer, sheet_name='ANDA', index=False)
        df_name['NDC']=df_name['NDC#'].astype(str).str[:5]+'-'+df_name['NDC#'].astype(str).str[5:9]+'-'+df_name['NDC#'].astype(str).str[-2:]
        df_name = df_name[['NDC','QUANTITY']]
        df_name[':QUANTITY']=df_name['QUANTITY'].apply(float)
        df_name = df_name.groupby(['NDC'], as_index=False).sum()
        df_name = df_name.rename(columns={'QUANTITY': 'VENDOR'})
        df_name = df_name[['NDC','VENDOR']]
    else:
        df_name = pd.DataFrame(columns=['NDC'])


reportDF0 = reportDF.drop_duplicates(subset=['NDC'], keep='first')





loadscript('Creating Report sheet', 6, .6)
reportDF1 = pd.merge(reportDF0, kinrayRXDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, kinrayOTCDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, MCKDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, TopRXDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, AmerisourceDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, OakDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, MaksDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, AlpineDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, HDSmithDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, AndaDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, CardinalDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, HealthcareDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, HerculesDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, IntegralRXDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, KyMEDSDF , on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, MastersDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, PaylessDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, PrimedDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, RedmondDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, RXSupplyDF, on=['NDC'], how='left')
reportDF = pd.merge(reportDF2, TRXadeDF, on=['NDC'], how='left')






reportPurchDF = reportDF
reportPurchDF = reportPurchDF.drop(['NDC','DRUG NAME','STRENGTH', 'PACK','QUANTITY','DISP'], axis=1)
reportPurchDF['TOTAL']= reportPurchDF.sum(axis=1)
reportPurchDF = reportPurchDF['TOTAL']
ReportDF = pd.concat([reportDF,reportPurchDF], axis=1)
ReportDF['DISC'] = ReportDF['TOTAL'] - ReportDF['DISP']


if not DrugID:
    pass
else:
    
    reportMultiDF1 = pd.merge(reportMultiDF, kinrayRXMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, kinrayOTCMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, MCKMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, TopRXMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, AmerisourceMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, OakMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, MaksMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, AlpineMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, HDSmithMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, AndaMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, CardinalMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, HealthcareMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, HerculesMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, IntegralRXMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, KyMEDSMultiDF , on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, MastersMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, PaylessMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, PrimedMultiDF, on=['DrugID'], how='left')
    reportMultiDF1 = pd.merge(reportMultiDF2, RedmondMultiDF, on=['DrugID'], how='left')
    reportMultiDF2 = pd.merge(reportMultiDF1, RXSupplyMultiDF, on=['DrugID'], how='left')
    reportMultiDF = pd.merge(reportMultiDF2, TRXadeMultiDF, on=['DrugID'], how='left')
    reportMultiDF = reportMultiDF1
    reportPurchMultiDF = reportMultiDF
    reportPurchMultiDF = reportPurchMultiDF.drop(['NDC','DrugID','DRUG NAME','STRENGTH','PACK','DISP'], axis=1)
    reportPurchMultiDF['TOTAL UNITS']= reportPurchMultiDF.sum(axis=1)
    reportPurchMultiDF = reportPurchMultiDF['TOTAL UNITS']
    ReportMultiDF = pd.concat([reportMultiDF,reportPurchMultiDF], axis=1)
    ReportMultiDF['DISC'] = ReportMultiDF['TOTAL UNITS'] - ReportMultiDF['DISP']
    
    ReportMultiDF = ReportMultiDF.drop_duplicates(subset='NDC', keep='first', inplace=False)
    #ReportMultiDF = ReportMultiDF.dropna(subset=['DISP'])

if not DrugID:
    ReportDF.to_excel(writer, sheet_name='Report', index=False)
else:
    ReportDF.to_excel(writer, sheet_name='Report', index=False)
    ReportMultiDF.to_excel(writer, sheet_name='Multi NDC', index=False)

writer.save()
#os.system("open -a 'Microsoft Excel.app' 'TrialReport.xlsx'")
# Windows - os.system('start excel.exe "%s\\TrialReport.xls"' % (sys.path[0], ))


