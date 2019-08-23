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
from tkinter import filedialog
from tkinter import *
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from functools import reduce




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
Kinray = ['KIN','kin','Kin']
Mckesson = ['Mckesson','MCK','mck','Mck']
Toprx = ['TOP','top','Top','Toprx','TOPRX','toprx']
ABC = ['ABC','abc','Abc']
Oak = ['OAK', 'Oak', 'oak',]
Maks = ['MAKS','Maks','maks']
Alpine = ['ALP','Alp','alp','Alpine','ALPINE']
HDSmith = ['HDSMITH','Hdsmith','HDS']
Anda = ['ANDA','anda','Anda']
Cardinal = ['Cardinal','CARDINAL','cardinal']
Healthcare = ['healthcare','Healthcare','HEALTCARE']
Hercules = ['Hercules','HERCULES','hercules']
Integralrx = ['integralrx','Integralrx','IntegralRX','INTEGRALRX']
Kymeds = ['Kymeds','KYMEDS','kymeds']
Masters = ['masters','Masters','MASTERS']
Payless = ['payless','Payless','PAYLESS']
Primed = ['primed','Primed','PRIMED']
Redmond = ['redmond','Redmond','REDMOND']
Rxsupply = ['rxsupply','Rxsupply','RXSUPPLY']
Trxade = ['trxade','Trxade','TRXADE']


writer = pd.ExcelWriter('TrialReport.xlsx', engine='xlsxwriter')



Tk().withdraw()



PrimeRXFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Dispense FILE') 
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
rawColumns = ['NDC','DRGNAME','DRUGNAME','DRUG NAME','DRUG NAME ','DRUGSTRONG','Pack','PACK','PACKAGESIZE','QTY','Quantity','QUANT','Quant','STRENGTH']
rxRawDF = rxRawDF[np.intersect1d(rxRawDF.columns, rawColumns)]
#rxRawDF = rxRawDF.rename(columns={'DRGNAME': 'DRUG NAME','DRUGNAME':'DRUG NAME', 'DRUGSTRONG':'STRENGTH', 'PACKAGESIZE':'PACK','QUANT':'QUANTITY' })
rxRawDF.columns = rxRawDF.columns.str.replace('DRGNAME','DRUG NAME')
rxRawDF.columns = rxRawDF.columns.str.replace('DRUGNAME','DRUG NAME')
rxRawDF.columns = rxRawDF.columns.str.replace('DRUG NAME ','DRUG NAME')
rxRawDF.columns = rxRawDF.columns.str.replace('DRUGSTRONG','STRENGTH')
rxRawDF.columns = rxRawDF.columns.str.replace('PACKAGESIZE','PACK')
rxRawDF.columns = rxRawDF.columns.str.replace('QUANT','QUANTITY')
#df.columns = df.columns.str.replace('agriculture', 'agri')
rxRawDF[['PACK','QUANTITY']]= rxRawDF[['PACK','QUANTITY']].apply(pd.to_numeric, errors='coerce')
loadscript('Copying over PrimeRX', 6, .4)
reportDF = rxRawDF[['NDC','DRUG NAME','STRENGTH', 'PACK','QUANTITY']].copy()
reportDF['QUANTITY'] = reportDF['QUANTITY'].apply(float)
reportDF['PACK'] = reportDF['PACK'].apply(float)
#reportDF1 = reportDF.groupby(['NDC','DRGNAME','DRUGSTRONG', 'PACKAGESIZE',], as_index=False).sum()
reportQuantDF = reportDF.groupby(['NDC'], as_index=False).sum()
reportQuantDF = reportQuantDF.drop(['PACK'], axis=1)
reportDF = reportDF.drop(['QUANTITY'], axis=1)
reportDF = pd.merge(reportDF, reportQuantDF, on=['NDC'], how='left')
reportDF['DISP'] = reportDF['QUANTITY']/reportDF['PACK']
reportDF['DISP'] = reportDF['DISP'].apply(lambda x:round(x,1))





if any([i for i in files if any(x in i for x in Kinray)]):
    Tk().withdraw()
    KINRXFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT KINRAY (RX) FILE') 
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
else:
    kinrayRXDF = pd.DataFrame(columns=['NDC'])






if any([i for i in files if any(x in i for x in Kinray)]):
    Tk().withdraw()
    KINOTCFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT KINRAY (OTC) FILE') 
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
else:
    kinrayOTCDF = pd.DataFrame(columns=['NDC'])

if any([i for i in files if any(x in i for x in Mckesson)]):
    Tk().withdraw()
    MCKFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT MCKESSON FILE') 
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
    MCKDF['Net'] = pd.to_numeric(MCKDF['Net'], errors='coerce')
    MCKDF = MCKDF.groupby(['NDC/UPC'], as_index=False).sum()
    MCKDF = MCKDF.rename(columns={'NDC/UPC': 'NDC'})
    MCKDF = MCKDF.rename(columns={'Net': 'MCK'})
else:
    MCKDF = pd.DataFrame(columns=['NDC']) 


if any([i for i in files if any(x in i for x in Toprx)]):
    Tk().withdraw()
    TopRXFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT TOPRX FILE') 
    if "xls" in TopRXFile:
            TopRXDF = pd.read_excel(TopRXFile,header= None, index= False)
    elif "csv" in TopRXFile:
            TopRXDF = pd.read_csv(TopRXFile,header= None, encoding='ISO-8859-1')
    TOPRXDF = TOPRXDF[((TOPRXDF.astype(str) == 'NDC#').cumsum()).any(1)]
    TopRXDF.columns = TopRXDF.iloc[0]
    TopRXDF = TopRXDF.drop(TopRXDF.index[0])
    TopRXDF.to_excel(writer, sheet_name='TOPRX', index=False)
    TopRXDF['NDC']=TopRXDF['NDC#'].astype(str).str[:5]+'-'+TopRXDF['NDC#'].astype(str).str[5:9]+'-'+TopRXDF['NDC#'].astype(str).str[-2:]
    TopRXDF = TopRXDF[['NDC','QUANTITY']]
    TopRXDF['QUANTITY']=TopRXDF['QUANTITY'].apply(float)
    TopRXDF = TopRXDF.groupby(['NDC'], as_index=False).sum()
    TopRXDF = TopRXDF.rename(columns={'QUANTITY': 'TopRX'})
else:
    TopRXDF = pd.DataFrame(columns=['NDC'])



if any([i for i in files if any(x in i for x in ABC )]):
    Tk().withdraw()
    ABCFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT ABC FILE') 
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
else:
    AmerisourceDF = pd.DataFrame(columns=['NDC'])


if any([i for i in files if any(x in i for x in Oak)]):
    Tk().withdraw()
    OakFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Oak FILE') 
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
else:
    OakDF = pd.DataFrame(columns=['NDC'])



if any([i for i in files if any(x in i for x in Maks)]):
    Tk().withdraw()
    MaksFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Maks FILE') 
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
else:
    MaksDF = pd.DataFrame(columns=['NDC'])


if any([i for i in files if any(x in i for x in Alpine)]):
    Tk().withdraw()

    
    AlpineFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Alpine FILE') 
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
else:
    AlpineDF = pd.DataFrame(columns=['NDC'])



if any([i for i in files if any(x in i for x in HDSmith)]):
    Tk().withdraw()
    HDSmithFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT HDSmith FILE') 
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
else:
    HDSmithDF = pd.DataFrame(columns=['NDC'])



if any([i for i in files if any(x in i for x in Anda)]):
    Tk().withdraw()
    AndaFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT ANDA FILE') 
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
else:
    AndaDF = pd.DataFrame(columns=['NDC'])


if any([i for i in files if any(x in i for x in Cardinal)]):
    Tk().withdraw()
    CardinalFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT CARDINAL FILE') 
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
else:
    CardinalDF = pd.DataFrame(columns=['NDC'])



if any([i for i in files if any(x in i for x in Healthcare)]):
    Tk().withdraw()
    HealthcareFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT HEALTHCARE FILE') 
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
else:
    HealthcareDF = pd.DataFrame(columns=['NDC']) 







if any([i for i in files if any(x in i for x in Hercules)]):
    Tk().withdraw()
    HerculesFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT HERCULES FILE') 
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
else:
    HerculesDF = pd.DataFrame(columns=['NDC'])



if any([i for i in files if any(x in i for x in Integralrx)]):
    Tk().withdraw()
    IntegralRXFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT INTEGRALRX FILE') 
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
else:
    IntegralRXDF = pd.DataFrame(columns=['NDC'])


if any([i for i in files if any(x in i for x in Kymeds)]):
    Tk().withdraw()
    KyMEDSFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT KYMEDS FILE') 
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
else:
    KyMEDSDF = pd.DataFrame(columns=['NDC'])


if any([i for i in files if any(x in i for x in Masters)]):
    Tk().withdraw()
    MastersFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Masters FILE') 
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
    MastersDF = MastersDF.rename(columns={'Qty': 'HDSMITH'})
else:
    MastersDF = pd.DataFrame(columns=['NDC'])

if any([i for i in files if any(x in i for x in Payless)]):
    Tk().withdraw()
    PaylessFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT PAYLESS FILE') 
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
else:
    PaylessDF = pd.DataFrame(columns=['NDC'])

if any([i for i in files if any(x in i for x in Primed)]):
    Tk().withdraw()
    PrimedFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Primed FILE') 
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
else:
    PrimedDF = pd.DataFrame(columns=['NDC'])

if any([i for i in files if any(x in i for x in Redmond)]):
    Tk().withdraw()
    RedmondFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Redmond FILE') 
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
else:
    RedmondDF = pd.DataFrame(columns=['NDC'])

if any([i for i in files if any(x in i for x in Rxsupply)]):
    Tk().withdraw()
    RXSupplyFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT RXSupply FILE') 
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
    RXSupplyDF = RXSupplyDF.rename(columns={'QUANTITY': 'REDMOND'})
else:
    RXSupplyDF = pd.DataFrame(columns=['NDC'])


if any([i for i in files if any(x in i for x in Trxade)]):
    Tk().withdraw()
    TRXadeFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT TRXADE FILE') 
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
else:
    TRXadeDF = pd.DataFrame(columns=['NDC'])




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

reportDF = reportDF.drop_duplicates(subset=['NDC'], keep='first')


loadscript('Creating Report sheet', 6, .6)
reportDF1 = pd.merge(reportDF, kinrayRXDF, on=['NDC'], how='left')
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
col_len = len(reportPurchDF.columns)*-1
reportPurchDF = reportPurchDF['TOTAL']
ReportDF = pd.concat([reportDF,reportPurchDF], axis=1)
ReportDF['DISC'] = ReportDF['TOTAL'] - ReportDF['DISP']
ReportDF.to_excel(writer, sheet_name='Report', index=False)
writer.save()
#os.system("open -a 'Microsoft Excel.app' 'TrialReport.xlsx'")
# Windows - os.system('start excel.exe "%s\\TrialReport.xls"' % (sys.path[0], ))


