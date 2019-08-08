import pandas as pd
import numpy as np
import codecs
import time
import os
import sys
from tkinter import filedialog
from tkinter import *
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def loadscript(a,b,c):
    print(a,b,c)

def loadscripts(step, trial, duration):
    print(step, end='', flush=True)
    for i in range(trial): 
        time.sleep(duration)
        print('.', end='', flush=True)
    time.sleep(.2)
    print("\n")

files = os.listdir(os.curdir)
Kinray = ['KIN','kin','Kin']
Mckesson = ['Mckesson','MCK','mck','Mck']
Toprx = ['TOP','top','Top']
ABC = ['ABC','abc','Abc']




writer = pd.ExcelWriter('TrialReport.xlsx', engine='xlsxwriter')
Tk().withdraw()
PrimeRXFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT Dispense FILE') 
if "xls" in PrimeRXFile:
    rxRawDF = pd.read_excel(PrimeRXFile,header= None, index= False)
elif "csv" in PrimeRXFile:
    rxRawDF = pd.read_csv(PrimeRXFile,header= None)
    rxRawDF.columns = rxRawDF.iloc[0]
    rxRawDF = rxRawDF.drop(rxRawDF.index[0])

rxRawDF.to_excel(writer, sheet_name='PrimeRx RAW', index=False)
rxRawDF[['PACKAGESIZE','QUANT']]= rxRawDF[['PACKAGESIZE','QUANT']].apply(pd.to_numeric, errors='coerce')
loadscript('Copying over PrimeRX', 6, .4)
reportDF = rxRawDF[['NDC','DRGNAME','DRUGSTRONG', 'PACKAGESIZE','QUANT']]
#print(rxRawDF.head(10))




if any([i for i in files if any(x in i for x in Kinray)]):
    Tk().withdraw()
    KINRXFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT KINRAY (RX) FILE') 
    if "xls" in KINRXFile:
            kinrayRXDF = pd.read_excel(KINRXFile,header= None, names=[0,1,2,3])
    elif "csv" in KINRXFile:
            kinrayRXDF = pd.read_csv(KINRXFile,header= None, names=[0,1,2,3])
    kinrayRXDF = kinrayRXDF.drop(kinrayRXDF.index[[0,1,2,3,4,5,6]])
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
            kinrayOTCDF = pd.read_excel(KINOTCFile,header= None, index= False , names=[0,1,2,3])
    elif "csv" in KINOTCFile:
            kinrayOTCDF = pd.read_csv(KINOTCFile,header= None, index= False , names=[0,1,2,3])
    kinrayOTCDF = kinrayOTCDF.drop(kinrayOTCDF.index[[0,1,2,3,4,5,6]])
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
            MCKDF = MCKDF.drop(MCKDF.index[[0,1,2,3,4,5,6,7,8,9]])
    elif "csv" in MCKFile:
            MCKDF = pd.read_csv(MCKFile,header= None, encoding='ISO-8859-1')
            MCKDF = MCKDF.drop(MCKDF.index[[0,1,2,3,4,5,6]])
    MCKDF.columns = MCKDF.iloc[0]
    MCKDF = MCKDF.drop(MCKDF.index[0])
    MCKDF.to_excel(writer, sheet_name='MCK', index=False)
    MCKDF['NDC']=MCKDF['NDC/UPC'].astype(str).str[:5]+'-'+MCKDF['NDC/UPC'].astype(str).str[5:9]+'-'+MCKDF['NDC/UPC'].astype(str).str[-2:]
    MCKDF = MCKDF[['NDC','Net']]
    MCKDF['Net']=MCKDF['Net'].apply(float)
    MCKDF = MCKDF.groupby(['NDC'], as_index=False).sum()
    MCKDF = MCKDF.rename(columns={'Net': 'MCK'})
    MCKDF = MCKDF[['NDC','MCK']]
else:
    MCKDF = pd.DataFrame(columns=['NDC']) 


if any([i for i in files if any(x in i for x in Toprx)]):
    Tk().withdraw()
    TopRXFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT TOPRX FILE') 
    if "xls" in TopRXFile:
            TopRXDF = pd.read_excel(TopRXFile,header= None, index= False)
    elif "csv" in TopRXFile:
            TopRXDF = pd.read_csv(TopRXFile,header= None, encoding='ISO-8859-1')
    TopRXDF.columns = TopRXDF.iloc[0]
    TopRXDF = TopRXDF.drop(TopRXDF.index[0])
    TopRXDF.to_excel(writer, sheet_name='TOPRX', index=False)
    TopRXDF['NDC']=TopRXDF['NDC#'].astype(str).str[:5]+'-'+TopRXDF['NDC#'].astype(str).str[5:9]+'-'+TopRXDF['NDC#'].astype(str).str[-2:]
    TopRXDF = TopRXDF[['NDC','QUANTITY']]
    TopRXDF['QUANTITY']=TopRXDF['QUANTITY'].apply(float)
    TopRXDF = TopRXDF.groupby(['NDC'], as_index=False).sum()
    TopRXDF = TopRXDF.rename(columns={'QUANTITY': 'TopRX'})
    TopRXDF = TopRXDF[['NDC','TopRX']]
else:
    TopRXDF = pd.DataFrame(columns=['NDC'])



if any([i for i in files if any(x in i for x in ABC )]):
    Tk().withdraw()
    ABCFile = askopenfilename(initialdir=os.getcwd(), title='PLEASE SELECT ABC FILE') 
    if "xls" in ABCFile:
        AmerisourceDF = pd.read_excel(ABCFile,header= None, index= False, sheet_name='Item Level Detail')
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
        AmerisourceDF.columns = AmerisourceDF.iloc[0]
        AmerisourceDF = AmerisourceDF.drop(AmerisourceDF.index[0])
        print(AmerisourceDF.head())
        AmerisourceDF['NDC']=AmerisourceDF['NDC'].astype(str).str[:5]+'-'+AmerisourceDF['NDC'].astype(str).str[5:9]+'-'+AmerisourceDF['NDC'].astype(str).str[-2:]
        AmerisourceDF = AmerisourceDF[['NDC','Shipped Qty']]
        AmerisourceDF['Shipped Qty']=AmerisourceDF['Shipped Qty'].apply(float)
        AmerisourceDF = AmerisourceDF.groupby(['NDC'], as_index=False).sum()
        AmerisourceDF = AmerisourceDF.rename(columns={'Shipped Qty': 'ABC'})
        AmerisourceDF = AmerisourceDF[['NDC','ABC']]
else:
    AmerisourceDF = pd.DataFrame(columns=['NDC'])




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
        df_name.to_excel(writer, sheet_name='TOPRX', index=False)
        df_name['NDC']=df_name['NDC#'].astype(str).str[:5]+'-'+df_name['NDC#'].astype(str).str[5:9]+'-'+df_name['NDC#'].astype(str).str[-2:]
        df_name = df_name[['NDC','QUANTITY']]
        df_name['QUANTITY']=df_name['QUANTITY'].apply(float)
        df_name = df_name.groupby(['NDC'], as_index=False).sum()
        df_name = df_name.rename(columns={'QUANTITY': 'VENDOR'})
        df_name = df_name[['NDC','VENDOR']]
    else:
        df_name = pd.DataFrame(columns=['NDC'])


loadscript('Creating Report sheet', 6, .6)
reportDF = reportDF.groupby(['NDC','DRGNAME','DRUGSTRONG', 'PACKAGESIZE',], as_index=False).sum()
reportDF['Total Dispense'] = reportDF['QUANT']/reportDF['PACKAGESIZE']
reportDF['Total Dispense'] = reportDF['Total Dispense'].apply(lambda x:round(x,1))
reportDF1 = pd.merge(reportDF, kinrayRXDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, kinrayOTCDF, on=['NDC'], how='left')
reportDF1 = pd.merge(reportDF2, MCKDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, TopRXDF, on=['NDC'], how='left')
reportDF = pd.merge(reportDF2, AmerisourceDF, on=['NDC'], how='left')

print(reportDF.head())
reportPurchDF = reportDF
reportPurchDF = reportPurchDF.drop(['NDC','DRGNAME','DRUGSTRONG', 'PACKAGESIZE','QUANT','Total Dispense'], axis=1)


col_len = len(reportPurchDF.columns)*-1
reportPurchDF['Total Purchase'] = reportPurchDF.iloc[:,col_len:]

print(reportPurchDF.head())




#if all([item in reportDF.columns for item in ['ABC']]):
 #   reportDF['Total Purchase'] = reportDF[['ABC']].sum(axis=1)
'''
if all([item in reportDF.columns for item in ['KIN RX','KIN OTC', 'MCK','TOP','ABC']]):
    reportDF['Total Purchase'] = reportDF[['KIN RX','KIN OTC', 'MCK','TOP','ABC']].sum(axis=1)
elif all([item in reportDF.columns for item in ['KIN RX','KIN OTC', 'MCK','TOP']]):
    reportDF['Total Purchase'] = reportDF[['KIN RX','KIN OTC', 'MCK','TOP']].sum(axis=1)
elif all([item in reportDF.columns for item in ['KIN RX','KIN OTC', 'MCK']]):
    reportDF['Total Purchase'] = reportDF[['KIN RX','KIN OTC', 'MCK']].sum(axis=1)
elif all([item in reportDF.columns for item in ['KIN RX','KIN OTC']]):
    reportDF['Total Purchase'] = reportDF[['KIN RX','KIN OTC']].sum(axis=1)
elif all([item in reportDF.columns for item in ['KIN RX']]):
    reportDF['Total Purchase'] = reportDF[['KIN RX']].sum(axis=1)
'''

ReportDF = pd.concat([reportDF,reportPurchDF], axis=1)

ReportDF['DISC'] = ReportDF['Total Dispense'] - ReportDF['Total Purchase']
ReportDF.to_excel(writer, sheet_name='Report', index=False)

writer.save()
os.system("open -a 'Microsoft Excel.app' 'TrialReport.xlsx'")
# Windows - os.system('start excel.exe "%s\\TrialReport.xls"' % (sys.path[0], ))

