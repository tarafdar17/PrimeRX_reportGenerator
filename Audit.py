import pandas as pd
import time
import os
import sys


def loadscript(a,b,c):
    print(a,b,c)

def loadscripts(step, trial, duration):
    print(step, end='', flush=True)
    for i in range(trial): 
        time.sleep(duration)
        print('.', end='', flush=True)
    time.sleep(.2)
    print("\n")

#loadscript('Copying over PrimeRX', 6, .4)
writer = pd.ExcelWriter('TrialReport.xlsx', engine='xlsxwriter')
rxRawDF = pd.read_csv('PrimeRx.csv', low_memory=False)
rxRawDF.to_excel(writer, sheet_name='PrimeRx RAW', index=False)
reportDF = rxRawDF[['NDC','DRGNAME','DRUGSTRONG', 'PACKAGESIZE','QUANT']]
#print(rxRawDF.head(10))


#loadscript('Copying over Kinray RX', 6, .4)
if os.path.exists('KINRX.xls'):
    kinrayRXDF = pd.read_excel('KINRX.xls',header= None, index= False , names=[0,1,2,3])
    kinrayRXDF = kinrayRXDF.drop(kinrayRXDF.index[[0,1,2,3,4,5,6]])
    kinrayRXDF.columns = kinrayRXDF.iloc[0]
    kinrayRXDF = kinrayRXDF.rename(columns={'Universal NDC': 'U NDC'})
    kinrayRXDF['NDC']=kinrayRXDF['U NDC'].astype(str).str[:5]+'-'+kinrayRXDF['U NDC'].astype(str).str[5:9]+'-'+kinrayRXDF['U NDC'].astype(str).str[-2:]
    kinrayRXDF = kinrayRXDF.drop(kinrayRXDF.index[0])
    kinrayRXDF = kinrayRXDF.groupby(['NDC'], as_index=False).sum()
    kinrayRXDF = kinrayRXDF[['NDC', 'Drug Name', 'Qty']]
    kinrayRX_sheet = kinrayRXDF[['NDC', 'Drug Name', 'Qty']]
    kinrayRXDF = kinrayRXDF.rename(columns={'Qty': 'KIN RX'})
    #print(kinrayRXDF.head())
    kinrayRXDF = kinrayRXDF[['NDC','KIN RX']]
    kinrayRX_sheet.to_excel(writer, sheet_name='KIN_RX', index=False)
else:
    kinrayRXDF = pd.DataFrame(columns=['NDC']) 


#loadscript('Copying over Kinray OTC', 6, .3)
if os.path.exists('KINOTC.xls'): 
    kinrayOTCDF = pd.read_excel('KINOTC.xls',header= None, index= False , names=[0,1,2,3])
    kinrayOTCDF = kinrayOTCDF.drop(kinrayOTCDF.index[[0,1,2,3,4,5,6]])
    kinrayOTCDF.columns = kinrayOTCDF.iloc[0]
    kinrayOTCDF = kinrayOTCDF.rename(columns={'Universal NDC': 'U NDC'})
    kinrayOTCDF['NDC']=kinrayOTCDF['U NDC'].astype(str).str[:5]+'-'+kinrayOTCDF['U NDC'].astype(str).str[5:9]+'-'+kinrayOTCDF['U NDC'].astype(str).str[-2:]
    kinrayOTCDF = kinrayOTCDF.drop(kinrayOTCDF.index[0])
    kinrayOTCDF = kinrayOTCDF.groupby(['NDC'], as_index=False).sum()
    kinrayOTCDF = kinrayOTCDF[['NDC', 'Drug Name', 'Qty']]
    kinrayOTC_sheet = kinrayOTCDF[['NDC', 'Drug Name', 'Qty']]
    kinrayOTCDF = kinrayOTCDF.rename(columns={'Qty': 'KIN OTC'})
    #print(kinrayOTCDF.head())
    kinrayOTCDF = kinrayOTCDF[['NDC','KIN OTC']]
    kinrayOTC_sheet.to_excel(writer, sheet_name='KIN_OTC',index=False)
else:
    kinrayOTCDF = pd.DataFrame(columns=['NDC']) 



#loadscript('Copying over McKesson', 6, .3)
if os.path.exists('MCK.csv'): 
    MCKDF = pd.read_csv('MCK.csv', header= None)
    MCKDF = MCKDF.drop(MCKDF.index[[0,1,2,3,4,5,6]])
    MCKDF.columns = MCKDF.iloc[0]
    MCKDF = MCKDF.drop(MCKDF.index[0])
    MCKDF['NDC']=MCKDF['NDC/UPC'].astype(str).str[:5]+'-'+MCKDF['NDC/UPC'].astype(str).str[5:9]+'-'+MCKDF['NDC/UPC'].astype(str).str[-2:]
    MCKDF['Net']=MCKDF['Net'].apply(float)
    MCK_sheet = MCKDF[['NDC','Item Description', 'Net']]
    #print(MCKDF.head())
    MCKDF = MCKDF.rename(columns={'Net': 'MCK'})
    MCKDF = MCKDF[['NDC','Item Description', 'MCK']]
    MCKDF = MCKDF[['NDC','MCK']]
    MCK_sheet.to_excel(writer, sheet_name='MCK', index=False)
else:
    kinrayRXDF = pd.DataFrame(columns=['NDC']) 


#loadscript('Creating Report sheet', 6, .6)
reportDF = reportDF.groupby(['NDC','DRGNAME','DRUGSTRONG', 'PACKAGESIZE'], as_index=False).sum()
reportDF['Total Dispense'] = reportDF['QUANT']/reportDF['PACKAGESIZE']
reportDF['Total Dispense'] = reportDF['Total Dispense'].apply(lambda x:round(x,1))
reportDF1 = pd.merge(reportDF, kinrayRXDF, on=['NDC'], how='left')
reportDF2 = pd.merge(reportDF1, kinrayOTCDF, on=['NDC'], how='left')
reportDF = pd.merge(reportDF2, MCKDF, on=['NDC'], how='left')
reportDF['Total Purchase'] = reportDF['KIN RX'] + reportDF['MCK']
reportDF['DISP'] = reportDF['Total Dispense'] - reportDF['Total Purchase']
reportDF.to_excel(writer, sheet_name='Report', index=False)

writer.save()
os.system("open -a 'Microsoft Excel.app' 'TrialReport.xlsx'")
# Windows - os.system('start excel.exe "%s\\TrialReport.xls"' % (sys.path[0], ))


