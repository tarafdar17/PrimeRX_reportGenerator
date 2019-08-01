import pandas as pd
from openpyxl import load_workbook
import time 


def loadscript(step, trial, duration):
    print(step, end='', flush=True)
    for i in range(trial): 
        time.sleep(duration)
        print('.', end='', flush=True)
    time.sleep(.2)
    print("\n")
    
#writer = pd.ExcelWriter('TrialReport.xlsx')

loadscript('Copying over PrimeRX', 6, .2)
df = pd.read_csv('PrimeRx.csv', low_memory=False)
#df = pd.read_csv('PrimeRx2.csv', low_memory=False)
df0 = df[['NDC','DRGNAME','DRUGSTRONG', 'PACKAGESIZE','QUANT']]
print(df.head())

loadscript('Copying over Kinray RX', 6, .1)
df1 = pd.read_excel('KINRX.xls',header= None, index= False , names=[0,1,2,3])
df1 = df1.drop(df1.index[[0,1,2,3,4,5,6]])
df1.columns = df1.iloc[0]
df1 = df1.rename(columns={'Universal NDC': 'U NDC'})
df1['NDC']=df1['U NDC'].astype(str).str[:5]+'-'+df1['U NDC'].astype(str).str[5:9]+'-'+df1['U NDC'].astype(str).str[-2:]
df1 = df1.drop(df1.index[0])

df1a = df1.groupby(['NDC'], as_index=False).sum()
df1b = df1a[['NDC', 'Drug Name', 'Qty']]
df1b = df1b.rename(columns={'Qty': 'KIN RX'})
print(df1b.head())
df1c = df1b[['NDC','KIN RX']]


loadscript('Copying over Kinray OTC', 6, .1)
df2 = pd.read_excel('KINOTC.xls',header= None, index= False , names=[0,1,2,3])
df2 = df2.drop(df2.index[[0,1,2,3,4,5,6]])
df2.columns = df2.iloc[0]
df2 = df2.rename(columns={'Universal NDC': 'U NDC'})
df2['NDC']=df2['U NDC'].astype(str).str[:5]+'-'+df2['U NDC'].astype(str).str[5:9]+'-'+df2['U NDC'].astype(str).str[-2:]
df2 = df2.drop(df2.index[0])

df2a = df2.groupby(['NDC'], as_index=False).sum()
df2b = df2a[['NDC', 'Drug Name', 'Qty']]
df2b = df2b.rename(columns={'Qty': 'KIN OTC'})
print(df2b.head())
df2c = df2b[['NDC','KIN OTC']]


loadscript('Copying over McKesson', 6, .1)
df3 = pd.read_csv('MCK.csv', header= None)
df3 = df3.drop(df3.index[[0,1,2,3,4,5,6]])
df3.columns = df3.iloc[0]
df3 = df3.drop(df3.index[0])
df3['NDC']=df3['NDC/UPC'].astype(str).str[:5]+'-'+df3['NDC/UPC'].astype(str).str[5:9]+'-'+df3['NDC/UPC'].astype(str).str[-2:]

df3a = df3[['NDC', 'Item Description', 'Net']]
df3b = df3a.groupby(['NDC', 'Item Description'], as_index=False).sum()
df3b = df3b.rename(columns={'Net': 'MCK'})
print(df3b.head())
df3c = df3b[['NDC','MCK']]


loadscript('Creating Report sheet', 6, .4)
df0 = df.groupby(['NDC','DRGNAME','DRUGSTRONG', 'PACKAGESIZE'], as_index=False).sum()
print(df0.head())

df0a = pd.merge(df0, df1c, on=['NDC'])
print(df0a.head())


loadscript('Setting up Excel Writing using Pandas DataFrames & OpenPYxl', 6, .2)
fn = 'TrialReport.xlsx'
writer = pd.ExcelWriter(fn, engine='openpyxl')
book = load_workbook(fn)
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

loadscript('Exporting to Excel Workbook: TrialReport.xlsx', 6, .6)
df.to_excel(writer, sheet_name='PrimeRxRAW', index=False)
df0a.to_excel(writer, sheet_name='PrimeRx', index=False)
df1.to_excel(writer, sheet_name='KIN_RX', index=False)
df2.to_excel(writer, sheet_name='KIN_OTC',index=False)
df3.to_excel(writer, sheet_name='MCK', index=False)
writer.save()
