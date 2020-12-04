# import packages
import time
import gspread
from datetime import datetime, timedelta
import pandas as pd
from gspread_pandas import Spread
from oauth2client.service_account import ServiceAccountCredentials

startTime = time.time()

# configure gspread google API's
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('Credentials.json', scope) # Your json file here
gc = gspread.authorize(credentials)

# get yesterday's date
yesterday = datetime.now() - timedelta(days = 1)
dateYestXL = yesterday.strftime('%Y-%m-%d') # "yyyy-mm-dd" format
dateYestGS = yesterday.strftime('%d %b %Y').lstrip('0').replace(' 0', ' ') # "d mmm yyyy" format
dateYestSFX = yesterday.strftime('%d-%b-%Y') # "dd-mmm-yyyy" format
dateYestBDFR = yesterday.strftime('%d/%m/%Y') # "dd/mm/yyyy" format

def sapDigital():
    try:
        # opening the source file with pandas to a dataframe -- pandas
        daily = pd.read_excel(r'C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/SAP/Daily Orders Report/SAP Daily Orders (B&QTPCAFR).xlsx')

        # filter 'Digital Sales' excel row by previous day and retailer -- pandas
        TPyest = daily[(daily['Calendar Date'] == dateYestXL) & (daily['Order Creation Site Label'] == 'TRADEPOINT WEBSITE')]
        DIYyest = daily[(daily['Calendar Date'] == dateYestXL) & (daily['Order Creation Site Label'] == 'DIY.COM')]
        CAFRyest = daily[(daily['Calendar Date'] == dateYestXL) & (daily['Order Creation Site Label'] == 'Castorama.fr site Web')]

        # slice filtered row ready for transfer -- pandas
        # 'digital sales'
        tp = TPyest.loc[:, ['Calendar Date', 'Realised Orders', 'Realised Sales', 'Realised AOV', 'Cash Sales']]
        diy = DIYyest.loc[:, ['Calendar Date', 'Realised Orders', 'Realised Sales', 'Realised AOV', 'Cash Sales']]
        cafr = CAFRyest.loc[:, ['Calendar Date', 'Realised Orders', 'Realised Sales', 'Realised AOV', 'Cash Sales']]

        # create dataframe
        # 'digital sales'
        dfTP = pd.DataFrame(tp)
        dfDIY = pd.DataFrame(diy)
        dfCAFR = pd.DataFrame(cafr)

        # update dataframe index as 'Calendar Date'
        # 'digital sales'
        dfTP.set_index('Calendar Date', inplace=True)
        dfDIY.set_index('Calendar Date', inplace=True)
        dfCAFR.set_index('Calendar Date', inplace=True)

        # opening destination file & sheets with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsTP = gsDaily.worksheet('TP')
        gsDIY = gsDaily.worksheet('B&Q')
        gsCAFR = gsDaily.worksheet('CAFR')

        # opening destination file with gspread-pandas
        TP = Spread('Daily Realised Sales Revenue', 'TP')
        DIY = Spread('Daily Realised Sales Revenue', 'B&Q')
        CAFR = Spread('Daily Realised Sales Revenue', 'CAFR')

        # locate yesterday's date in gspread and parse row-number for reference cell
        cellTP = gsTP.find(dateYestGS)
        row_numDigTP = 'D'+(str(cellTP.row))

        cellDIY = gsTP.find(dateYestGS)
        row_numDigDIY = 'D'+(str(cellDIY.row))

        cellCAFR = gsTP.find(dateYestGS)
        row_numDigCAFR = 'D'+(str(cellCAFR.row))
        
        # append slice to google sheet
        # 'digital sales'
        TP.df_to_sheet(dfTP, index=False, headers=False, sheet='TP', start=row_numDigTP, replace=False)
        DIY.df_to_sheet(dfDIY, index=False, headers=False, sheet='B&Q', start=row_numDigDIY, replace=False)
        CAFR.df_to_sheet(dfCAFR, index=False, headers=False, sheet='CAFR', start=row_numDigCAFR, replace=False)

    except Exception as ex:
        print("** SAP Digital failed: ", ex)


def sapTotal():
    try:
        # opening the source file with pandas to a dataframe -- pandas
        totalGBP = pd.read_excel(r'C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/SAP/Daily Orders Report/Total Daily GBP/Store Sales Performance - Daily GBP.xlsx')
        totalEUR = pd.read_excel(r'C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/SAP/Daily Orders Report/Total Daily EUR/Store Sales Performance - Daily EU.xlsx')

        # filter 'Total Sales' excel row by previous day and retailer -- pandas
        totalYestGBP = totalGBP[(totalGBP['Calendar Date'] == dateYestXL)]
        totalYestEUR = totalEUR[(totalEUR['Calendar Date'] == dateYestXL)]

        # slice filtered row ready for transfer -- pandas
        # 'total sales'
        totalTP = totalYestGBP.loc[:, ['Calendar Date', 'Total Trade Sales GBP']]
        totalBQ = totalYestGBP.loc[:, ['Calendar Date', 'Total B&Q Sales GBP']]
        totalCAFR = totalYestEUR.loc[:, ['Calendar Date', 'Total CAFR Sales EU Inc VAT']]

        # create dataframe
        # 'total sales'
        dfTotalTP = pd.DataFrame(totalTP, columns = ['Calendar Date', 'Total Trade Sales GBP'])
        dfTotalBQ = pd.DataFrame(totalBQ, columns = ['Calendar Date', 'Total B&Q Sales GBP'])
        dfTotalCAFR = pd.DataFrame(totalCAFR, columns = ['Calendar Date', 'Total CAFR Sales EU Inc VAT'])

        # update dataframe index as 'Calendar Date'
        # 'total sales'
        dfTotalTP.set_index('Calendar Date', inplace=True)
        dfTotalBQ.set_index('Calendar Date', inplace=True)
        dfTotalCAFR.set_index('Calendar Date', inplace=True)

        # opening destination file & sheets with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsTP = gsDaily.worksheet('TP')
        gsDIY = gsDaily.worksheet('B&Q')
        gsCAFR = gsDaily.worksheet('CAFR')

        # opening destination file with gspread-pandas
        TP = Spread('Daily Realised Sales Revenue', 'TP')
        DIY = Spread('Daily Realised Sales Revenue', 'B&Q')
        CAFR = Spread('Daily Realised Sales Revenue', 'CAFR')

        # locate yesterday's date in gspread and parse row-number for reference cell
        cellTP = gsTP.find(dateYestGS)
        row_numTotTP = 'I'+(str(cellTP.row))

        cellDIY = gsDIY.find(dateYestGS)
        row_numTotDIY = 'I'+(str(cellDIY.row))

        cellCAFR = gsCAFR.find(dateYestGS)
        row_numTotCAFR = 'N'+(str(cellCAFR.row))

        # append slice to google sheet
        # 'total sales'
        TP.df_to_sheet(dfTotalTP, index=False, headers=False, sheet='TP', start=row_numTotTP, replace=False)
        DIY.df_to_sheet(dfTotalBQ, index=False, headers=False, sheet='B&Q', start=row_numTotDIY, replace=False)
        CAFR.df_to_sheet(dfTotalCAFR, index=False, headers=False, sheet='CAFR', start=row_numTotCAFR, replace=False)
    except Exception as ex:
        print("** SAP Total failed: ", ex)
        

def sfxWeb():
    try:
        # opening the source file with pandas to a dataframe -- pandas
        dailyXL = pd.read_excel('C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/SFX/Daily Sales/Daily COVID19 submission.xlsx')

        # Create DataFrame
        df = pd.DataFrame(dailyXL)

        # update dataframe index as 'Calendar Date'
        df.set_index('Date', inplace=True)
        
        # fill merged date cells
        df.index = pd.Series(df.index).fillna(method='ffill')
     
        # filter 'Web Sales' excel row by previous day and 'Origin Channel' -- pandas
        SFXyestWeb = df[(df.index == dateYestSFX) & (df['Origin Channel'] == 'WEB')]

        # slice filtered row ready for transfer -- pandas
        # 'digital sales'
        sfxW = SFXyestWeb.loc[:, ['# Sales Orders', 'Sales Net excl VAT £', 'AOV £', 'Demand Net excl VAT £']]

        # create datafram
        dfSFX = pd.DataFrame(sfxW)

        # opening destination file & sheets with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsSFX = gsDaily.worksheet('SFX')

        # opening destination file with gspread-pandas
        SFX = Spread('Daily Realised Sales Revenue', 'SFX')

        # locate yesterday's date in gspread and parse row-number for reference cell
        
        cellSFX = gsSFX.find(dateYestGS)
        row_numDigSFX = 'C'+(str(cellSFX.row))

        # append slice to google sheet
        # 'digital sales'
        SFX.df_to_sheet(dfSFX, index=False, headers=False, sheet='SFX', start=row_numDigSFX, replace=False)
    except Exception as ex:
        print("** SFX Web failed: ", ex)
        

def sfxTotal():
    try:
        # opening the source file with pandas to a dataframe -- pandas
        dailyXL = pd.read_excel('C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/SFX/Daily Sales/Daily COVID19 submission.xlsx')

        # Create DataFrame
        df = pd.DataFrame(dailyXL)

        # update dataframe index as 'Calendar Date'
        df.set_index('Date', inplace=True)
        
        # fill merged date cells
        df.index = pd.Series(df.index).fillna(method='ffill')

        # slice filtered row ready for transfer -- pandas
        # 'total sales'
        SFXyest = df.loc[dateYestSFX, ['# Sales Orders', 'Sales Net excl VAT £', 'AOV £', 'Demand Net excl VAT £']].sum()

        SFXyest = pd.DataFrame(SFXyest).transpose()

        # create SFXyest dataframe
        df2 = pd.DataFrame(SFXyest)
        
        # opening destination file & sheets with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsSFX = gsDaily.worksheet('SFX')

        # opening destination file with gspread-pandas
        SFX = Spread('Daily Realised Sales Revenue', 'SFX')

        # locate yesterday's date in gspread and parse row-number for reference cell
        
        cellSFX = gsSFX.find(dateYestGS)
        row_numTotSFX = 'G'+(str(cellSFX.row))
        
        # append slice to google sheet
        # 'total sales'
        SFX.df_to_sheet(SFXyest, index=False, headers=False, sheet='SFX', start=row_numTotSFX, replace=False)
    except Exception as ex:
        print("** SFX Total failed: ", ex)
        

def sfxTotalAOV():
    try:
        # opening the source file with pandas to a dataframe -- pandas
        dailyXL = pd.read_excel('C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/SFX/Daily Sales/Daily COVID19 submission.xlsx')

        # Create DataFrame
        df = pd.DataFrame(dailyXL)

        # update dataframe index as 'Calendar Date'
        df.set_index('Date', inplace=True)
        
        # fill merged date cells
        df.index = pd.Series(df.index).fillna(method='ffill')

        
        # slice filtered row ready for transfer -- pandas
        # 'total sales'
        SFXyest = df.loc[dateYestSFX, ['Sales Net excl VAT £', '# Sales Orders']].sum()
        aov = df.loc[dateYestSFX, 'AOV £'] = SFXyest['Sales Net excl VAT £'] / SFXyest['# Sales Orders']

        # opening destination file & sheets with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsSFX = gsDaily.worksheet('SFX')

        # opening destination file with gspread-pandas
        SFX = Spread('Daily Realised Sales Revenue', 'SFX')

        # locate yesterday's date in gspread and parse row-number for reference cell
        cellSFX = gsSFX.find(dateYestGS)
        
        # update 'total AOV' -- gspread

        headers = gsSFX.row_values(1)
        colToUpdate = headers.index('SFX Total AOV')+1
        cellToUpdate = gsSFX.cell(cellSFX.row, colToUpdate)
        aovTotal = cellToUpdate.value = aov
        cell_list = []
        cell_list.append(cellToUpdate)
        gsSFX.update_cells(cell_list)
    except Exception as ex:
        print("** SFX Total AOV failed: ", ex)


def bdfr():
    try:
        # opening the source file with pandas to a dataframe -- pandas
        dailyXL = pd.read_excel('C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/BDFR/Daily Sales/Daily DATA BD.xlsx')

        # create dataframe
        df = pd.DataFrame(dailyXL)
         
        # set header -- pandas
        dailyXL.columns = dailyXL.iloc[1]
        dailyXL = dailyXL.drop([0,1])
        dailyXL = dailyXL.reset_index(drop=True)

        # convert 'Date' column to a datetime series -- pandas
        dailyXL['Date'] = pd.to_datetime(dailyXL['Date'])

        # filter 'Sales' excel row by previous day -- pandas
        bdfrYest = dailyXL[(dailyXL['Date'] == dateYestXL)]
            
        
        # slice filtered row ready for transfer -- pandas
        # first slice
        bdfr_1 = bdfrYest.loc[:, ['Date', 'Nombre de factures WEB', 'CA HT WEB']]

        # create 1st dataframe
        dfBDFR_1 = pd.DataFrame(bdfr_1)

        #update dataframe index as 'Date'
        dfBDFR_1.set_index('Date', inplace=True)

        # opening destination file & sheet with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsBDFR = gsDaily.worksheet('BDFR')

        # opening destination file with gspread-pandas
        BDFR = Spread('Daily Realised Sales Revenue', 'BDFR')

        # locate yesterday's date in gspread and parse row-number for reference cell
        cellBDFR = gsBDFR.find(dateYestGS)
        row_numBDFR = 'D'+(str(cellBDFR.row))

        # append slice to google sheet
        BDFR.df_to_sheet(dfBDFR_1, index=False, headers=False, sheet='BDFR', start=row_numBDFR, replace=False)


        # slice filtered row ready for transfer -- pandas
        # second slice
        bdfr_2 = bdfrYest.loc[:, ['Date', 'Commandé WEB', 'Passages caisse', 'CA HT']]

        # create 1st dataframe
        dfBDFR_2 = pd.DataFrame(bdfr_2)

        #update dataframe index as 'Date'
        dfBDFR_2.set_index('Date', inplace=True)

        # opening destination file & sheet with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsBDFR = gsDaily.worksheet('BDFR')

        # opening destination file with gspread-pandas
        BDFR = Spread('Daily Realised Sales Revenue', 'BDFR')

        # locate yesterday's date in gspread and parse row-number for reference cell
        cellBDFR = gsBDFR.find(dateYestGS)
        row_numBDFR = 'G'+(str(cellBDFR.row))

        # append slice to google sheet
        BDFR.df_to_sheet(dfBDFR_2, index=False, headers=False, sheet='BDFR', start=row_numBDFR, replace=False)


        # slice filtered row ready for transfer -- pandas
        # third slice
        bdfr_3 = bdfrYest.loc[:, ['Date', 'CC transactions', 'HD transactions',
                                'CC cash sales revenue', 'HD cash sales revenue']]
        # create 1st dataframe
        dfBDFR_3 = pd.DataFrame(bdfr_3)

        #update dataframe index as 'Date'
        dfBDFR_3.set_index('Date', inplace=True)

        # opening destination file & sheet with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsBDFR = gsDaily.worksheet('BDFR')

        # opening destination file with gspread-pandas
        BDFR = Spread('Daily Realised Sales Revenue', 'BDFR')

        # locate yesterday's date in gspread and parse row-number for reference cell
        cellBDFR = gsBDFR.find(dateYestGS)
        row_numBDFR = 'Q'+(str(cellBDFR.row))

        # append slice to google sheet
        BDFR.df_to_sheet(dfBDFR_3, index=False, headers=False, sheet='BDFR', start=row_numBDFR, replace=False)

    except Exception as ex:
        print("** BDFR failed: ",ex)


def capl():
    try:
        # opening the source file with pandas to a dataframe -- pandas
        dailyXL = pd.read_excel('C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/CAPL/Daily/CPL digital  total sales-no links.xlsx')

        # set header -- pandas
        dailyXL.columns = dailyXL.iloc[0]
        dailyXL = dailyXL.drop([0,1])
        dailyXL = dailyXL.reset_index(drop=True)

        # convert 'Date' column to a datetime series -- pandas
        dailyXL['Date'] = pd.to_datetime(dailyXL['Date'])

        # filter 'Sales' excel row by previous day -- pandas
        caplYest = dailyXL[(dailyXL['Date'] == dateYestGS)]
                 
        # slice filtered row ready for transfer -- pandas
        caplDig = caplYest.iloc[:, [2, 3, 4, 3]]
        caplTot = caplYest.iloc[:, [6, 7, 8]]

        # opening destination file & sheet with gspread
        gsDaily = gc.open('Daily Realised Sales Revenue')
        gsCAPL = gsDaily.worksheet('CAPL')

        # opening destination file with gspread-pandas
        CAPL = Spread('Daily Realised Sales Revenue', 'CAPL')

        # locate yesterday's date in gspread and parse row-number for reference cell
        cellCAPL = gsCAPL.find(dateYestGS)
        row_numDig = 'C'+(str(cellCAPL.row))
        row_numTot = 'G'+(str(cellCAPL.row))


        # append slice to google sheet
        CAPL.df_to_sheet(caplDig, index=False, headers=False, sheet='CAPL', start=row_numDig, replace=False)
        CAPL.df_to_sheet(caplTot, index=False, headers=False, sheet='CAPL', start=row_numTot, replace=False)
    except Exception as ex:
        print("** CAPL failed:", ex)
            
        
sapDigital()
time.sleep(5)

sapTotal()
time.sleep(5)

sfxWeb()
time.sleep(5)

sfxTotal()
time.sleep(5)

sfxTotalAOV()
time.sleep(5)

bdfr()
time.sleep(5)

capl()
time.sleep(5)

endTime = time.time()
print('Took %s seconds to run.' % round(endTime - startTime))

"""# Retailer Google Sheet

TP = complete

BQ = complete

CAFR = complete

SFX = complete

BDFR = complete

CAPL = complete

BDRO = gc['BDRO'] # Update BDRO Daily Sales

BDES = gc['BDES'] # Update BDES Daily Sales

BDPO = gc['BDPO'] # Update BDPO Daily Sales
"""
