# import packages
import gspread
from datetime import datetime, timedelta
import pandas as pd
from gspread_pandas import Spread
from oauth2client.service_account import ServiceAccountCredentials

# configure gspread google API's
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('Credentials.json', scope) # Your json file here
gc = gspread.authorize(credentials)

# get yesterday's date
yesterday = datetime.now() - timedelta(days = 1)
dateYestGS = yesterday.strftime('%d %b %Y').lstrip('0').replace(' 0', ' ') # "d mmm yyyy" format
dateYestSFX = yesterday.strftime('%d-%b-%Y') # "dd-mmm-yyyy" format



# opening the source file with pandas to a dataframe -- pandas
dailyXL = pd.read_excel('C:/Users/forbej06/OneDrive - Kingfisher PLC/email data/SFX/Daily Sales/Daily COVID19 submission.xlsx')

# Create DataFrame
df = pd.DataFrame(dailyXL)

# update dataframe index as 'Calendar Date'
df.set_index('Date', inplace=True)
    
# fill merged date cells
df.index = pd.Series(df.index).fillna(method='ffill')

def sfxWeb():
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


def sfxTotal():

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




def sfxTotalAOV():

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
    # row_numAovSFX = 'i'+(str(cellSFX.row))


    # update 'total AOV' -- gspread

    headers = gsSFX.row_values(1)
    colToUpdate = headers.index('SFX Total AOV')+1
    cellToUpdate = gsSFX.cell(cellSFX.row, colToUpdate)
    aovTotal = cellToUpdate.value = aov
    cell_list = []
    cell_list.append(cellToUpdate)
    gsSFX.update_cells(cell_list)


sfxWeb()

sfxTotal()

sfxTotalAOV()
