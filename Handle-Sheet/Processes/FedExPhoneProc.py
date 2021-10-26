from re import I, S
from os.path import exists
from openpyxl.worksheet import worksheet
import pandas
from openpyxl import load_workbook
import warnings

FILENAME = 'FedEx-Customer-Phone-Numbers.xlsx'

#Get Data from Excel File
def get_data(filename):
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        return pandas.read_excel(filename, index_col=None, header = 0, engine='openpyxl')

#format DataFrame
def format_data(df):
    #convert to string
    df['cust_nbr'] = df['cust_nbr'].astype(str)
    df['cust_nbr'] = df['cust_nbr'].str.strip()
    df['ph_nbr_area_cd'] = df['ph_nbr_area_cd'].astype(str)
    df['ph_nbr_area_cd'] = df['ph_nbr_area_cd'].str.strip()
    df['ph_nbr'] = df['ph_nbr'].astype(str)
    df['ph_nbr'] = df['ph_nbr'].str.strip()

    #trim last two characters
    df['ph_nbr_area_cd'] = df['ph_nbr_area_cd'].str[:-2]


    #drop last column
    df.drop('ph_ext_nbr', axis=1, inplace=True)

    #apply padding to dataframe
    df['cust_nbr'] = df['cust_nbr'].apply(lambda x: pad_string(x, 12))
    
    #apply quality check
    quality_check(df)
    return df

#Pad string with spaces
def pad_string(string, length):
    return string.ljust(length, ' ')


#Quality Check
def quality_check(df):
    #check row width
    for row in df.itertuples():
        ph_nbr_len = len(row.ph_nbr)
        area_cd_len = len(row.ph_nbr_area_cd)
        if ph_nbr_len != 7:
            print("Row: " + str(row) + " phone number is not 7 characters long")
        if area_cd_len != 3:
            print("Row: " + str(row) + " area code is not 3 characters long")

def FedexPhoneProc():
    df = get_data(FILENAME)
    df = format_data(df)

    #write df to DAT
    df.to_csv('FEDEX-DEBTOR-PHONE-LIST.DAT', sep='|', header=False, index=False)

    #remove | from outputFile
    with open('FEDEX-DEBTOR-PHONE-LIST.DAT', 'r') as file:
        filedata = file.read()
    filedata = filedata.replace('|', '')
    with open('FEDEX-DEBTOR-PHONE-LIST.DAT', 'w') as file:
        file.write(filedata)
    print(df)
