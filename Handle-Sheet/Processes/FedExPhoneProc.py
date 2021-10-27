from re import I, S
from os.path import exists
from openpyxl.worksheet import worksheet
import pandas
from openpyxl import load_workbook
import warnings
import math

#Get Data from Excel File
def get_data(filename: str):
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        return pandas.read_excel(filename, index_col=None, header = 0, engine='openpyxl')

#format DataFrame
trim_deci = lambda x: x.replace('.0', '')
def format_data(df):
    #convert to string

    df['cust_nbr'] = df['cust_nbr'].astype(str)
    df['cust_nbr'] = df['cust_nbr'].str.strip()

    df['ph_nbr_area_cd'] = df['ph_nbr_area_cd'].fillna('')
    df['ph_nbr_area_cd'] = df['ph_nbr_area_cd'].astype(str)
    df['ph_nbr_area_cd'] = df['ph_nbr_area_cd'].str.strip()
   

    df['ph_nbr'] = df['ph_nbr'].astype(str)
    df['ph_nbr'] = df['ph_nbr'].str.strip()

    #remove '.0 from' cust_nbr and ph_nbr_area_cd
    df['cust_nbr'] = df['cust_nbr'].apply(trim_deci)
    df['ph_nbr_area_cd'] = df['ph_nbr_area_cd'].apply(trim_deci)
    

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

def FedexPhoneProc(filename: str):
    print("Processing " + filename)

    df = get_data(filename)
    df = format_data(df)
    #Remove 'FedEx-Customer-Phone-Numbers' from filename
    o_trunc = filename.replace('FedEx-Customer-Phone-Numbers', '')
    #remove exty from otrunc
    o_trunc = o_trunc.replace('.xlsx', '')
    #write df to DAT
    df.to_csv('FEDEX-DEBTOR-PHONE-LIST'+o_trunc+'.DAT', sep='|', header=False, index=False)

    #remove | from outputFile
    with open('FEDEX-DEBTOR-PHONE-LIST'+ o_trunc + '.DAT', 'r') as file:
        filedata = file.read()
    filedata = filedata.replace('|', '')
    with open('FEDEX-DEBTOR-PHONE-LIST'+o_trunc+'.DAT', 'w') as file:
        file.write(filedata)
    
    print("FEDEX-DEBTOR-PHONE-LIST"+o_trunc+".DAT created")
