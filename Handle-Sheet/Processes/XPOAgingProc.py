from re import I, S
from os.path import exists
from openpyxl.worksheet import worksheet
import pandas as pd
from openpyxl import load_workbook

AGING_WITH_ADDED_COLUMNS = 'AGING-WITH-ADDED-COLUMNS.TXT'
IN_AGING_NOT_IN_SYSTEM = 'IN-AGING-NOT-IN-SYSTEM.TXT'
IN_SYSTEM_NOT_IN_AGING = 'IN-SYSTEM-NOT-IN-AGING.TXT'

#create Excel file
def createExcelFile(added_columns, not_in_aging, not_in_system):
    with pd.ExcelWriter('XPO_AGING_REPORT.xlsx') as writer:  # doctest: +SKIP
        added_columns.to_excel(writer, sheet_name='AGING_WITH_ADDED_COLUMNS', index = False)
        not_in_aging.to_excel(writer, sheet_name='IN_SYSTEM_NOT_IN_AGING', index = False)
        not_in_system.to_excel(writer, sheet_name='IN_AGING_NOT_IN_SYSTEM', index = False)


def XPOAgingProc():
    print("Processing XPO AGING Report... ")
    #read all aging files
    in_aging_not_in_system = pd.read_csv(IN_AGING_NOT_IN_SYSTEM, sep='~', header=0, index_col=False)
    in_system_not_in_aging = pd.read_csv(IN_SYSTEM_NOT_IN_AGING, sep='~', header=0, index_col=False)
    added_columns = pd.read_csv(AGING_WITH_ADDED_COLUMNS, sep='~', header=0, index_col=False)

    #create Excel file
    createExcelFile(added_columns, in_system_not_in_aging, in_aging_not_in_system)
    
    print("XPO AGING Report Processed Successfully")