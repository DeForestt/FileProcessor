from re import I, S
from os.path import exists
import os
from openpyxl.worksheet import worksheet
import pandas
from openpyxl import load_workbook
from Processes.Lumen import lumenProc
from Processes.PYProc import PyProc
from Processes.FedExPhoneProc import FedexPhoneProc
from Processes.XPOAgingProc import XPOAgingProc
import datetime
import threading




    

def main():
    lumen = threading.Thread(target=lumenProc)
    tlumen = False

    PY = threading.Thread(target=PyProc)
    tPY = False

    fedex = threading.Thread(target=FedexPhoneProc)
    tFedEx = False

    XPOAging = threading.Thread(target=XPOAgingProc)
    tXPOAging = False

    #create current date processed folder
    if not exists('Processed'): os.mkdir('Processed')
    if not exists('Processed\\' + datetime.datetime.now().strftime("%Y-%m-%d")): os.mkdir('Processed\\' + datetime.datetime.now().strftime("%Y-%m-%d"))

    if exists("LumenInput.xlsx"):
        lumen.start()
        tlumen = True

    if exists("PYInput.xlsx"):
        PY.start()
        tPY = True
    
    if exists("FedEx-Customer-Phone-Numbers.xlsx"):
        fedex.start()
        tFedEx = True

    if exists("AGING-WITH-ADDED-COLUMNS.TXT") and exists("IN-AGING-NOT-IN-SYSTEM.TXT") and exists("IN-SYSTEM-NOT-IN-AGING.TXT"):
        XPOAging.start()
        tXPOAging = True

    if tlumen:
        lumen.join()
        #Move Lumen to processed folder
        os.rename("LumenInput.xlsx", "Processed\\" + datetime.datetime.now().strftime("%Y-%m-%d") + "\\LumenInput.xlsx")

    if tPY:
        PY.join()
        #Move PY to processed folder
        os.rename("PYInput.xlsx", "Processed\\" + datetime.datetime.now().strftime("%Y-%m-%d") + "\\PYInput.xlsx")

    if tFedEx:
        fedex.join()
        #Move FedEx to processed folder
        os.rename("FedEx-Customer-Phone-Numbers.xlsx", "Processed\\" + datetime.datetime.now().strftime("%Y-%m-%d") + "\\FedEx-Customer-Phone-Numbers.xlsx")

    if tXPOAging: 
        XPOAging.join()
        #Move XPO Aging to processed folder
        os.rename("AGING-WITH-ADDED-COLUMNS.TXT", "Processed\\" + datetime.datetime.now().strftime("%Y-%m-%d") + "\\AGING-WITH-ADDED-COLUMNS.TXT")
        os.rename("IN-AGING-NOT-IN-SYSTEM.TXT", "Processed\\" + datetime.datetime.now().strftime("%Y-%m-%d") + "\\IN-AGING-NOT-IN-SYSTEM.TXT")
        os.rename("IN-SYSTEM-NOT-IN-AGING.TXT", "Processed\\" + datetime.datetime.now().strftime("%Y-%m-%d") + "\\IN-SYSTEM-NOT-IN-AGING.TXT")

    print("All Processes Complete")

if __name__ == "__main__":
    main()
