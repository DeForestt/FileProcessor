from re import I, S
from os.path import exists
from openpyxl.worksheet import worksheet
import pandas
from openpyxl import load_workbook
from Processes.Lumen import lumenProc
from Processes.PYProc import PyProc
from Processes.FedExPhoneProc import FedexPhoneProc
from Processes.XPOAgingProc import XPOAgingProc
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

    if tlumen: lumen.join()
    if tPY: PY.join()
    if tFedEx: fedex.join()
    if tXPOAging: XPOAging.join()

    print("All Processes Complete")

if __name__ == "__main__":
    main()
