from re import I, S
from os.path import exists
from openpyxl.worksheet import worksheet
import pandas
from openpyxl import load_workbook
from Processes.Lumen import lumenProc
from Processes.PYProc import PyProc
import threading




    

def main():
    lumen = threading.Thread(target=lumenProc)
    tlumen = False

    PY = threading.Thread(target=PyProc)
    tPY = False

    if exists("LumenInput.xlsx"):
        lumen.start()
        tlumen = True

    if exists("PYInput.xlsx"):
        PY.start()
        tPY = True

    if tlumen: lumen.join()
    if tPY: PY.join()

    print("All Processes Complete")

if __name__ == "__main__":
    main()