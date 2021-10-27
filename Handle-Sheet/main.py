from re import I, S, T
from os.path import exists
import os
from openpyxl.worksheet import worksheet
import pandas
from openpyxl import load_workbook
from Processes.Lumen import lumenProc
from Processes.PYProc import PyProc
from Processes.FedExPhoneProc import FedexPhoneProc
from Processes.XPOAgingProc import XPOAgingProc
from Processes.ENSProc import ENSProc
import datetime
import threading



#struct to hold a process name and a boolean to indicate if it is running
class Process:
    def __init__(self, file_names, running, function):
        self.file_names = file_names
        self.running = running
        self.function = function
        self.function.start()
    

def main():

    #create current date processed folder
    if not exists('Processed'): os.mkdir('Processed')
    if not exists('Processed\\' + datetime.datetime.now().strftime("%Y-%m-%d")): os.mkdir('Processed\\' + datetime.datetime.now().strftime("%Y-%m-%d"))
    
    #create list of processes
    Processes = []

    #loop through every file that starts with 'PY'
    for file in os.listdir('.'):
        if file.startswith('PY'):
            PY = threading.Thread(target=PyProc, args=(file,))
            Processes.append(Process([file], True, PY))
        if file.startswith('ENS'):
            ENS = threading.Thread(target=ENSProc, args=(file,))
            Processes.append(Process([file], True, ENS))
        if file.startswith('FedEx-Customer-Phone-Numbers'):
            FedexPhone = threading.Thread(target=FedexPhoneProc, args=(file,))
            Processes.append(Process([file], True, FedexPhone))
        if file.startswith('LumenInput'):
            lumen = threading.Thread(target=lumenProc, args=(file,))
            Processes.append(Process([file], True, lumen))

    if exists("AGING-WITH-ADDED-COLUMNS.TXT") and exists("IN-AGING-NOT-IN-SYSTEM.TXT") and exists("IN-SYSTEM-NOT-IN-AGING.TXT"):
        XPOAging = threading.Thread(target=XPOAgingProc, args=())
        Processes.append(Process(["AGING-WITH-ADDED-COLUMNS.TXT", "IN-AGING-NOT-IN-SYSTEM.TXT", "IN-SYSTEM-NOT-IN-AGING.TXT"], True, XPOAging))

    for process in Processes:
        process.function.join()
        for file in process.file_names:
            #get ext from file
            ext = file.split('.')[-1]
            #remove exte from filname
            ofile = file.split('.')[0]
            #move file to processed folder with current date
            os.rename(file, 'Processed\\' + datetime.datetime.now().strftime("%Y-%m-%d") + '\\' + ofile + '-' + datetime.datetime.now().strftime("%HH-%MM-%SS") + '.' + ext)
        
    print("All Processes Complete")

if __name__ == "__main__":
    main()
