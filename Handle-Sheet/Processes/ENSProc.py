import pandas as pd
import os

ENS = "ENSInput.xlsx"

def ENSProc():
    print("Processing ENS File...")
    
    xl = pd.ExcelFile(ENS)
    
    #check for number of sheets
    if len(xl.sheet_names) > 1:
        df = pd.read_excel(xl, sheet_name='Details', index_col=None, header=None)
    else:
        df = pd.read_excel(xl, index_col=None, header=None)
    
    #close file
    xl.close()

    #Convert to CSV file
    df.to_csv(f"BanUpdate1.prn", index=False, header=False)

    #read CSV file
    df = pd.read_csv(f"BanUpdate1.prn", header=None, index_col=None)

    #insert columns
    df.insert(1, 'AMAL', "AMAL")
    df.insert(2, '    ', "   ")

    #Set DATE Column to user input
    df["DATE"] = input("ENS Paused:: Enter Date (YYYYMMDD): ")

    df.T.apply(lambda row: ''.join(map(str, row)))
    #remove first three rows
    df = df.drop(df.index[0:3])

    #write to CSV file
    df.to_csv(f"BanUpdate1.prn", index=False, header=False)

    #remove ',' from CSV file
    with open(f"BanUpdate1.prn", "r") as infile, open(r"BanUpdate.prn", 'w') as outfile:
        for line in infile:
            outfile.write(line.replace(",", ""))
    
    #remove infile
    os.remove(F"BanUpdate1.prn")

