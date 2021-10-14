import pandas

def PyProc():
    print("Starting PY Process...")
    xl =pandas.ExcelFile("PYInput.xlsx")
    if len(xl.sheet_names) > 1:
        df = pandas.read_excel(xl, "Details", index_col=None, header=None)
    else:
        df = pandas.read_excel(xl, index_col=None, header=None)
    xl.close()

    df.to_csv(f"INPUT-FILE.DAT", sep='\t')

    df = pandas.read_csv("INPUT-FILE.DAT", sep='\t', index_col=0, header=0, converters={'0': lambda x: x[:50], '1': lambda x: x[:20], '2': lambda x: x[:20], '3': lambda x: x[:10]})

    df.to_csv(f"INPUT-FILE.DAT", sep='\t', index=False, header=False)
    print("PY Process Complete")