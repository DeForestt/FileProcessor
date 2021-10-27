import pandas

def PyProc(fileName: str):
    print("Starting PY Process for" + fileName + "...")
    xl =pandas.ExcelFile(fileName)
    if len(xl.sheet_names) > 1:
        df = pandas.read_excel(xl, "Details", index_col=None, header=None)
    else:
        df = pandas.read_excel(xl, index_col=None, header=None)
    xl.close()

    #replace 'PY' in filename with ''
    fileName = fileName.replace('PY', '')
    #remove extension from filename
    fileName = fileName.replace('.xlsx', '')

    df.to_csv("INPUT-FILE-" + fileName + ".DAT", sep='\t')

    df = pandas.read_csv("INPUT-FILE-" + fileName + ".DAT", sep='\t', index_col=0, header=0, converters={'0': lambda x: x[:50], '1': lambda x: x[:20], '2': lambda x: x[:20], '3': lambda x: x[:10]})
    output_file_name ="INPUT-FILE-" + fileName + ".DAT"
    df.to_csv(output_file_name, sep='\t', index=False, header=False)
    print("PY Process Complete")