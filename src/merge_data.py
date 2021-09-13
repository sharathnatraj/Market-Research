import pandas as pd
import numpy as np
import os

if __name__ == '__main__':
    datadirectory = "./../data/"
    isFirstFile = True
    for filename in os.listdir(datadirectory):
        if filename.endswith(".xlsx"):
            xls = pd.ExcelFile(os.path.join(datadirectory, filename), engine='openpyxl')
            print("file " + filename)
            df = pd.read_excel(xls, 'Data Sheet', header=None)
            comp = df.iloc[0][1]

            # Profit and loss
            idx_st = np.where(df[0] == 'PROFIT & LOSS')[0][0]
            idx_ed = np.where(df[0] == 'Quarters')[0][0]
            dfpl = df.iloc[idx_st:idx_ed]
            dfpl = dfpl.dropna(axis=1, how="all")
            dfpl = dfpl.dropna().T
            dfpl.columns = dfpl.iloc[0]
            dfpl = dfpl[1:]  # take the data less the header row
            dfpl['Company'] = comp

            # Balance sheet
            idx_st = np.where(df[0] == 'BALANCE SHEET')[0][0]
            idx_ed = np.where(df[0] == 'CASH FLOW:')[0][0]
            dfbl = df.iloc[idx_st:idx_ed]
            dfbl = dfbl.dropna(axis=1, how="all")
            dfbl = dfbl.dropna().T
            dfbl.columns = dfbl.iloc[0]
            dfbl = dfbl[1:]

            # Cashflow
            idx_st = np.where(df[0] == 'CASH FLOW:')[0][0]
            idx_ed = np.where(df[0] == 'DERIVED:')[0][0]
            dfcf = df.iloc[idx_st:idx_ed]
            dfcf = dfcf.dropna(axis=1, how="all")
            dfcf = dfcf.dropna().T
            dfcf.columns = dfcf.iloc[0]
            dfcf = dfcf[1:]
            dfpl.merge(dfbl).merge(dfcf).to_csv('./../data/consolidated_data.csv', mode='a', index=False,
                                                header=isFirstFile)
            isFirstFile = False
