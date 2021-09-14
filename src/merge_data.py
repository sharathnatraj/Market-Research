import pandas as pd
import numpy as np
import os


def extract_section(df, section_start_heading, section_end_heading):
    idx_st = np.where(df[0] == section_start_heading)[0][0]
    idx_ed = np.where(df[0] == section_end_heading)[0][0]
    dfpl = df.iloc[idx_st + 1:idx_ed - 1]
    dfpl = dfpl.dropna(how="all").dropna(axis=1, how='all').T
    dfpl.columns = dfpl.iloc[0]
    dfpl = dfpl[1:]  # take the data less the header row
    return dfpl


if __name__ == '__main__':
    datadirectory = "./../data/"
    isFirstFile = True
    columns = []
    for filename in os.listdir(datadirectory):
        if filename.endswith(".xlsx"):
            xls = pd.ExcelFile(os.path.join(datadirectory, filename), engine='openpyxl')
            print("Processing file " + filename)
            df = pd.read_excel(xls, 'Data Sheet', header=None)
            comp = df.iloc[0][1]

            # Profit and loss
            dfpl = extract_section(df, section_start_heading='PROFIT & LOSS', section_end_heading='Quarters')
            dfbl = extract_section(df, section_start_heading='BALANCE SHEET', section_end_heading='CASH FLOW:')
            dfcf = extract_section(df, section_start_heading='CASH FLOW:', section_end_heading='DERIVED:')
            dfpl['Company'] = comp

            merged_data = dfpl.merge(dfbl).merge(dfcf)
            if isFirstFile:
                columns = merged_data.columns.values
            merged_data.to_csv('./../data/consolidated_data.csv', mode='a', index=False,
                               header=isFirstFile, columns=columns)
            isFirstFile = False
