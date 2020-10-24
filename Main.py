import sys
import os
import pandas as pd
import xlrd
import re
import numpy as np

def getCompName(inputFileName):  # just working with name here
    workBook = xlrd.open_workbook(inputFileName)
    workSheet = workBook.sheet_by_index(0)
    inputString = str(workSheet.cell(0, 0))
    splatString = re.split('for | as', inputString)
    compName = splatString[1]
    return compName


def readFromFile(inputFile, frames):  # only working with name here
    companyName = getCompName(inputFile)
    df = pd.read_excel(inputFile, skiprows=0, header=1)
    df['Category'].replace('', np.nan, inplace=True)
    df.dropna(subset=['Category'], inplace=True)
    df = df[df.Category != 'Non ITO']
    df = df[df.Category != 'Category']
    df.insert(0, "Company", [companyName] * len(df), True)
    frames.append(df)


if __name__ == '__main__':
    directory = sys.argv[1]
    xls_files = os.listdir(directory)
    frames = []
    for file in xls_files:
        readFromFile(directory + '\\' + file, frames)
    final_table = pd.concat(frames)
    final_table.to_excel('myOutPut.xlsx')
