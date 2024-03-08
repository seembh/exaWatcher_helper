import tkinter as tk
from tkinter import filedialog
import re
import pandas as pd


def search_for_file_path():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()
    print(file_path)
    if len(file_path) > 0:
        return file_path
    else:
        raise ValueError('Bad file path')


def getXlsHeader(inputLog):
    temp = inputLog[0].split()
    header = temp[1:]
    header[0] = 'Snap Time'
    return header


def getNumCols(header):
    numCols = len(header)
    return numCols


def flatten_comprehension(matrix):
    return [item for row in matrix for item in row]


def findContent(inputLog):
    content = []
    for i in inputLog[1:]:
        line = []
        line.extend(re.findall('^[A-z]{3} [A-z]{3} [0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2} EET [0-9]{4}', i))
        line.extend(i.split()[6:])
        content.append(line)
    return content


def openLog(name):
    with open(name) as f:
        lines = f.readlines()
        return lines


def writeToXlsx(content, header, numCols):
    df = pd.DataFrame(content, columns=header)
    writer = pd.ExcelWriter('report.xlsx',
                            # engine='openpyxl'
                            engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_numbers': True}}
                            )
    df.to_excel(writer, sheet_name='Data', index=False)
    # workbook = writer.book
    worksheet = writer.sheets['Data']

    crRow = 0
    crCol = numCols + 2
    for i in range(1, numCols):
        worksheet.write(crRow, crCol + i - 1, header[i])

    crRow = 1
    worksheet.write(crRow, crCol - 1, 'Min')
    worksheet.write_formula(crRow, crCol, '=MIN(B:B)')

    crRow = 2
    worksheet.write(crRow, crCol - 1, 'Time of Min')
    worksheet.write_formula(crRow, crCol, '=INDEX($A:$A,MATCH(MIN(B:B),B:B,0))')

    crRow = 3
    worksheet.write(crRow, crCol - 1, 'Max')
    worksheet.write_formula(crRow, crCol, '=MAX(B:B)')

    crRow = 4
    worksheet.write(crRow, crCol - 1, 'Time of Max')
    worksheet.write_formula(crRow, crCol, '=INDEX($A:$A,MATCH(MAX(B:B),B:B,0))')

    crRow = 5
    worksheet.write(crRow, crCol - 1, 'Average')
    worksheet.write_formula(crRow, crCol, '=AVERAGE(B:B)')

    writer.close()


def output():
    # Get file path
    file_path_variable = search_for_file_path()
    # Open file for reading
    inputLog = openLog(file_path_variable)

    header = getXlsHeader(inputLog)
    numCols = getNumCols(header)
    content = findContent(inputLog)
    writeToXlsx(content, header, numCols)
    # createBiutifulFormulas(content, header, numCols)


if __name__ == '__main__':
    output()
