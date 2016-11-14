import sys
import os.path
from openpyxl import Workbook, load_workbook#, compat.range

def main():
    if(len(sys.argv) < 3):
        print('You must provide at least 2 spreadsheets to combine')
        print('\t{} <sheet1.xlsx> <sheet2.xlsx> [sheet3.xlsx ...]'.format(sys.argv[0]))
        sys.exit(1)

    output = Workbook()
    outputWs = output.create_sheet('combined',0)
    first = True
    rowTotal = 0
    for fileName in sys.argv[1:]:
        if(not os.path.isfile(fileName)):
            print('Cannot find the file "{}". Check the name and try again'.format(fileName))
            sys.exit(2)
        wb = Workbook()
        wb = load_workbook(filename=fileName, read_only=True)
        skip=True
        for rowNum,r in enumerate(wb.active.rows, 1):
            if(skip and not first):
                skip=False
                continue
            skip=False
            for i, cell in enumerate(r, 1):
                outputWs.cell(row=rowNum+rowTotal, column=i).value = cell.value
                # print('output[{}:{}] = {}'.format(rowNum,i,outputWs.cell(row=rowNum, column=i).value))
        rowTotal += rowNum - 1
        first = False
        print(output["combined"]["A1"].value)
        output.save('output.xlsx')

if __name__ == '__main__':
    main()
