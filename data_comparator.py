import pandas as pd
import os
from pandas import ExcelWriter

def main():
    input_file = os.path.abspath(str(input("Enter the path: ")))
    xes_name = str(input("Enter xlsx file name to be generated: "))
    xes_name += ".xlsx"

    df1 = pd.read_excel(open(input_file,'rb'), sheetname='VSheet',index=False)

    df2 = pd.read_excel(open(input_file,'rb'), sheetname='RSheet',index=False)

    vSheet = [x for x in df1.itertuples()]
    rSheetDict = dict(zip(df2.keys(), df2.get_values()[0]))

    rSheetSet = set([x for x in rSheetDict])
    rSheetSet.discard('Currency')

    outputDict = {}

    for x in vSheet:
        if (x[1] in rSheetSet):
            val = abs(x[2]-rSheetDict[x[1]])
            if (val != 0):
                outputDict[x[1]] = val
            elif val == 0:
                outputDict[x[1]] = "OK"

    finalList = [(x,outputDict[x]) for x in outputDict]

    dfS1 = pd.DataFrame(data=finalList, columns=['Currency', 'Absolute Difference'])
    #writer = pd.ExcelWriter(xes_name, engine='xlsxwriter')
    writer = ExcelWriter(xes_name)

    #dfS1 = dfS1.sort_values(by=['Currency'], ascending=False)

    dfS1.to_excel(writer, index=False, sheet_name='Result')

    writer.save()

if __name__ == '__main__':
    try:
        main()
        while True:
            inp = str(input("Do you want to continue(Y/N)?"))
            if inp != None and (inp.lower() != "Y".lower()):
                break
            else:
                main()
    except Exception as err:
        print('An exception happened ' + str(err))
