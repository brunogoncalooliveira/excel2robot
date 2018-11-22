# coding=utf8
from openpyxl import load_workbook
import re
import argparse
from pprint import pprint



def sanitize_data(value):
    r = re.compile("\s{2,}")
    arr =  sorted( r.findall(value), key=len, reverse=True)
    for i in arr:
         value = value.replace(i, '${SPACE * ' + str(len(i)) + '}')
    #print(value)

    return value.replace("\n",'')

def SheetToTxt(sheet):
    txt = ''
    for row in sheet.rows:
        for i in row:
            if i.value is not None:
                if isinstance(i.value, int):
                    v = str(i.value).encode("utf-8").decode("utf-8")
                    txt += v + "  "
                else:
                    v = sanitize_data( i.value )
                    txt += v + "  "
            else:
                txt += "  "
        txt += "\n"
    return txt


def convert(inputfile, outputfile):

    sheet_names = ['Settings', "Variables", 'Test Cases', "Keywords"]

    wb2 = load_workbook(inputfile)
    Sheets = [i.title for i in wb2.worksheets]

    txt = ''
    for i in sheet_names:
        if i in Sheets:
            txt += '*** ' + i + " ***\n"
            txt += SheetToTxt(wb2[i])
            txt += "\n"

    fo = open(outputfile, 'w')
    fo.write(txt)
    fo.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Convert excel file to robot framework"
    )

    # Add the positional parameter
    parser.add_argument('inputfile', help="Excel input file")
    parser.add_argument('outputfile', help="Output file (in robot framework text format)")

    # Parse the arguments
    arguments = parser.parse_args()

    # Finally print the passed string
    #print(arguments.inputfile)
    #print(arguments.outputfile)

    convert(arguments.inputfile, arguments.outputfile)
