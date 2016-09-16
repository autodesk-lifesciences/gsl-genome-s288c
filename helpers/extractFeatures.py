#!/usr/bin/python

from openpyxl import load_workbook
import sys
import json

"""
Tool to extract columns of interest from a tab delimited file of features, and write them to a .json file.
"""
def main():
    if len(sys.argv) < 4:
        print "Please supply the .xmlx filepath, sheet name and output json file as arguments"
        sys.exit(1)

    print "Filename: {}".format(sys.argv[1])
    print "Sheet Name: {}".format(sys.argv[2])
    wb = load_workbook(filename=sys.argv[1])
    ws = wb[sys.argv[2]]
    geneList = []
    geneItem= {}
    for row in ws.rows:
        # Get the systematic name (Cell index 3)
        if (row[0].value and row[3].value):
            geneItem= {}
            geneItem['systematicName'] = str(row[3].value) if row[3].value else ''
            # Get the common name (Cell index 4)
            geneItem['commonName'] = str(row[4].value) if row[4].value else ''
            # Get the description ( Cell index 15)
            geneItem['description'] = str(row[15].value) if row[15].value else ''
            geneList.append(geneItem)

    print "Writing file to {}".format(sys.argv[3])
    json.dump(geneList, open(sys.argv[3], 'w'), indent=2)


if __name__ == '__main__':
    main()
