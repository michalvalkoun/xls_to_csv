#!/usr/bin/env python
""" Konverze xls do csv
    @Author: Michal Valkoun
    @Date: 13/12/2021
"""

import sys
import csv

import xlrd

if len(sys.argv) != 3:
    sys.exit('Chybný počet argumentů.\nPříkaz [vstupní soubor] [výstupní soubor]')

if not sys.argv[1].endswith('.xls'):
    sys.exit('\"' + sys.argv[1] + '\" nemá příponu .xls')
if not sys.argv[2].endswith('.csv'):
    sys.exit('\"' + sys.argv[2] + '\" nemá příponu .csv')

try:
    book = xlrd.open_workbook(sys.argv[1])
except FileNotFoundError:
    sys.exit('Soubor \"' + sys.argv[1] + '\" neexistuje.')

sh = book.sheet_by_index(0)

try:
    with open(sys.argv[2], mode='w', newline='', encoding='utf8') as cs:
        wr = csv.writer(cs, delimiter=';', quoting=csv.QUOTE_MINIMAL)
        for x in range(sh.nrows):
            wr.writerow(sh.row_values(x))
except PermissionError:
    sys.exit('Nemohu otevřít \"' + sys.argv[2] + '\", zřejmě ho máš otevřený.')
