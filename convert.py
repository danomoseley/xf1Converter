#!/usr/bin/python
import sys
import csv
import time

def convert_to_xf1():
    input_filename = raw_input("Enter filename: ")
    #filename = 'Test Price Export.TXT'
    product_code_prefix = raw_input("Enter product code prefix: ")
    #product_code_prefix = '66'

    t = time.localtime()
    formated_date = time.strftime('%m%d%y', t)
    output_filename = 'Master'+formated_date+'Ing'+product_code_prefix+'.xf1'

    prefix_exclusions = {}
    with open('exclusions.txt', 'rb') as exclusionsfile:
        reader = csv.reader(exclusionsfile, delimiter=',', quotechar='|')
        for row in reader:
            if row[0]:
                prefix_exclusions[row[0]] = True

    with open(input_filename, 'rb') as csvfile:
        f = open(output_filename,'wb')
        reader = csv.reader(csvfile, delimiter='\t', quotechar='|')
        f.write('XF Version = 5\r\n')
        next(reader)
        next(reader)
        next(reader)
        next(reader)
        next(reader)
        next(reader)
        for row in reader:
            if row:
                if row[0] == 'Product' or row[0] == 'Code':
                    continue
                f.write('IP0')
                product_code = row[0].strip()
                product_code = product_code.zfill(4)
                if product_code not in prefix_exclusions:
                    product_code = product_code_prefix + product_code
                f.write((product_code).rjust(12))
                price_ton = float(row[2].strip())
                price = "%.4f" % round(price_ton/20 ,4)
                f.write(price.rjust(12))
                f.write(product_code.rjust(128))
                f.write('\r\n')

convert_to_xf1()