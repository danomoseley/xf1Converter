#!/usr/bin/python
import sys
import csv
import time
import win32com.client
import win32clipboard

product_codes = {}
product_codes_sorted = []
shell = win32com.client.Dispatch("WScript.Shell")

def convert_to_xf1():
    input_filename = raw_input("Enter filename: ")
    product_code_prefix = raw_input("Enter product code prefix: ")

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
                product_codes[product_code] = True
                product_codes_sorted.append(product_code)
                f.write((product_code).rjust(12))
                price_ton = float(row[2].strip())
                price = "%.4f" % round(price_ton/20 ,4)
                f.write(price.rjust(12))
                f.write(product_code.rjust(128))
                f.write('\r\n')
        product_codes_sorted.sort()

def copy_and_get_clipboard_data():
    shell.SendKeys("^c", 0)
    time.sleep(0.1)
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    return data

def ingredient_selection():
    raw_input("Press enter when ready for ingredient selection")
    time.sleep(5)

    last_data = ''
    data = copy_and_get_clipboard_data()
    shell.SendKeys("{DOWN}", 0)
    max_product_code = product_codes_sorted[-1]

    found = 0
    while data < max_product_code and data != last_data:
        last_data = data
        data = copy_and_get_clipboard_data()
        if data in product_codes:
            print "%s found" % data
            shell.SendKeys("{LEFT}", 0)
            shell.SendKeys(" ", 0)
            shell.SendKeys("{RIGHT}", 0)
            found = found + 1
        shell.SendKeys("{DOWN}", 0)
    print "%d codes in file" % len(product_codes_sorted)
    print "%d codes found and selected" % found
    raw_input("Press enter to quit")

convert_to_xf1()
ingredient_selection()
