#!/usr/bin/python
IS_WIN = True
import sys, csv, time, os, pprint, traceback
from Tkinter import Tk
import Tkinter
from tkFileDialog import askopenfilename
if IS_WIN:
    import win32com.client
    import win32clipboard

    shell = win32com.client.Dispatch("WScript.Shell")

t = time.localtime()
formated_date = time.strftime('%m%d%y', t)

directory = formated_date
if not os.path.exists(directory):
    os.makedirs(directory)

def get_required_file(message):
    root = Tkinter.Tk()
    root.withdraw()
    root.overrideredirect(True)
    root.geometry('0x0+0+0')
    root.deiconify()
    root.lift()
    root.focus_force()
    input_filename = askopenfilename(title=message)
    root.destroy()

    if not input_filename:
        print "Missing required file, exiting."
        sys.exit()
    return input_filename

def read_product_cost():
    try:
        input_filename = get_required_file("Choose Agris csv cost file")
        #input_filename = 'Misc Product Cost 8.30.13.csv'
        product_cost = {}
        with open(input_filename, 'rbU') as csvfile:
            reader = csv.reader(csvfile, delimiter=',', quotechar='|')
            for row in reader:
                if row[0] != 'ITE' and row[0] != 'LOC':
                    loc = row[0].replace('\xff', '').replace('\xa0','').strip()
                    item_number = row[3].replace('\xff', '').replace('\xa0', '')
                    if loc not in product_cost:
                        product_cost[loc] = {}
                    product_cost[loc][item_number] = float(row[10])
        return product_cost
    except Exception, e:
        print "Error encountered in read_product_cost: " + str(e)
        print traceback.format_exc()
        raw_input("Press enter to quit")
        sys.exit()

def convert_ingredient_list():
    try:
        product_cost = read_product_cost()
        plant_file_550 = get_required_file("Choose Brill ingredient list for Adams Center (550 / 64)")
        while(raw_input("\r\nAre you sure this file is correct for Adams Center (550 / 64)? [y/n]\r\n"+plant_file_550+" ").lower() != 'y'):
            plant_file_550 = get_required_file("Choose Brill ingredient list for Adams Center (550 / 64)")


        plant_file_560 = get_required_file("Choose Brill ingredient list for Augusta (560 / 61)")
        while(raw_input("\r\nAre you sure this file is correct for Augusta (560 / 61)? [y/n\r\n"+plant_file_560+" ").lower() != 'y'):
            plant_file_560 = get_required_file("Choose Brill ingredient list for Augusta (560 / 61)")

        plant_file_570 = get_required_file("Choose Brill ingredient list for Brandon (570 / 68)")
        while(raw_input("\r\nAre you sure this file is correct for Brandon (570 / 68)? [y/n]\r\n"+plant_file_570+" ").lower() != 'y'):
            plant_file_570 = get_required_file("Choose Brill ingredient list for Brandon (570 / 68)")

        plant_file_580 = get_required_file("Choose Brill ingredient list for Sangerfield (580 / 66)")
        while(raw_input("\r\nAre you sure this file is correct for Sangerfield (580 / 66)? [y/n]\r\n"+plant_file_580+" ").lower() != 'y'):
            plant_file_580 = get_required_file("Choose Brill ingredient list for Sangerfield (580 / 66)")

        plant_files = [
            [plant_file_550, '550'],
            [plant_file_560, '560'],
            [plant_file_570, '570'],
            [plant_file_580, '580']
        ]
        f = open(directory+os.sep+'cost_output_'+formated_date+'.xf1','wb')
        exception_file = directory+os.sep+'cost_exception_report_'+formated_date+'.txt'
        exception_fh = open(exception_file,'wb')
        counts_by_plant = {}
        for plant_file in plant_files:
            filepath = plant_file[0]
            with open(filepath, 'rb') as csvfile:
                reader = csv.reader(csvfile, delimiter='\t', quotechar='|')
                next(reader)
                loc = plant_file[1]
                counts_by_plant[loc] = 0
                for row in reader:
                    product_number = row[0].strip()
                    if product_number != 'Code':
                        if product_number in product_cost[loc]:
                            cost_ton = product_cost[loc][product_number]
                            cost = "%.4f" % round(cost_ton/20 ,4)
                            f.write('IP'+loc)
                            f.write(product_number.rjust(10))
                            f.write(cost.rjust(12))
                            f.write('\r\n')
                            counts_by_plant[loc] = counts_by_plant[loc] + 1
                        else:
                            exception_fh.write("Cost not found for " + product_number + " (" + row[1] + ") @ " + loc)
                            exception_fh.write('\r\n')

            print "Exception report written to %s" % exception_file
        pprint.pprint(counts_by_plant)
    except Exception, e:
        print "Error encountered in convert_ingredient_list: " + str(e)
        print traceback.format_exc()
        raw_input("Press enter to quit")
        sys.exit()

def convert_to_xf1(product_code_prefix, plant_number, plant_name):
    try:
        input_filename = get_required_file("Press enter to choose SS ingredient cost file for "+plant_name+" (" + plant_number + " / " + product_code_prefix + ")")
        #input_filename = 'BRANDON TEST MO 09.02.13.TXT'

        output_filename = 'Master'+formated_date+'Ing'+product_code_prefix+'.xf1'

        prefix_exclusions = {}
        with open('exclusions.txt', 'rb') as exclusionsfile:
            reader = csv.reader(exclusionsfile, delimiter=',', quotechar='|')
            for row in reader:
                if row[0]:
                    prefix_exclusions[row[0]] = True

        product_codes_sorted = []
        product_codes = {}
        with open(input_filename, 'rb') as csvfile:
            f = open(directory+os.sep+output_filename,'wb')
            reader = csv.reader(csvfile, delimiter='\t', quotechar='|')
            f.write('XF Version = 5\r\n')
            next(reader)
            next(reader)
            plant = next(reader)

            if plant[0].lower() != plant_name.lower():
                print "%s is for %s, %s expected" % (input_filename, plant[0], plant_name)
                convert_to_xf1(product_code_prefix, plant_number, plant_name)
                return false

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
                    else:
                        print "Skipping prefix for " + product_code
                    product_codes[product_code] = True
                    product_codes_sorted.append(product_code)
                    f.write((product_code).rjust(12))
                    price_ton = float(row[2].strip())
                    price = "%.4f" % round(price_ton/20 ,4)
                    f.write(price.rjust(12))
                    f.write(product_code.rjust(128))
                    f.write('\r\n')
                    print row
            product_codes_sorted.sort()
        if IS_WIN:
            ingredient_selection(product_codes, product_codes_sorted)
        return True
    except Exception, e:
        print "Error encountered in convert_to_xf1: " + str(e)
        print traceback.format_exc()
        raw_input("Press enter to quit")
        sys.exit()

def copy_and_get_clipboard_data():
    shell.SendKeys("^c", 0)
    time.sleep(0.1)
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    return data

def ingredient_selection(product_codes, product_codes_sorted):
    try:
        print "You will have 5 seconds to select the code column of a row on the export screen."
        input = raw_input("Press enter when ready for ingredient selection")

        if input.lower() != 'n':
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
    except Exception, e:
        print "Error encountered in ingredient_selection: " + str(e)
        print traceback.format_exc()
        raw_input("Press enter to quit")
        sys.exit()

monthly = raw_input("Is this a monthly update? [y/n]: ")
if monthly.lower() == 'y':
    convert_ingredient_list()
convert_to_xf1('61', '560', 'Augusta Mill')
convert_to_xf1('64', '550', 'Adams Center')
convert_to_xf1('66', '580', 'Sangerfield Mill')
convert_to_xf1('68', '570', 'Brandon Mill')
raw_input("Press enter to quit")
