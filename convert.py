#!/usr/bin/python
import sys, csv, time, os, traceback, collections, platform
import subprocess
from Tkinter import Tk
import Tkinter
from tkFileDialog import askopenfilename
IS_WIN = platform.system().lower() == 'windows'

print 'Checking for latest version'
p = subprocess.Popen(["git", "pull"], shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
out, err = p.communicate()
if not 'Already up-to-date.' in out:
    print 'Version updated, restarting'
    python = sys.executable
    os.execl(python, python, * sys.argv)
else:
    print 'Up to date'

if IS_WIN:
    import win32com.client
    import win32clipboard

    shell = win32com.client.Dispatch("WScript.Shell")

t = time.localtime()
formated_date = time.strftime('%m%d%y', t)

directory = formated_date
if not os.path.exists(directory):
    os.makedirs(directory)

def get_required_file(message, filetypes=None):
    root = Tkinter.Tk()
    root.withdraw()
    root.overrideredirect(True)
    root.geometry('0x0+0+0')
    root.deiconify()
    root.lift()
    root.focus_force()
    print "Choose file for %s" % message 
    input_filename = askopenfilename(title=message)
    root.destroy()

    if not input_filename:
        input = raw_input("Missing required file, try again? [y/n]: ")
        if input.lower() == 'n':
            sys.exit()
        else:
            return get_required_file(message)
    return input_filename

def get_required_file_confirm(description, filetypes=None):
    input_file = get_required_file(description)

    prompt = "\r\nAre you sure this file is correct for %s? [y/n]\r\n%s " % (description, input_file)

    while(raw_input(prompt).lower() not in ['y', '']):
        input_file = get_required_file(description, filetypes)
    return input_file

def read_product_cost():
    try:
        input_filename = get_required_file_confirm("Agris csv cost file", filetypes=(("CSV", "*.csv"),
                                           ("All files", "*.*") ))
        product_cost = {}
        with open(input_filename, 'rbU') as csvfile:
            reader = csv.reader(csvfile, delimiter=',', quotechar='"')
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

def convert_cost_list(product_costs):
    cost_list_files = []
    cost_list_files.append(get_required_file_confirm("Misc Ingredient list for NY"))
    cost_list_files.append(get_required_file_confirm("Misc Ingredient list for NE"))

    products = {}

    for cost_list_file in cost_list_files:
        with open(cost_list_file, 'rb') as csvfile:
            reader = csv.reader(csvfile, delimiter='\t', quotechar='"')
            for i in range(5):
                next(reader)
            for row in reader:
                if row:
                    if row[0].lower() == 'product' or row[0].lower() == 'code' or len(row) == 1:
                        continue
                    full_product_code = row[0].strip()
                    product_description = row[1].strip()
                    product_size = int(row[2])

                    product_code_parts = full_product_code.split('-')
                    product_code = full_product_code
                    product_cost = 'N/A'

                    if product_code_parts and len(product_code_parts) > 1:
                        plant_letter = product_code_parts[0].strip()
                        product_code = product_code_parts[1]
                        if plant_letter.lower() == 's':
                            plant_number = '580'
                        elif plant_letter.lower() == 'a':
                            plant_number = '560'
                        if plant_number:
                            if product_code in product_costs[plant_number]:
                                product_cost = product_costs[plant_number][product_code]

                    products[full_product_code] = {
                        'full_code': full_product_code,
                        'code': product_code,
                        'description': product_description,
                        'cost': product_cost,
                        'size': product_size
                    }

    ordered_products = collections.OrderedDict(sorted(products.items()))

    cost_report_filename = directory+os.sep+'misc_ingredient_report_'+formated_date+'.csv'

    with open(cost_report_filename, 'wb') as cost_report_fh:
        csvwriter = csv.writer(cost_report_fh, delimiter=',',
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
        csvwriter.writerow(['Product Code', 'Product Description', 'Size', 'Cost', 'Extended Cost'])

        for k, product in ordered_products.iteritems():
            if product['cost'] != 'N/A':
                extended_product_cost = "%.4f" % round(((float(product['cost'])/2000)*product_size) ,4)
            else:
                extended_product_cost = 'N/A'
            csvwriter.writerow([product['full_code'], product['description'], product['size'], product['cost'], extended_product_cost])

    print "\r\nCost report written to %s" % cost_report_filename 

def convert_ingredient_list():
    try:
        product_cost = read_product_cost()
        convert_cost_list(product_cost)
        
        plant_file_550 = get_required_file_confirm("Brill ingredient list for Adams Center (550 / 64)")
        plant_file_560 = get_required_file_confirm("Brill ingredient list for Augusta (560 / 61)")
        plant_file_570 = get_required_file_confirm("Brill ingredient list for Brandon (570 / 68)")
        plant_file_580 = get_required_file_confirm("Brill ingredient list for Sangerfield (580 / 66)")

        plant_files = [
            [plant_file_550, '550'],
            [plant_file_560, '560'],
            [plant_file_570, '570'],
            [plant_file_580, '580']
        ]
        output_filename = directory+os.sep+'cost_output_'+formated_date+'.xf1'
        f = open(output_filename,'wb')
        exception_file = directory+os.sep+'cost_exception_report_'+formated_date+'.txt'
        exception_fh = open(exception_file,'wb')
        counts_by_plant = {}
        for plant_file in plant_files:
            filepath = plant_file[0]
            with open(filepath, 'rb') as csvfile:
                reader = csv.reader(csvfile, delimiter='\t', quotechar='"')
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

        print "\r\nException report written to %s" % exception_file
        print "\r\nCost output written to %s" % output_filename
        print "%d costs for Adams Center (550 / 64)" % counts_by_plant['550']
        print "%d costs for Augusta (560 / 61)" % counts_by_plant['560']
        print "%d costs for Brandon (570 / 68)" % counts_by_plant['570']
        print "%d costs for Sangerfield (580 / 66)" % counts_by_plant['580']
    except Exception, e:
        print "Error encountered in convert_ingredient_list: " + str(e)
        print traceback.format_exc()
        raw_input("Press enter to quit")
        sys.exit()

def convert_to_xf1(plants):
    try:
        plant_product_codes = {}
        output_filename = directory+os.sep+'Master%sIng.xf1' % formated_date
        f = open(output_filename,'wb')
        f.write('XF Version = 5\r\n')

        prefix_exclusions = {}
        with open('exclusions.txt', 'rb') as exclusionsfile:
            reader = csv.reader(exclusionsfile, delimiter=',', quotechar='"')
            for row in reader:
                if row[0]:
                    prefix_exclusions[row[0]] = True

        for plant in plants:
            reader = None
            valid_input_file = False
            while not valid_input_file:
                prompt = "SS ingredient cost file for %s (%s / %s)" % (plant['plant_name'], plant['plant_number'], plant['product_code_prefix'])
                input_filename = get_required_file(prompt)
                csvfile = open(input_filename, 'rb')
                reader = csv.reader(csvfile, delimiter='\t', quotechar='"')
                next(reader)
                next(reader)
                plant_name_line = next(reader)
                input_file_plant_name = plant_name_line[0]
                if input_file_plant_name.lower() != plant['plant_name'].lower():
                    print "\r\n%s is for %s, %s expected" % (input_filename, input_file_plant_name, plant['plant_name'])
                    prompt = "Try again for %s? [y/n]: " % plant['plant_name']
                    user_input = raw_input(prompt)
                    if user_input.lower() == 'n':
                        continue
                else:
                    valid_input_file = True

            next(reader)
            next(reader)
            next(reader)
            product_codes_sorted = []
            product_codes = {}
            for row in reader:
                if row:
                    if row[0] == 'Product' or row[0] == 'Code':
                        continue
                    f.write('IP0')
                    product_code = row[0].strip()
                    product_code = product_code.zfill(4)
                    if product_code not in prefix_exclusions:
                        product_code = plant['product_code_prefix'] + product_code
                    else:
                        print "Skipping prefix for %s" % product_code
                    product_codes[product_code] = True
                    product_codes_sorted.append(product_code)
                    f.write((product_code).rjust(12))
                    price_ton = float(row[2].strip())
                    price = "%.4f" % round(price_ton/20 ,4)
                    f.write(price.rjust(12))
                    f.write(product_code.rjust(128))
                    f.write('\r\n')

            product_codes_sorted.sort()
            plant_product_codes[plant['plant_name']] = {
                'product_codes': product_codes,
                'product_codes_sorted': product_codes_sorted,
            }

        if IS_WIN:
            ingredient_selection(plant_product_codes)
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

def ingredient_selection(plant_product_codes):
    try:
        for plant_name in plant_product_codes:
            print "You will have 5 seconds to select the code column of a row on the export screen."
            prompt = "Press enter when ready for ingredient selection for %s, or n to skip: " % plant_name
            user_input = raw_input(prompt)

            if user_input.lower() != 'n':
                time.sleep(5)

                last_data = ''
                data = copy_and_get_clipboard_data()
                shell.SendKeys("{DOWN}", 0)
                max_product_code = plant_product_codes[plant_name]['product_codes_sorted'][-1]

                found = 0
                while data < max_product_code and data != last_data:
                    last_data = data
                    data = copy_and_get_clipboard_data()
                    if data in plant_product_codes[plant_name]['product_codes']:
                        print "%s found" % data
                        shell.SendKeys("{LEFT}", 0)
                        shell.SendKeys(" ", 0)
                        shell.SendKeys("{RIGHT}", 0)
                        found = found + 1
                    shell.SendKeys("{DOWN}", 0)
                print "%d codes in file" % len(plant_product_codes[plant_name]['product_codes_sorted'])
                print "%d codes found and selected" % found
    except Exception, e:
        print "Error encountered in ingredient_selection: " + str(e)
        print traceback.format_exc()
        raw_input("Press enter to quit")
        sys.exit()

monthly = raw_input("Is this a monthly update? [y/n]: ")
if monthly.lower() in ['y','']:
    convert_ingredient_list()
    process_ss = raw_input("Process Solid Solutions? [y/n]: ")
else:
    process_ss = 'y'

if process_ss.lower() in ['y','']:
    convert_to_xf1([
        {'product_code_prefix': '61', 'plant_number': '560', 'plant_name': 'Augusta Mill'},
        {'product_code_prefix': '64', 'plant_number': '550', 'plant_name': 'Adams Center'},
        {'product_code_prefix': '66', 'plant_number': '580', 'plant_name': 'Sangerfield Mill'},
        {'product_code_prefix': '68', 'plant_number': '570', 'plant_name': 'Brandon Mill'}
    ])

raw_input("Process complete. Press enter to quit")
