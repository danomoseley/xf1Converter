#!/usr/bin/python
from flask import Flask, request, url_for, jsonify
app = Flask(__name__)

import sys, csv, time, os, pprint, traceback, collections, platform, json
import subprocess
from Tkinter import Tk
import Tkinter
from tkFileDialog import askopenfilename
import tkSimpleDialog
IS_WIN = platform.system().lower() == 'windows'

'''
print 'Checking for latest version'
p = subprocess.Popen(["git", "pull"], shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
out, err = p.communicate()
if not 'Already up-to-date.' in out:
    print 'Version updated, restarting'
    python = sys.executable
    os.execl(python, python, * sys.argv)
else:
    print 'Up to date'
'''

import os
from flask import Flask, request, redirect, url_for
from werkzeug import secure_filename

t = time.localtime()
formated_date = time.strftime('%m%d%y', t)
def get_export_directory():
    directory = formated_date
    if not os.path.exists(directory):
        os.makedirs(directory)
    return directory
 
UPLOAD_FOLDER = get_export_directory()
ALLOWED_EXTENSIONS = set(['txt','TXT','csv'])
 
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
 
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'GET':
        # we are expected to return a list of dicts with infos about the already available files:
        file_infos = []
        for file_name in list_files():
            file_url = url_for('download', file_name=file_name)
            file_size = get_file_size(file_name)
            file_infos.append(dict(name=file_name,
                                   size=file_size,
                                   url=file_url))
        return jsonify(files=file_infos)

    if request.method == 'POST':
        directory = 'static/uploads/'+formated_date
        if not os.path.exists(directory):
            os.makedirs(directory)

        return_files = []
        files = request.files
        for upload_file in files.getlist('files[]'):
            file_name = upload_file.filename
            upload_file.save(os.path.join(directory, file_name))
            file_url = directory+'/'+file_name
            pprint.pprint(upload_file)
            return_files.append({
                'url': file_url,
                'name': file_name
            })
        return json.dumps({'files':return_files})

if IS_WIN:
    import win32com.client
    import win32clipboard

    shell = win32com.client.Dispatch("WScript.Shell")

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
    input_filename = get_required_file_confirm("Agris csv cost file", filetypes=(("CSV", "*.csv"),
                                           ("All files", "*.*") ))
    read_product_cost_file(input_filename)

def read_product_cost_file(input_filename):
    try:
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

def convert_cost_list_files(cost_list_files, product_costs, base_file_path):
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

    directory = get_export_directory()
    cost_report_filename = base_file_path+os.sep+'misc_ingredient_report_'+formated_date+'.csv'

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

    return cost_report_filename 

def convert_ingredient_list():
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

    product_cost = read_product_cost()

    cost_list_files = []
    cost_list_files.append(get_required_file_confirm("Misc Ingredient list for NY"))
    cost_list_files.append(get_required_file_confirm("Misc Ingredient list for NE"))
    convert_cost_list_files(cost_list_files, product_costs)

    convert_ingredient_list_files(plant_files)

def convert_ingredient_list_files(plant_files, product_cost, base_file_path):
    try:
        output_filename = base_file_path+os.sep+'cost_output_'+formated_date+'.xf1'
        f = open(output_filename,'wb')
        exception_file = base_file_path+os.sep+'cost_exception_report_'+formated_date+'.txt'
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

        return exception_file, output_filename, counts_by_plant
    except Exception, e:
        print "Error encountered in convert_ingredient_list: " + str(e)
        print traceback.format_exc()
        raw_input("Press enter to quit")
        sys.exit()

def convert_to_xf1(input_filename, product_code_prefix, plant_number, plant_name, base_file_path):
    try:
        #prompt = "SS ingredient cost file for %s (%s / %s)" % (plant_name, plant_number, product_code_prefix)
        #input_filename = get_required_file(prompt)
        #input_filename = 'BRANDON TEST MO 09.02.13.TXT'

        output_filename = base_file_path+os.sep+'Master%sIng%s.xf1' % (formated_date, product_code_prefix)

        prefix_exclusions = {}
        with open('exclusions.txt', 'rb') as exclusionsfile:
            reader = csv.reader(exclusionsfile, delimiter=',', quotechar='"')
            for row in reader:
                if row[0]:
                    prefix_exclusions[row[0]] = True

        product_codes_sorted = []
        product_codes = {}
        with open(input_filename, 'rb') as csvfile:
            f = open(output_filename,'wb')
            reader = csv.reader(csvfile, delimiter='\t', quotechar='"')
            f.write('XF Version = 5\r\n')
            next(reader)
            next(reader)
            plant = next(reader)

            if plant and plant[0].lower() != plant_name.lower():
                print "\r\n%s is for %s, %s expected" % (input_filename, plant[0], plant_name)
                convert_to_xf1(product_code_prefix, plant_number, plant_name)
                return False

            next(reader)
            next(reader)
            next(reader)
            product_codes_loaded = None
            product_codes_found = None
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
        if IS_WIN:
            product_codes_loaded, product_codes_found = ingredient_selection(product_codes, product_codes_sorted)
        return output_filename, product_codes_loaded, product_codes_found
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
                shell.SendKeys("{LEFT}", 0)
                shell.SendKeys(" ", 0)
                shell.SendKeys("{RIGHT}", 0)
                found = found + 1
            shell.SendKeys("{DOWN}", 0)

            return len(product_codes_sorted), found
    except Exception, e:
        print "Error encountered in ingredient_selection: " + str(e)
        print traceback.format_exc()
        raw_input("Press enter to quit")
        sys.exit()


@app.route('/convert_ingredients', methods=['GET'])
def convert_ingredients():
    input_filename = 'static/uploads/090213/Misc Product Cost 8.30.13.csv'
    product_costs = read_product_cost_file(input_filename)

    plant_files = [
        ['static/uploads/090213/SEPT EXPORT 550.TXT', '550'],
        ['static/uploads/090213/SEPT EXPORT 560.TXT', '560'],
        ['static/uploads/090213/SEPT EXPORT 570.TXT', '570'],
        ['static/uploads/090213/SEPT EXPORT 580.TXT', '580']
    ]

    cost_list_files = ['static/uploads/090213/NE MISC INGRED CODES.TXT', 'static/uploads/090213/NY NISC ING CODES.TXT']
    cost_report_path = convert_cost_list_files(cost_list_files, product_costs, 'static/uploads/090213')

    exception_file, output_filename, counts_by_plant = convert_ingredient_list_files(plant_files, product_costs, 'static/uploads/090213')

    return jsonify({
        'cost_report': 'http://localhost:5001/'+cost_report_path,
        'exception_report': 'http://localhost:5001/'+exception_file,
        'ingredient_export': 'http://localhost:5001/'+output_filename,
        'count_by_plant': counts_by_plant
    })

@app.route('/convert_solid_solutions', methods=['GET'])
def convert_solid_solutions():
    ss_convert_filepath, product_codes_loaded, product_codes_found = convert_to_xf1('static/uploads/090213/BRANDON TEST MO 09.02.13.TXT', '68', '570', 'Brandon Mill', 'static/uploads/090213')
    return jsonify({
        '570': {
            'file': 'http://localhost:5001/'+ss_convert_filepath,
            'loaded': product_codes_loaded,
            'found': product_codes_found
        }
    })

'''
monthly = raw_input("Is this a monthly update? [y/n]: ")
if monthly.lower() in ['y','']:
    convert_ingredient_list()
    process_ss = raw_input("Process Solid Solutions? [y/n]: ")
else:
    process_ss = 'y'

if process_ss.lower() in ['y','']:
    convert_to_xf1('61', '560', 'Augusta Mill')
    convert_to_xf1('64', '550', 'Adams Center')
    convert_to_xf1('66', '580', 'Sangerfield Mill')
    convert_to_xf1('68', '570', 'Brandon Mill')

raw_input("Process complete. Press enter to quit")
'''

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5001, debug=True)
