from __future__ import unicode_literals, print_function, division
import os
import sys
import xlrd
import json
import configparser
from option import option
import locale
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')


def get_cell_values(sheet, row_col):
    if "__" in row_col:
        return get_cell_value(sheet, row_col.split('__')[0]), get_cell_value(sheet, row_col.split('__')[1])
    else:
        return get_cell_value(sheet, row_col)


def get_cell_value(sheet, row_col):
    row = row_col.split(', ')[0]
    col = row_col.split(', ')[1]
    return get_cell_value_with_separate_coordinates(sheet, row, col)


def get_cell_value_with_separate_coordinates(sheet, row, col):
    try:
        val = sheet.cell(int(row), int(col))
        str_val = val.value
        return str_val
    except:
        return ""


def convert_xlsx_to_json(folder_path, output_folder_path, file_name, config):
    try:
        wb = xlrd.open_workbook(folder_path + file_name)
        sheet = wb.sheet_by_index(0)
        obj_option = option
        for sec in config.sections():
            for key in config[sec]:
                if "__" in config[sec][key]:
                    values = get_cell_values(sheet, config[sec][key])
                    obj_option[sec][key.split('__')[0]][key.split('__')[1]], \
                    obj_option[sec][key.split('__')[0]][key.split('__')[2]] = values
                else:
                    obj_option[sec][key] = get_cell_value(sheet, config[sec][key])
        with open(output_folder_path + file_name.split("_")[0] + '_Option.json', 'w') as outfile:
            json.dump(obj_option, outfile)
    except:
        print(file_name + ": ERROR_3")


if len(sys.argv) < 3:
    print("USAGE: python3 convertor.py FOLDER_PATH OUTPUT_FOLDER_PATH FILE_NAME")
    sys.exit()

config = configparser.ConfigParser()
config.read('xlsx_dictionary.ini')

if len(sys.argv) == 3:
    for filename in os.listdir(sys.argv[1]):
        if filename.endswith(".xlsx"):
            convert_xlsx_to_json(sys.argv[1], sys.argv[2], filename, config)
elif len(sys.argv) == 4:
    convert_xlsx_to_json(sys.argv[1], sys.argv[2], sys.argv[3], config)

    


# python converter.py /home/ashkhen/Documents/xls/ /home/ashkhen/Documents/jsn/