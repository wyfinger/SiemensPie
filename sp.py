#!/usr/bin/python
# -*- coding: utf-8 -*-
import argparse
import codecs
import json
import os
import re
import sys
import binascii

import time
from lxml import etree as lxml

from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet

import winreg

__author__ = 'Wyfinger'
__version__ = '2020-02-20'

xriofile = ""
xmlfile = ""
config_tree = {}
xrio_tree = None
xml_tree = None
primary = True
ktt = 1
ktn = 1
book = Workbook
sheet = Worksheet
cell_formats = {}

cur_row = 0
stash = {}
last_h1_title = ""


# TODO: уставка 7137 в 7SJ не выводится в Excel, хотя в конфиге, если посмотреть в Digsi, есть.

def print_small_help():
    # print small help tip to console, for use in error in parameters

    print("Use: sp.exe [-c] [xml or xrio file] [xml or xrio file]")
    print("  -c  - path to config.json file")
    print("  set one (xml or xrio) file if they have the same name")
    print("")

    return


def read_config(config_path):
    # read config file (Json)

    try:
        # print('Reading config_tree from: ' + config_path)
        with codecs.open(config_path, 'r', 'utf-8') as param_file:
            rez = json.load(param_file)
    except:
        print("Error at read config.json file.\n")
        print_small_help()
        time.sleep(5)
        sys.exit()

    return rez


def print_h1(text):
    # insert chapter header

    global cur_row, last_h1_title

    # check text to titles_correct
    text = config_tree['titles_correct'].get(text, text)

    sheet.merge_range(cur_row, 0, cur_row, 7, text, cell_formats[9])
    cur_row = cur_row + 1

    # save last h1 header for titles_correct inn PrintH2
    last_h1_title = text

    return


def print_h2(text):
    # insert chapter sub-header

    global cur_row

    if text != "" and text != last_h1_title:
        # check text to titles_correct
        text = config_tree['titles_correct'].get(last_h1_title + "|" + text, text)

        sheet.merge_range(cur_row, 0, cur_row, 7, text, cell_formats[10])
        cur_row = cur_row + 1

    return


def print_h3(title_prefix='', groups=False):
    # insert parameter groups header

    global cur_row

    if title_prefix != '':
        if groups:  # 3 rows
            sheet.merge_range(cur_row, 0, cur_row + 2, 0, "Адрес", cell_formats[11])  # column 1
            sheet.merge_range(cur_row, 1, cur_row + 2, 1, "Параметр", cell_formats[11])  # column 2
            sheet.write(cur_row, 2, "Значение/диапазон/шаг", cell_formats[12])  # column 3
            sheet.merge_range(cur_row + 1, 2, cur_row + 2, 2, "(вторичные величины)", cell_formats[13])
            sheet.merge_range(cur_row, 3, cur_row, 6, "Задаваемый параметр", cell_formats[12])  # column 4
            sheet.merge_range(cur_row + 1, 3, cur_row + 1, 6, title_prefix, cell_formats[13])
            sheet.write(cur_row + 2, 3, "Группа A", cell_formats[11])
            sheet.write(cur_row + 2, 4, "Группа B", cell_formats[11])
            sheet.write(cur_row + 2, 5, "Группа C", cell_formats[11])
            sheet.write(cur_row + 2, 6, "Группа D", cell_formats[11])
            sheet.merge_range(cur_row, 7, cur_row + 2, 7, "Комментарий", cell_formats[11])  # column 5
            cur_row = cur_row + 3
        else:  # 2 rows
            sheet.merge_range(cur_row, 0, cur_row + 1, 0, "Адрес", cell_formats[11])  # column 1
            sheet.merge_range(cur_row, 1, cur_row + 1, 1, "Параметр", cell_formats[11])  # column 2
            sheet.write(cur_row, 2, "Значение/диапазон/шаг", cell_formats[12])  # column 3
            sheet.write(cur_row + 1, 2, "(вторичные величины)", cell_formats[13])
            sheet.merge_range(cur_row, 3, cur_row, 6, "Задаваемый параметр", cell_formats[12])  # column 4
            sheet.merge_range(cur_row + 1, 3, cur_row + 1, 6, title_prefix, cell_formats[13])
            sheet.merge_range(cur_row, 7, cur_row + 1, 7, "Комментарий", cell_formats[11])  # column 5
            cur_row = cur_row + 2
    else:
        if groups:  # 2 rows
            sheet.merge_range(cur_row, 0, cur_row + 1, 0, "Адрес", cell_formats[11])  # column 1
            sheet.merge_range(cur_row, 1, cur_row + 1, 1, "Параметр", cell_formats[11])  # column 2
            sheet.merge_range(cur_row, 2, cur_row + 1, 2, "Значение/диапазон/шаг", cell_formats[11])  # column 3
            sheet.merge_range(cur_row, 3, cur_row, 6, "Задаваемый параметр", cell_formats[11])  # column 4
            sheet.write(cur_row + 1, 3, "Группа A", cell_formats[11])
            sheet.write(cur_row + 1, 4, "Группа B", cell_formats[11])
            sheet.write(cur_row + 1, 5, "Группа C", cell_formats[11])
            sheet.write(cur_row + 1, 6, "Группа D", cell_formats[11])
            sheet.merge_range(cur_row, 7, cur_row + 1, 7, "Комментарий", cell_formats[11])  # column 5
            cur_row = cur_row + 2
        else:  # 1 rows
            sheet.write(cur_row, 0, "Адрес", cell_formats[11])  # column 1
            sheet.write(cur_row, 1, "Параметр", cell_formats[11])  # column 2
            sheet.write(cur_row, 2, "Значение/диапазон/шаг", cell_formats[11])  # column 3
            sheet.merge_range(cur_row, 3, cur_row, 6, "Задаваемый параметр", cell_formats[11])  # column 4
            sheet.write(cur_row, 7, "Комментарий", cell_formats[11])  # column 5
            cur_row = cur_row + 1

    return


def extract_parameter_name(address):
    # get parameter info from XRio file

    parameter_name = xrio_tree.xpath("//ForeignId[text()='" + address + "']/parent::*/Name/text()")
    if (parameter_name is not None) and (len(parameter_name) > 0):
        parameter_name = str(parameter_name[0])
    else:
        parameter_name = ""

    return parameter_name


def extract_parameter_precision(address):
    # get parameter precision from XRio file

    parameter_precision = xrio_tree.xpath("//ForeignId[text()='" + address + "']/parent::*/Unit")
    if (parameter_precision is not None) and (len(parameter_precision) > 0):
        return int(parameter_precision[0].attrib['DecimalPlaces'])
    else:
        return 0


def convert_to_primary(address, value, dimension, secondary_precision):
    # convert electrical value to primary

    # do not convert special addresses

    # this is a number
    rez = value
    if re.search(r"\d+(\.|)\d*", value, re.MULTILINE):
        value = float(value)
        if dimension == "А":
            rez = "%g" % round(value * ktt, secondary_precision - 1) + " " + dimension
        elif dimension == "В":
            rez = "%g" % (value * ktn / 1000) + " кВ"
        elif dimension == "Ом":  # 2018-03-23: в 7SA в первичных 3 знака после запятой, в 7SD - два, везде делаем 3
            rez = "%g" % round(value * ktn / ktt, secondary_precision) + " " + dimension
        elif dimension == "Ом / км":
            rez = "%g" % round(value * ktn / ktt, secondary_precision - 1) + " " + dimension
        elif dimension == "ВА":
            rez = "%g" % round(value * ktn * ktt / 1000000, secondary_precision + 1) + " МВА"
        elif dimension == "мкФ/км":
            rez = "%g" % round(value * ktt / ktn, secondary_precision + 1) + " " + dimension
        else:
            rez = "%g" % float(rez) + " " + dimension

    return str(rez)


def extract_parameter_values(parameter):
    # extract parameter values in all groups of parameters

    parameter_addr = parameter.attrib['DAdr']
    parameter_type = parameter.attrib['Type']

    parameter_value = parameter.find(r"Value")
    parameter_value_a = parameter.find(r"Value[@SettingGroup='A']")
    parameter_value_b = parameter.find(r"Value[@SettingGroup='B']")
    parameter_value_c = parameter.find(r"Value[@SettingGroup='C']")
    parameter_value_d = parameter.find(r"Value[@SettingGroup='D']")

    if parameter_value_a is None:
        parameter_value_a = parameter_value.text
        parameter_value_b = parameter_value.text
        parameter_value_c = parameter_value.text
        parameter_value_d = parameter_value.text
    else:
        parameter_value_a = parameter_value_a.text
        parameter_value_b = parameter_value_b.text
        parameter_value_c = parameter_value_c.text
        parameter_value_d = parameter_value_d.text

    if parameter_type == "Txt":
        parameter_value_a = parameter.find(r"Comment[@Number='" + parameter_value_a + "']").attrib['Name']
        parameter_value_b = parameter.find(r"Comment[@Number='" + parameter_value_b + "']").attrib['Name']
        parameter_value_c = parameter.find(r"Comment[@Number='" + parameter_value_c + "']").attrib['Name']
        parameter_value_d = parameter.find(r"Comment[@Number='" + parameter_value_d + "']").attrib['Name']
    else:
        dimension = parameter.find('Comment[@Dimension]')
        if dimension is not None:
            dimension = dimension.attrib.get('Dimension')
        else:
            dimension = ''

        # convert to primary if needed
        if (primary is False) | (parameter_addr in config_tree['non_electrical']):
            # if value is "oo" - do not display dimension
            # call ConvertToPrimary for calc 'group_has_elec_values' variable
            convert_to_primary(parameter_addr, parameter_value_a, dimension, extract_parameter_precision(parameter_addr))
            parameter_value_a = parameter_value_a if parameter_value_a == "oo" else parameter_value_a + " " + dimension
            parameter_value_b = parameter_value_b if parameter_value_b == "oo" else parameter_value_b + " " + dimension
            parameter_value_c = parameter_value_c if parameter_value_c == "oo" else parameter_value_c + " " + dimension
            parameter_value_d = parameter_value_d if parameter_value_d == "oo" else parameter_value_d + " " + dimension
        else:
            secondary_precision = extract_parameter_precision(parameter_addr)
            parameter_value_a = convert_to_primary(parameter_addr, parameter_value_a, dimension, secondary_precision)
            parameter_value_b = convert_to_primary(parameter_addr, parameter_value_b, dimension, secondary_precision)
            parameter_value_c = convert_to_primary(parameter_addr, parameter_value_c, dimension, secondary_precision)
            parameter_value_d = convert_to_primary(parameter_addr, parameter_value_d, dimension, secondary_precision)

    return [parameter_value_a.strip(), parameter_value_b.strip(), parameter_value_c.strip(), parameter_value_d.strip()]


def extract_parameter_range(parameter):
    # extract parameter range

    parameter_type = parameter.attrib['Type']
    range_text = ''
    precision = 0
    if parameter_type == "Txt":
        comments = parameter.findall('Comment[@Name]')
        for comment in comments:
            if len(range_text) != 0:
                range_text = range_text + "\r\n"
            if comment.attrib['Name'] != '':
                range_text = range_text + comment.attrib['Name']
    elif parameter_type == "Dec":
        comment = parameter.find('Comment')
        dimension = comment.attrib.get('Dimension')
        if dimension is None:
            dimension = ''
        min_value = comment.attrib['MinValue']
        max_value = comment.attrib['MaxValue']
        precision = len(min_value) - min_value.rfind(".")
        if precision == len(min_value) + 1:
            precision = 0
        additional_valid_values = comment.attrib.get('AdditionalValidValues')
        if additional_valid_values is None:
            range_text = min_value + " … " + max_value + " " + dimension
        else:
            range_text = min_value + " … " + max_value + " " + dimension + "; " + additional_valid_values

    return [range_text, precision]


def print_parameter_data(parameter_data, highlight=False):
    # paste parameter info to output excel sheet

    global cur_row

    # write data from config_tree then correct it by "params_correct" config_tree section
    if len(parameter_data['Address']) > 6:
        sheet.write(cur_row, 0, parameter_data['Address'], cell_formats[15])
    elif highlight:
        sheet.write(cur_row, 0, parameter_data['Address'], cell_formats[14])
    else:
        sheet.write(cur_row, 0, parameter_data['Address'], cell_formats[1])
    sheet.write(cur_row, 1, parameter_data['Name'], cell_formats[2])
    sheet.write(cur_row, 2, parameter_data['Range'], cell_formats[3])

    parameter_values = parameter_data['Values']

    # if values are equal merge cells
    if parameter_values[0] == parameter_values[1] == parameter_values[2] == parameter_values[3]:
        sheet.merge_range(cur_row, 3, cur_row, 6, parameter_values[0], cell_formats[4])
    elif parameter_values[0] == parameter_values[1] == parameter_values[2]:
        sheet.merge_range(cur_row, 3, cur_row, 5, parameter_values[0], cell_formats[4])
        sheet.write(cur_row, 6, parameter_values[3], cell_formats[7])
    elif parameter_values[0] == parameter_values[1]:
        sheet.merge_range(cur_row, 3, cur_row, 4, parameter_values[0], cell_formats[4])
        if parameter_values[2] == parameter_values[3]:
            sheet.merge_range(cur_row, 5, cur_row, 6, parameter_values[2], cell_formats[6])
        else:
            sheet.write(cur_row, 5, parameter_values[2], cell_formats[6])
            sheet.write(cur_row, 6, parameter_values[3], cell_formats[7])
    elif parameter_values[1] == parameter_values[2] == parameter_values[3]:
        sheet.write(cur_row, 3, parameter_values[0], cell_formats[4])
        sheet.merge_range(cur_row, 4, cur_row, 6, parameter_values[1], cell_formats[5])
    elif parameter_values[2] == parameter_values[3]:
        sheet.write(cur_row, 3, parameter_values[0], cell_formats[4])
        sheet.write(cur_row, 4, parameter_values[1], cell_formats[5])
        sheet.merge_range(cur_row, 5, cur_row, 6, parameter_values[2], cell_formats[6])
    elif parameter_values[1] == parameter_values[2]:
        sheet.write(cur_row, 3, parameter_values[0], cell_formats[4])
        sheet.merge_range(cur_row, 4, cur_row, 5, parameter_values[1], cell_formats[5])
        sheet.write(cur_row, 6, parameter_values[3], cell_formats[7])
    else:
        sheet.write(cur_row, 3, parameter_values[0], cell_formats[4])
        sheet.write(cur_row, 4, parameter_values[1], cell_formats[5])
        sheet.write(cur_row, 5, parameter_values[2], cell_formats[6])
        sheet.write(cur_row, 6, parameter_values[3], cell_formats[7])

    sheet.write(cur_row, 7, parameter_data['Description'], cell_formats[8])

    # and correct fields from 'params_correct' section of config
    addr = parameter_data['Address']
    need_correct = config_tree["params_correct"].get(addr, None)
    if need_correct is not None:
        if not isinstance(need_correct[0], list):
            need_correct = [need_correct]
        for patch in need_correct:
            col_no = patch[0]
            col_val = patch[1]
            sheet.write(cur_row, int(col_no), col_val, cell_formats[col_no + 1] if col_no in range(0, 7) else cell_formats[0])

    cur_row = cur_row + 1

    # insert formula with comments
    # !!!
    # sheet.Cells(currow, 9).FormulaR1C1 = '=IFERROR(IF(TRIM(RC[-8])<>"",INDEX(\'\\\\Prim-fs-serv\\rdu\СРЗА\\Уставки\\РАСЧЕТЫ УСТАВОК\\
    # [!!!Siemens, общие комментарии.xlsx]7SD\'!C1:C2,MATCH(RC[-8],\'\\\\Prim-fs-serv\\rdu\\СРЗА\\Уставки\\РАСЧЕТЫ УСТАВОК\\[!!!Siemens,
    # общие комментарии.xlsx]7SD\'!C1,0),2),""),"")'
    # sheet.Cells(currow, 9).VerticalAlignment = -4108 # xlCenter

    return


def register_xrio_ext():
    # register .xrio extention for siemens py

    try:
        if getattr(sys, 'frozen', False):
            exe_path = sys.executable
            ico_path = os.path.join(os.path.dirname(sys.executable), 'doc.ico')
        else:
            return

        exe_path = '"' + exe_path + '" "%1"'
        ico_path = ico_path + ",0"
        key_xrio = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r".xrio")
        winreg.SetValue(key_xrio, None, winreg.REG_SZ, r"SiemensPie.XRio")
        winreg.SetValue(key_xrio, r"Content Type", winreg.REG_SZ, r"text/html")
        winreg.CloseKey(key_xrio)
        key_xrio = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"SiemensPie.XRio")
        key_di = winreg.CreateKey(key_xrio, r"DefaultIcon")
        winreg.SetValue(key_di, None, winreg.REG_SZ, ico_path)
        winreg.CloseKey(key_di)
        key_sh = winreg.CreateKey(key_xrio, r"shell")
        key_op = winreg.CreateKey(key_sh, r"open")
        key_cmd = winreg.CreateKey(key_op, r"command")
        winreg.SetValue(key_cmd, None, winreg.REG_SZ, exe_path)
        winreg.CloseKey(key_cmd)
        winreg.CloseKey(key_op)
        winreg.CloseKey(key_sh)

    except:
        print("Error at register shell extension for .xrio files, work continues.. \n")

    return


def process_command_line():
    # command line parameter analyses
    global xriofile, xmlfile, config_tree

    if getattr(sys, 'frozen', False):
        config_path = os.path.join(os.path.dirname(sys.executable), "config.json")
    else:
        config_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "config.json")

    parser = argparse.ArgumentParser()
    parser.add_argument('-c', '--config', nargs='?', default=config_path, help='Json file with parameters')
    parser.add_argument('file1', nargs='?')
    parser.add_argument('file2', nargs='?')
    namespace = parser.parse_args()

    # find .xml and .xrio files
    if namespace.file1 is not None:
        file1name, file1extension = os.path.splitext(namespace.file1)
        if file1extension.lower() == '.xrio':
            xriofile = namespace.file1
        if (file1extension.lower() == '.xml'):
            xmlfile = namespace.file1
    if namespace.file2 is not None:
        file2name, file2extension = os.path.splitext(namespace.file2)
        if file2extension.lower() == '.xrio':
            xriofile = namespace.file2
        if file2extension.lower() == '.xml':
            xmlfile = namespace.file2

    if (xriofile == "") & (xmlfile != ""):
        xriofile = os.path.splitext(xmlfile)[0] + '.xrio'
        if not os.path.isfile(xriofile):
            xriofile = ""
    if (xriofile != "") & (xmlfile == ""):
        xmlfile = os.path.splitext(xriofile)[0] + '.xml'
        if not os.path.isfile(xmlfile):
            xmlfile = ""

    if namespace.config is not None:
        config_path = namespace.config
    config_tree = read_config(config_path)

    if (xmlfile == "") | (xriofile == ""):
        print("Error. XML or XRio file is not exists.")
        print_small_help()
        time.sleep(5)
        sys.exit()

    print("XML: " + xmlfile)
    print("XRio: " + xriofile)

    return


def create_output_file():
    # create output excel file
    global book, sheet

    try:
        xlsx_path, ext = os.path.splitext(xmlfile)
        book = Workbook(xlsx_path + '.xlsx')
        sheet = book.add_worksheet()
    except:
        print("Error at create Excel file to output.\n")
        time.sleep(5)
        sys.exit()

    return


def page_setup():
    # insert header and page stylization

    # page margins, headers and footers
    sheet.set_margins(0.4, 0.4, 0.9, 0.8)
    sheet.set_header("", {'margin': 0.12})
    sheet.set_footer("&amp;R&amp;F\n&amp;P", {'margin': 0.12})
    sheet.set_zoom(90)
    sheet.set_landscape()
    sheet.set_paper(9)
    sheet.fit_to_pages(1, 0)
    sheet.set_footer("&R&F\r&R&P", {'margin': 0.25})

    # text formats
    global cell_formats

    # cell formats array
    cell_formats = {
        0: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 0}),  # default
        1: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),  # address
        2: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),  # name
        3: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),  # range
        4: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),  # value, g.A
        5: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),  # value, g.B
        6: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),  # value, g.C
        7: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),  # value, g.D
        8: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),  # desc
        9: book.add_format(
            {'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'font_size': 13, 'bg_color': '#EEECE1', 'border': 1}),
        # h1
        10: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#EEECE1', 'border': 1}),
        # h2
        11: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#EEECE1', 'border': 1}),
        # h3
        12: book.add_format(
            {'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#EEECE1', 'top': 1, 'left': 1,
             'right': 1, 'bottom': 0}),  # h3 without bottom border
        13: book.add_format(
            {'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#EEECE1', 'top': 0, 'left': 1,
             'right': 1, 'bottom': 1}),  # h3 without top border
        14: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': 'yellow', 'border': 1}),
        # highlight address
        15: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'color': 'white', 'font_size': '3', 'border': 1})
        # hide address
    }

    # set column width and formats
    sheet.set_column(0, 0, 6)  # address
    sheet.set_column(1, 1, 23)  # name
    sheet.set_column(2, 2, 30)  # range
    sheet.set_column(3, 6, 8.3)  # values in groups
    sheet.set_column(7, 7, 41.6)  # description

    return


def process_all():
    # start of data process

    global xrio_tree, xml_tree
    global config_tree, primary, ktt, ktn, cur_row

    # load xml files
    try:
        xrio_tree = lxml.parse(xriofile)
        xml_tree = lxml.parse(xmlfile)
    except:
        print("Error at read XML and XRio files.\n")
        print_small_help()
        time.sleep(5)
        sys.exit()

    # select config_tree section by MLFB code
    MLFBDIGSI = xml_tree.xpath('.//General/GeneralData[@Name="MLFBDIGSI"]/@ID')[0]
    for k in config_tree.keys():
        if MLFBDIGSI[0:len(k)] == k:
            config_tree = config_tree[k]
            break

    # 7SD52 may not include Set of Step Protections and may not contain voltage_primary parameters
    ktn = 1
    ktt = 1
    primary = config_tree['convert_to_primary'].lower() == 'true'
    if primary:
        try:
            v_primary = xml_tree.xpath(config_tree['voltage_primary'])
            v_primary = 1 if len(v_primary) == 0 else float(re.sub(r"[^\d+.]", "", v_primary[0], 0, 0)) * 1000  # voltage in kilovolts
            v_second = xml_tree.xpath(config_tree['voltage_second'])
            v_second = 1 if len(v_second) == 0 else float(re.sub(r"[^\d+.]", "", v_second[0], 0, 0))
            ktn = v_primary / v_second
            c_primary = xml_tree.xpath(config_tree['current_primary'])[0]
            c_primary = int(re.sub(r"[^\d+\.]", "", c_primary, 0, 0))
            c_second = xml_tree.xpath(config_tree['current_second'])[0]
            c_second = int(re.sub(r"[^\d+\.]", "", c_second, 0, 0))
            ktt = c_primary / c_second
        except:
            print("Can't find primary/secondary values of current and voltage.\n")
            time.sleep(5)
            sys.exit()

    # first cycle for each parameters, check 'params_to_rearrange'.
    # this section allows you to move some settings by placing them after the specified address.
    # it also allows you to add a new setpoint that does not exist in .xrio
    extract_parameters_to_rearrange()

    # paste overview info about terminal
    # MLFB code
    sheet.merge_range(cur_row, 0, cur_row, 1, "MLFB Код", cell_formats[0])
    sheet.merge_range(cur_row, 2, cur_row, 7, MLFBDIGSI, cell_formats[0])
    cur_row = cur_row + 1
    # Version
    sheet.merge_range(cur_row, 0, cur_row, 1, "Версия ПО терминала", cell_formats[0])
    # sheet.merge_range(cur_row, 2, cur_row, 7, xml_tree.xpath(
    #   './/General/GeneralData[@Name="Version"]/@ID')[0], frm_name)   # версия ПО из XML файла
    sheet.merge_range(cur_row, 2, cur_row, 7, xrio_tree.xpath(
        '//XRio/CUSTOM/Block/Block[@Id="GENERALINFO"]/Block/Parameter[@Id="SERIAL_NUMBER"]/Value/text()')[0],
                      cell_formats[0])  # версия ПО из XRio файла
    cur_row = cur_row + 1
    # Topology
    sheet.merge_range(cur_row, 0, cur_row, 1, "Топология", cell_formats[0])
    sheet.merge_range(cur_row, 2, cur_row, 7, xml_tree.xpath('.//General/GeneralData[@Name="Topology"]/@ID')[0], cell_formats[0])
    cur_row = cur_row + 1
    # Self version (crc32)
    sheet.merge_range(cur_row, 0, cur_row, 1, "Версия SiemensPie", cell_formats[0])
    sheet.merge_range(cur_row, 2, cur_row, 7, __version__, cell_formats[0])
    cur_row = cur_row + 1

    # main work
    function_groups = xml_tree.findall('Settings/FunctionGroup')
    for function_group in function_groups:
        process_function_group(function_group)

    return


def extract_parameters_to_rearrange():
    # process all XML file and extract params for rearrange to stash list

    global stash

    all_params = xml_tree.findall('Settings//Parameter')
    params_addrs = {}

    for p in all_params:
        params_addrs[p.attrib['DAdr']] = True

        parameter_data = {
            'Address': p.attrib['DAdr'],
            'Name': extract_parameter_name(p.attrib['DAdr']),
            'Range': extract_parameter_range(p)[0],
            'Values': extract_parameter_values(p),
            'Description': p.attrib['Name'],
        }

        # pop_after may be "auto" or parameter address
        pop_after = config_tree['params_to_rearrange'].get(parameter_data['Address'], 0)
        if pop_after != 0:
            stash[parameter_data['Address']] = {
                'PopAfter': pop_after,
                'ParameterData': parameter_data
            }

    # TODO: Add parameters from the 'params_correct' section with column 0, this is for
    #  special addresses that are in Digsi, but which are not in the .xml and .xrio files,
    #  for example 7137 in 7SJ.

    # check stash params for PopAfter exists in config, if in confid have not param with
    # address from 'PopAfter' - delete this param from stash
    for s in stash.copy():
        if stash[s]['PopAfter'] not in params_addrs:
            stash.pop(s)

    return


def process_function_group(function_group):
    # process function group

    function_group_name = function_group.attrib['Name']
    print_h1(function_group_name)
    setting_pages = function_group.findall("SettingPage")
    for setting_page in setting_pages:
        process_setting_page(setting_page)

    return


def process_setting_page(setting_page):
    # process settings page

    global cur_row

    setting_page_name = setting_page.attrib['Name']
    print_h2(setting_page_name)

    parameters = setting_page.findall("Parameter")

    page_have_elec_param = False
    page_have_groups = False

    for parameter in parameters:
        if parameter.attrib['Type'] != "Txt":
            dimension = parameter.find('Comment[@Dimension]')
            if dimension is not None:
                dimension = dimension.attrib.get('Dimension')
            else:
                dimension = ''
            if (dimension in ["А", "В", "Ом", "Ом / км", "ВА", "мкФ/км"]) & (not parameter.attrib['DAdr'] in config_tree['non_electrical']):
                page_have_elec_param = True
            if parameter.find(r"Value[@SettingGroup='A']") is not None:
                page_have_groups = True

    if page_have_elec_param:
        if primary:
            print_h3('(первичные величины)', page_have_groups)
        else:
            print_h3('(вторичные величины)', page_have_groups)
    else:
        print_h3('', page_have_groups)

    for parameter in parameters:
        process_parameter(parameter)

    return


def process_parameter(parameter):
    # precess one parameter / address

    address = parameter.attrib['DAdr']
    parameter_data = {
        'Address': address,
        'Name': extract_parameter_name(address),
        'Range': extract_parameter_range(parameter)[0],
        'Values': extract_parameter_values(parameter),
        'Description': parameter.attrib['Name'],
    }

    print(address)

    insert_parameter(parameter_data)

    return


def insert_parameter(parameter_data, rearrange=False):
    # precess one parameter / address

    global stash

    # print current address if this address is absent in stash
    if (parameter_data['Address'] not in stash) | rearrange:
        print_parameter_data(parameter_data, rearrange)

    # if after current param we need to insert stashed param - do it!
    for s in stash.copy():
        if stash[s]['PopAfter'] == parameter_data['Address']:
            insert_parameter(stash[s]['ParameterData'], True)
            stash.pop(s)

    return


#
# Main work stats from print hello message
#
print("SiemensPie - tool for convert Siemens Siprotec relay protection config")
print("to readable and editable format (Excel .xlsx file)")
print("(C) " + __version__ + " " + __author__ + ", https://github.com/wyfinger/SiemensPie")
print("")

register_xrio_ext()
process_command_line()
create_output_file()
page_setup()
process_all()
book.close()

time.sleep(5)
