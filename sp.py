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
from xlsxwriter.format import Format

import winreg

__author__ = 'wyfinger'
__version__ = '2020-01-28'

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
group_has_elec_values = False
group_has_group_values = False
stash = []
inserted_stash = []
last_printed_address = '0'
last_h1_title = ""

# TODO: уставка 7137 в 7SJ не выводится в Excel, хотя в конфиге, если посмотреть в Digsi, есть.


'''
 print small help tip to console, for use in error in parameters
'''
def print_small_help():

    print("Use: sp.exe [-c] [xml or xrio file] [xml or xrio file]")
    print("  -c  - path to config.json file")
    print("  set one (xml or xrio) file if they have the same name")
    print("")

    return


'''
 read config file (Json)
'''
def read_config(config_path):

    try:
        # print('Reading config_tree from: ' + config_path)
        with codecs.open(config_path, 'r', 'utf-8') as param_file:
            return json.load(param_file)
    except:
        print("Error at read config.json file.\n")
        print_small_help()
        time.sleep(5)
        sys.exit()

    return


'''
 command line parameter analyses
'''
def process_command_line():

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
    """create output excel file
    """
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


'''
 process stash scrap
'''
def print_consumed_params():

    global stash, cur_row
    global group_has_elec_values, group_has_group_values

    # if stash is clear (all parameters are inserted) we need return
    if any(p['Inserted']== True for p in stash):  return

    # else let's print consumed parameters

    cur_row = cur_row + 3

    print_h1("Потерянные уставки")

    group_has_elec_values = False
    group_has_group_values = False
    header_row = cur_row
    cur_row = cur_row + 2

    for st in stash:
        if st['ParameterData']['Address'] not in inserted_stash:
            print_parameter_data(st['ParameterData'])
            sheet.write(cur_row - 1, 14, 'Pop After Address: ' + st['PopAfter'])

    #print_group_header(header_row)

    return


'''
 start of data process
'''
def process_all():

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

    extract_parameters_to_rearrange();

    # paste overview info about terminal
    # MLFB code
    sheet.merge_range(cur_row, 0, cur_row, 1, "MLFB Код", cell_formats[0])
    sheet.merge_range(cur_row, 2, cur_row, 7, MLFBDIGSI, cell_formats[0])
    cur_row = cur_row + 1
    # Version
    sheet.merge_range(cur_row, 0, cur_row, 1, "Версия ПО терминала", cell_formats[0])
    # sheet.merge_range(cur_row, 2, cur_row, 7, xml_tree.xpath('.//General/GeneralData[@Name="Version"]/@ID')[0], frm_name)   # версия ПО из XML файла
    sheet.merge_range(cur_row, 2, cur_row, 7, xrio_tree.xpath(
        '//XRio/CUSTOM/Block/Block[@Id="GENERALINFO"]/Block/Parameter[@Id="SERIAL_NUMBER"]/Value/text()')[0], cell_formats[0])  # версия ПО из XRio файла
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
    FunctionGroups = xml_tree.findall('Settings/FunctionGroup')
    for FunctionGroup in FunctionGroups:
        process_function_group(FunctionGroup)

    # iss9: if ref address to move parameters is not exists we can still this parameter.
    #  But we can drop this parameters to end of exported list.
    print_consumed_params()

    return


def page_setup():
    """ insert header and page stylization
    """

    # page margins, headers and footers
    sheet.set_margins(0.4, 0.4, 0.9, 0.8)
    sheet.set_header("", {'margin': 0.12})
    sheet.set_footer("&amp;R&amp;F\n&amp;P", {'margin': 0.12})
    sheet.set_zoom(90)
    sheet.set_landscape()
    sheet.set_paper(9)
    sheet.fit_to_pages(1, 0)
    sheet.set_footer("&R&F\r&R&P", {'margin': 0.25})

    # text formaats
    global cell_formats

    # cell formats array
    cell_formats = {
        0: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 0}),     # default
        1: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),   # address
        2: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),     # name
        3: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),     # range
        4: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),   # value, g.A
        5: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),   # value, g.B
        6: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),   # value, g.C
        7: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),   # value, g.D
        8: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1}),     # desc
        9: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'font_size': 13, 'bg_color': '#EEECE1', 'border': 1}),  # h1
       10: book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#EEECE1', 'border': 1}),  # h2
       11: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#EEECE1', 'border': 1}), # h3
       12: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#EEECE1', 'top': 1, 'left': 1, 'right': 1, 'bottom': 0}),  # h3 without bottom border
       13: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#EEECE1', 'top': 0, 'left': 1, 'right': 1, 'bottom': 1}),  # h3 without top border
       14: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': 'yellow', 'border': 1}),  # highlight address
       15: book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'color': 'white', 'font_size': '3', 'border': 1})   # hide address
    }

    # set column width and formats
    sheet.set_column(0, 0, 6)  # address
    sheet.set_column(1, 1, 23)  # name
    sheet.set_column(2, 2, 30)  # range
    sheet.set_column(3, 6, 8.3)  # values in groups
    sheet.set_column(7, 7, 41.6)  # description

    return


'''
 insert chapter header
'''
def print_h1(text):

    global cur_row, last_h1_title, frm_h1

    # check text to titles_correct
    text = config_tree['titles_correct'].get(text, text)

    sheet.merge_range(cur_row, 0, cur_row, 7, text, cell_formats[9])
    cur_row = cur_row + 1

    # save last h1 header for titles_correct inn PrintH2
    last_h1_title = text

    return


'''
 insert chapter sub-header
'''
def print_h2(text):

    if text != "" and text != last_h1_title:
        global cur_row, frm_h2

        # check text to titles_correct
        text = config_tree['titles_correct'].get(last_h1_title + "|" + text, text)

        sheet.merge_range(cur_row, 0, cur_row, 7, text, cell_formats[10])
        cur_row = cur_row + 1

    return


def print_h3(title_prefix='', groups=False):
    """ insert parameter groups header
    """

    global cur_row, cell_formats

    if title_prefix != '':
        if groups:                                                                                     ### 3 rows
            sheet.merge_range(cur_row, 0, cur_row + 2, 0, "Адрес", cell_formats[11])                     # column 1
            sheet.merge_range(cur_row, 1, cur_row + 2, 1, "Параметр", cell_formats[11])                  # column 2
            sheet.write(cur_row, 2, "Значение/диапазон/шаг", cell_formats[12])                           # column 3
            sheet.merge_range(cur_row + 1, 2, cur_row + 2, 2, "(вторичные величины)", cell_formats[13])
            sheet.merge_range(cur_row, 3, cur_row, 6, "Задаваемый параметр", cell_formats[12])           # column 4
            sheet.merge_range(cur_row + 1, 3, cur_row + 1, 6, title_prefix, cell_formats[13])
            sheet.write(cur_row + 2, 3, "Группа A", cell_formats[11])
            sheet.write(cur_row + 2, 4, "Группа B", cell_formats[11])
            sheet.write(cur_row + 2, 5, "Группа C", cell_formats[11])
            sheet.write(cur_row + 2, 6, "Группа D", cell_formats[11])
            sheet.merge_range(cur_row, 7, cur_row + 2, 7, "Комментарий", cell_formats[11])               # column 5
            cur_row = cur_row + 3
        else:                                                                                          ### 2 rows
            sheet.merge_range(cur_row, 0, cur_row + 1, 0, "Адрес", cell_formats[11])                     # column 1
            sheet.merge_range(cur_row, 1, cur_row + 1, 1, "Параметр", cell_formats[11])                  # column 2
            sheet.write(cur_row, 2, "Значение/диапазон/шаг", cell_formats[12])                           # column 3
            sheet.write(cur_row + 1, 2, "(вторичные величины)", cell_formats[13])
            sheet.merge_range(cur_row, 3, cur_row, 6, "Задаваемый параметр", cell_formats[12])           # column 4
            sheet.merge_range(cur_row + 1, 3, cur_row + 1, 6, title_prefix, cell_formats[13])
            sheet.merge_range(cur_row, 7, cur_row + 1, 7, "Комментарий", cell_formats[11])               # column 5
            cur_row = cur_row + 2
    else:
        if groups:                                                                                     ### 2 rows
            sheet.merge_range(cur_row, 0, cur_row + 1, 0, "Адрес", cell_formats[11])                     # column 1
            sheet.merge_range(cur_row, 1, cur_row + 1, 1, "Параметр", cell_formats[11])                  # column 2
            sheet.merge_range(cur_row, 2, cur_row + 1, 2, "Значение/диапазон/шаг", cell_formats[11])     # column 3
            sheet.merge_range(cur_row, 3, cur_row, 6, "Задаваемый параметр", cell_formats[11])           # column 4
            sheet.write(cur_row + 1, 3, "Группа A", cell_formats[11])
            sheet.write(cur_row + 1, 4, "Группа B", cell_formats[11])
            sheet.write(cur_row + 1, 5, "Группа C", cell_formats[11])
            sheet.write(cur_row + 1, 6, "Группа D", cell_formats[11])
            sheet.merge_range(cur_row, 7, cur_row + 1, 7, "Комментарий", cell_formats[11])               # column 5
            cur_row = cur_row + 2
        else:                                                                                          ### 1 rows
            sheet.write(cur_row, 0, "Адрес", cell_formats[11])                                           # column 1
            sheet.write(cur_row, 1, "Параметр", cell_formats[11])                                        # column 2
            sheet.write(cur_row, 2, "Значение/диапазон/шаг", cell_formats[11])                           # column 3
            sheet.merge_range(cur_row, 3, cur_row, 6, "Задаваемый параметр", cell_formats[11])           # column 4
            sheet.write(cur_row, 7, "Комментарий", cell_formats[11])                                     # column 5
            cur_row = cur_row + 1

    return


'''
 update column header if chapter have electric paramaters 
'''


def update_column_header(RowNo, addtext_range="", addtext_value=""):
    if addtext_range != "":
        sheet.Cells(RowNo, 3).Value = sheet.Cells(RowNo, 3).Value + "\r\n(" + addtext_range + ")"
    if addtext_value != "":
        sheet.Cells(RowNo, 4).Value = sheet.Cells(RowNo, 4).Value + "\r\n(" + addtext_value + ")"
    sheet.Rows(RowNo).RowHeight = 32
    return


'''
 insert group header if needed
'''


# добавление текста групп уставок в шапку таблицы
def insert_group_header(RowNo):
    global cur_row
    global group_has_elec_values

    sheet.Cells(RowNo, 1).EntireRow.Insert(1)
    sheet.Rows(RowNo).RowHeight = 15

    # велосипед на велосипеде !!!
    if group_has_elec_values:
        sheet.Rows(RowNo - 1).RowHeight = 32
    else:
        sheet.Rows(RowNo - 1).RowHeight = 22

    sheet.Cells(RowNo, 4).Value = "Группа A"
    sheet.Cells(RowNo, 5).Value = "Группа B"
    sheet.Cells(RowNo, 6).Value = "Группа C"
    sheet.Cells(RowNo, 7).Value = "Группа D"

    # sheet.Rows.AutoFit()

    sheet.Range("A" + str(RowNo - 1) + ":A" + str(RowNo)).Merge()
    sheet.Range("B" + str(RowNo - 1) + ":B" + str(RowNo)).Merge()
    sheet.Range("C" + str(RowNo - 1) + ":C" + str(RowNo)).Merge()
    sheet.Range("H" + str(RowNo - 1) + ":H" + str(RowNo)).Merge()

    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).HorizontalAlignment = -4108  # win32.constants.xlCenter
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.Pattern = 1  # win32.constants.xlSolid
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.ThemeColor = 1  # win32.constants.xlThemeColorDark1
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.TintAndShade = -0.149998474074526
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Font.Bold = True

    cur_row = cur_row + 1

    return


'''
 get parameter info from XRio file
'''
def extract_parameter_name(Address):

    global xrio_tree

    ParameterName = xrio_tree.xpath("//ForeignId[text()='" + Address + "']/parent::*/Name/text()")
    if (ParameterName != None) and (len(ParameterName) > 0):
        ParameterName = str(ParameterName[0])
    else:
        ParameterName = ""

    return ParameterName


'''
 get parameter precision from XRio file
'''
def extract_parameter_precision(Address):

    global xrio_tree

    ParameterPrecision = xrio_tree.xpath("//ForeignId[text()='" + Address + "']/parent::*/Unit")
    if (ParameterPrecision != None) and (len(ParameterPrecision) > 0):
        return int(ParameterPrecision[0].attrib['DecimalPlaces'])
    else:
        return 0


'''
 convert electrical value to primary
'''
def convert_to_primary(Address, Value, Dimension, SecondaryPrecision):

    # do not convert special addresses

    global group_has_elec_values

    # this is a number
    rez = Value
    if re.search(r"\d+(\.|)\d*", Value, re.MULTILINE):
        Value = float(Value)
        if Dimension == "А":
            rez = "%g" % round(Value * ktt, SecondaryPrecision - 1) + " " + Dimension
            group_has_elec_values = True
        elif Dimension == "В":
            rez = "%g" % (Value * ktn / 1000) + " кВ"
            group_has_elec_values = True
        elif Dimension == "Ом":  # 2018-03-23: в 7SA в первичных 3 знака после запятой, в 7SD - два, везде делаем 3
            rez = "%g" % round(Value * ktn / ktt, SecondaryPrecision) + " " + Dimension
            group_has_elec_values = True
        elif Dimension == "Ом / км":
            rez = "%g" % round(Value * ktn / ktt, SecondaryPrecision - 1) + " " + Dimension
            group_has_elec_values = True
        elif Dimension == "ВА":
            rez = "%g" % round(Value * ktn * ktt / 1000000, SecondaryPrecision + 1) + " МВА"
            group_has_elec_values = True
        elif Dimension == "мкФ/км":
            rez = "%g" % round(Value * ktt / ktn, SecondaryPrecision + 1) + " " + Dimension
            group_has_elec_values = True
        else:
            rez = "%g" % float(rez) + " " + Dimension

    return str(rez)





'''
 extract parameter values in all groups of parameters
'''


def extract_parameter_values(Parameter):
    ParameterAddr = Parameter.attrib['DAdr']
    ParameterType = Parameter.attrib['Type']

    global group_has_group_values#, primary

    ParameterValue = Parameter.find(r"Value")
    ParameterValueA = Parameter.find(r"Value[@SettingGroup='A']")
    ParameterValueB = Parameter.find(r"Value[@SettingGroup='B']")
    ParameterValueC = Parameter.find(r"Value[@SettingGroup='C']")
    ParameterValueD = Parameter.find(r"Value[@SettingGroup='D']")

    if ParameterValueA == None:
        ParameterValueA = ParameterValue.text
        ParameterValueB = ParameterValue.text
        ParameterValueC = ParameterValue.text
        ParameterValueD = ParameterValue.text
    else:
        ParameterValueA = ParameterValueA.text
        ParameterValueB = ParameterValueB.text
        ParameterValueC = ParameterValueC.text
        ParameterValueD = ParameterValueD.text
        group_has_group_values = True

    if ParameterType == "Txt":
        ParameterValueA = Parameter.find(r"Comment[@Number='" + ParameterValueA + "']").attrib['Name']
        ParameterValueB = Parameter.find(r"Comment[@Number='" + ParameterValueB + "']").attrib['Name']
        ParameterValueC = Parameter.find(r"Comment[@Number='" + ParameterValueC + "']").attrib['Name']
        ParameterValueD = Parameter.find(r"Comment[@Number='" + ParameterValueD + "']").attrib['Name']
    else:
        Dimension = Parameter.find('Comment[@Dimension]')
        if Dimension != None:
            Dimension = Dimension.attrib.get('Dimension')
        else:
            Dimension = ''

        # convert to primary if needed
        if (primary == False) | (ParameterAddr in config_tree['params_without_convert']):
            # if value is "oo" - do not display dimension
            # call ConvertToPrimary for calc 'group_has_elec_values' variable
            convert_to_primary(ParameterAddr, ParameterValueA, Dimension, extract_parameter_precision(ParameterAddr))
            ParameterValueA = ParameterValueA if ParameterValueA == "oo" else ParameterValueA + " " + Dimension
            ParameterValueB = ParameterValueB if ParameterValueB == "oo" else ParameterValueB + " " + Dimension
            ParameterValueC = ParameterValueC if ParameterValueC == "oo" else ParameterValueC + " " + Dimension
            ParameterValueD = ParameterValueD if ParameterValueD == "oo" else ParameterValueD + " " + Dimension
        else:
            SecondaryPrecision = extract_parameter_precision(ParameterAddr)
            ParameterValueA = convert_to_primary(ParameterAddr, ParameterValueA, Dimension, SecondaryPrecision)
            ParameterValueB = convert_to_primary(ParameterAddr, ParameterValueB, Dimension, SecondaryPrecision)
            ParameterValueC = convert_to_primary(ParameterAddr, ParameterValueC, Dimension, SecondaryPrecision)
            ParameterValueD = convert_to_primary(ParameterAddr, ParameterValueD, Dimension, SecondaryPrecision)

    return [ParameterValueA.strip(), ParameterValueB.strip(), ParameterValueC.strip(), ParameterValueD.strip()]


'''
 extract parameter range
'''


def extract_parameter_range(Parameter):
    ParameterType = Parameter.attrib['Type']
    RangeText = ''
    Precision = 0
    if ParameterType == "Txt":
        Comments = Parameter.findall('Comment[@Name]')
        for Comment in Comments:
            if len(RangeText) != 0:
                RangeText = RangeText + "\r\n"
            if Comment.attrib['Name'] != '':
                RangeText = RangeText + Comment.attrib['Name']
    elif ParameterType == "Dec":
        Comment = Parameter.find('Comment')
        Dimension = Comment.attrib.get('Dimension')
        if Dimension == None:
            Dimension = ''
        MinValue = Comment.attrib['MinValue']
        MaxValue = Comment.attrib['MaxValue']
        Precision = len(MinValue) - MinValue.rfind(".")
        if Precision == len(MinValue) + 1:
            Precision = 0
        AdditionalValidValues = Comment.attrib.get('AdditionalValidValues')
        if AdditionalValidValues == None:
            RangeText = MinValue + " … " + MaxValue + " " + Dimension
        else:
            RangeText = MinValue + " … " + MaxValue + " " + Dimension + "; " + AdditionalValidValues

    return [RangeText, Precision]


'''
 paste parameter info to output excel sheet
'''


def print_parameter_data(ParameterData, highlight=False):

    global cur_row, last_printed_address

    # write data from config_tree then correct it by "params_correct" config_tree section
    if len(ParameterData['Address']) > 6:
        sheet.write(cur_row, 0, ParameterData['Address'], cell_formats[15])
    elif highlight:
        sheet.write(cur_row, 0, ParameterData['Address'], cell_formats[14])
    else:
        sheet.write(cur_row, 0, ParameterData['Address'], cell_formats[1])
    sheet.write(cur_row, 1, ParameterData['Name'], cell_formats[2])
    sheet.write(cur_row, 2, ParameterData['Range'], cell_formats[3])

    ParameterValues = ParameterData['Values']

    # if values are equal merge cells
    if ParameterValues[0] == ParameterValues[1] == ParameterValues[2] == ParameterValues[3]:
        sheet.merge_range(cur_row, 3, cur_row, 6, ParameterValues[0], cell_formats[4])
    elif ParameterValues[0] == ParameterValues[1] == ParameterValues[2]:
        sheet.merge_range(cur_row, 3, cur_row, 5, ParameterValues[0], cell_formats[4])
        sheet.write(cur_row, 6, ParameterValues[3], cell_formats[7])
    elif ParameterValues[0] == ParameterValues[1]:
        sheet.merge_range(cur_row, 3, cur_row, 4, ParameterValues[0], cell_formats[4])
        if ParameterValues[2] == ParameterValues[3]:
            sheet.merge_range(cur_row, 5, cur_row, 6, ParameterValues[2], cell_formats[6])
        else:
            sheet.write(cur_row, 5, ParameterValues[2], cell_formats[6])
            sheet.write(cur_row, 6, ParameterValues[3], cell_formats[7])
    elif ParameterValues[1] == ParameterValues[2] == ParameterValues[3]:
        sheet.write(cur_row, 3, ParameterValues[0], cell_formats[4])
        sheet.merge_range(cur_row, 4, cur_row, 6, ParameterValues[1], cell_formats[5])
    elif ParameterValues[2] == ParameterValues[3]:
        sheet.write(cur_row, 3, ParameterValues[0], cell_formats[4])
        sheet.write(cur_row, 4, ParameterValues[1], cell_formats[5])
        sheet.merge_range(cur_row, 5, cur_row, 6, ParameterValues[2], cell_formats[6])
    elif ParameterValues[1] == ParameterValues[2]:
        sheet.write(cur_row, 3, ParameterValues[0], cell_formats[4])
        sheet.merge_range(cur_row, 4, cur_row, 5, ParameterValues[1], cell_formats[5])
        sheet.write(cur_row, 6, ParameterValues[3], cell_formats[7])
    else:
        sheet.write(cur_row, 3, ParameterValues[0], cell_formats[4])
        sheet.write(cur_row, 4, ParameterValues[1], cell_formats[5])
        sheet.write(cur_row, 5, ParameterValues[2], cell_formats[6])
        sheet.write(cur_row, 6, ParameterValues[3], cell_formats[7])

    sheet.write(cur_row, 7, ParameterData['Description'], cell_formats[8])

    # and correct (or add new)
    addr = ParameterData['Address']
    need_correct = config_tree["params_correct"].get(addr, None)
    if need_correct != None:
        col_no = config_tree["params_correct"].get(addr, None)[0]
        col_val = config_tree["params_correct"].get(addr, None)[1]
        sheet.write(cur_row, int(col_no), col_val, cell_formats[col_no+1] if col_no in range(0, 7) else cell_formats[0])

    cur_row = cur_row + 1
    last_printed_address = ParameterData['Address']

    # insert formula with comments
    # !!!
    # sheet.Cells(currow, 9).FormulaR1C1 = '=IFERROR(IF(TRIM(RC[-8])<>"",INDEX(\'\\\\Prim-fs-serv\\rdu\СРЗА\\Уставки\\РАСЧЕТЫ УСТАВОК\\[!!!Siemens, общие комментарии.xlsx]7SD\'!C1:C2,MATCH(RC[-8],\'\\\\Prim-fs-serv\\rdu\\СРЗА\\Уставки\\РАСЧЕТЫ УСТАВОК\\[!!!Siemens, общие комментарии.xlsx]7SD\'!C1,0),2),""),"")'
    # sheet.Cells(currow, 9).VerticalAlignment = -4108 # xlCenter

    pass


'''
 stash parameters for rearrange, push
'''


def stash_parameters_push(ParameterData):
    global stash

    PopAfter = config_tree['params_to_rearrange'].get(ParameterData['Address'], 0)
    if (PopAfter == 0):
        return False
    else:
        stash.append({
            'PopAfter': PopAfter,
            'ParameterData': ParameterData
        })
        return True


'''
 stash parameters for rearrange, pop (for stashed parameter with PopAfter number)
'''


def stash_parameters_pop(ParameterAddress):
    global stash

    for i in range(len(stash)):
        if (stash[i]['PopAfter'] == ParameterAddress):  # past stashed parameter with PopAfter number
            ParameterData = stash[i]['ParameterData']
            stash.pop(i)
            print_parameter_data(ParameterData)
            return ParameterData['Address']

    return False


'''
 stash parameters for rearrange, pop (for stashed parameter with PopAfter = "auto")
 past before current parameter
'''


def stash_parameters_pop_auto(ParameterAddress):
    global stash

    for i in range(len(stash)):
        if ((stash[i]['PopAfter'].lower() == 'auto') &  # auto rearrange
                (int(re.sub('[^\d]', '', ParameterAddress)) > int(
                    re.sub('[^\d]', '', stash[i]['ParameterData']['Address'])))):
            ParameterData = stash[i]['ParameterData']
            stash.pop(i)
            print_parameter_data(ParameterData)
            return ParameterData['Address']

    return False


def insert_parameter(parameter_data, rearrange=False):
    """ precess one parameter / address
    """

    global stash

    # print current address if this address is absent in stash
    if ((not (any(p['ParameterData']['Address'] == parameter_data['Address'] for p in stash))) |
        (rearrange)):
        print_parameter_data(parameter_data, rearrange)

        # rearrange: Num (post after current address)
        for i in range(len(stash)):
            if stash[i]['PopAfter'] == parameter_data['Address']:
                insert_parameter(stash[i]['ParameterData'], True)
                stash[i]['Inserted'] = True

    return

def process_parameter(parameter):
    """ precess one parameter / address
    """

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

    # print parameter to report and stashed params for rearrange if it need
    global stash
    global inserted_stash

    # rearrange: auto (post before current address)
    need_loop = True
    while need_loop:
        need_loop = False
        for i in range(len(stash)):
            if (stash[i]['PopAfter'].lower() == 'auto') & \
                    (int(re.sub('[^\d]', '', stash[i]['ParameterData']['Address'])) <
                     int(re.sub('[^\d]', '', address))) & \
                    (int(re.sub('[^\d]', '', stash[i]['ParameterData']['Address'])) >
                     int(re.sub('[^\d]', '', last_printed_address))) & \
                    (stash[i]['ParameterData']['Address'] not in inserted_stash) & \
                    (not (any(p['ParameterData']['Address'] == address for p in stash))):

                print_parameter_data(stash[i]['ParameterData'], True)
                inserted_stash.append(stash[i]['ParameterData']['Address'])
                need_loop = True


    # print current address parameter
    if not (any(p['ParameterData']['Address'] == address for p in stash)):
        print_parameter_data(parameter_data)

    # rearrange: Num (post after current address)
    for i in range(len(stash)):
        if stash[i]['PopAfter'] == address:
            print_parameter_data(stash[i]['ParameterData'], True)
            inserted_stash.append(stash[i]['ParameterData']['Address'])

    return


def is_elec_value(dimension):
    return dimension in ["А", "В", "Ом", "Ом / км", "ВА", "мкФ/км"]


def process_setting_page(setting_page):
    """process settings page
    """
    global cur_row, primary

    setting_page_name = setting_page.attrib['Name']
    print_h2(setting_page_name)

    parameters = setting_page.findall("Parameter")

    page_have_elec_param = False
    page_have_param_to_convert = False
    page_have_groups = False

    for parameter in parameters:
        if parameter.attrib['Type'] != "Txt":
            dimension = parameter.find('Comment[@Dimension]')
            if dimension != None:
                dimension = dimension.attrib.get('Dimension')
            else:
                dimension = ''
            if dimension in ["А", "В", "Ом", "Ом / км", "ВА", "мкФ/км"]:
                page_have_elec_param = True
                if not parameter.attrib['DAdr'] in config_tree['params_without_convert']:
                    page_have_param_to_convert = True
            if parameter.find(r"Value[@SettingGroup='A']") is not None:
                page_have_groups = True

    if page_have_elec_param & page_have_param_to_convert:
        if primary:
            print_h3('(первичные величины)', page_have_groups)
        else:
            print_h3('(вторичные величины)', page_have_groups)
    else:
        print_h3('', page_have_groups)

    for parameter in parameters:
        process_parameter(parameter)

    return


'''
 process function group
'''


def process_function_group(FunctionGroup):
    FunctionGroupName = FunctionGroup.attrib['Name']
    print_h1(FunctionGroupName)
    SettingPages = FunctionGroup.findall("SettingPage")
    for SettingPage in SettingPages:
        process_setting_page(SettingPage)

    return


'''
  process all XML file and extract params for rearrange to stash list
'''
def extract_parameters_to_rearrange():

    global xrio_tree, xml_tree
    global config_tree
    global stash

    all_params = xml_tree.findall('Settings//Parameter')
    for p in all_params:
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
            stash.append({
                'PopAfter': pop_after,
                'ParameterData': parameter_data,
                'Inserted': False
            })

    # sort stash list for insert in normal order
    stash.sort(key=lambda p: p['ParameterData']['Address'])

    return


'''
 register .xrio extention for siemens py 
'''
def register_xrio_ext():

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


'''
  Main work stats from print hello message
'''
print("SiemensPie - tool for convert Siemens Siprotec relay protection config")
print("to readable and editable format (Excel .xlsx file)")
print("(C) "+__version__+" Wyfinger, https://github.com/wyfinger/SiemensPie")
print("")

register_xrio_ext()
process_command_line()
create_output_file()
page_setup()
process_all()
book.close()

time.sleep(5)
