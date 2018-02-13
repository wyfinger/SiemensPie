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


__author__ = 'prim_miv'

xriofile = ""
xmlfile = ""
primary = False
config = {}
xrio_tree = None
xml_tree = None
ktt = 1
ktn = 1
book = Workbook
sheet = Worksheet
frm_address = Format
frm_name = Format
frm_range = Format
frm_values = Format
frm_desc = Format
frm_h1 = Format
frm_h2 = Format
frm_h = Format
cur_row = 0
group_has_elec_values = False
group_has_group_values = False
stash = []
last_h1_title = ""

#TODO: уставка 7137 в 7SJ не выводится в Excel, хотя в конфиге, если посмотреть в Digsi, есть.
#TODO: уставка 1124 в 7SD, описание "Центральная фаза", мы обычно дописываем "... присоединения", поэтому
#      нужен механизм правки через конфиг не только имени уставки, но и других (любых) столбцов бланка.



# print small help tip to console, for use in error in parameters
def PrintSmallHelp():

    print("Use: sp.exe [-p] [-c] [xml or xrio file] [xml or xrio file]")
    print("  -p  - convert electrical values to primary")
    print("  -c  - path to config file (Json)")
    print("  set one (xml or xrio) file if they have the same name")

    return


'''
 read config Json
'''
def ReadConfig(config_path):

    try:
        #print('Reading config from: ' + config_path)
        with codecs.open(config_path, 'r', 'utf-8') as param_file:
            return json.load(param_file)
    except:
        print("Error at read config file.\n")
        PrintSmallHelp()
        sys.exit()

    return


'''
 command line parameter analyses
'''
def ProcessCommandLine():

    global xriofile, xmlfile, primary, config

    if getattr(sys, 'frozen', False):
        config_path = os.path.join(os.path.dirname(sys.executable), "config.json")
    else:
        config_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "config.json")

    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--primary', nargs='?', default='true', help='Convert values to a prymary')
    parser.add_argument('-c', '--config', nargs='?', default=config_path, help='Json file with parameters')
    parser.add_argument('file1', nargs='?')
    parser.add_argument('file2', nargs='?')
    namespace = parser.parse_args()

    # convert params to primary? in's stored in secondary
    primary = (namespace.primary != 'false')

    # find .xml and .xrio files
    if (namespace.file1 != None):
        file1name, file1extension = os.path.splitext(namespace.file1)
        if (file1extension.lower() == '.xrio'):
            xriofile = namespace.file1
        if (file1extension.lower() == '.xml'):
            xmlfile = namespace.file1
    if (namespace.file2 != None):
        file2name, file2extension = os.path.splitext(namespace.file2)
        if (file2extension.lower() == '.xrio'):
            xriofile = namespace.file2
        if (file2extension.lower() == '.xml'):
            xmlfile = namespace.file2

    if (xriofile == "") & (xmlfile != ""):
        xriofile = os.path.splitext(xmlfile)[0] + '.xrio'
        if (os.path.isfile(xriofile) == False):
            xriofile = ""
    if (xriofile != "") & (xmlfile == ""):
        xmlfile = os.path.splitext(xriofile)[0] + '.xml'
        if (os.path.isfile(xmlfile) == False):
            xmlfile = ""

    if namespace.config != None:
        config_path = namespace.config
    config = ReadConfig(config_path)

    if (xmlfile == "") | (xriofile == ""):
        print("Error. XML or XRio file is not exists.")
        PrintSmallHelp()
        sys.exit()

    print("XML: " + xmlfile)
    print("XRio: " + xriofile)

    return


'''
 calculate crc32 of this file
'''
def SelfCRC32():
    buf = open(sys.argv[0],'rb').read()
    buf = (binascii.crc32(buf) & 0xFFFFFFFF)
    return "%08X" % buf


'''
 create outpot excel file
'''
def CreateOutputFile():

    global book, sheet

    try:
        xlsx_path, ext = os.path.splitext(xmlfile)
        book = Workbook(xlsx_path + '.xlsx')
        sheet = book.add_worksheet()
    except:
        print("Error at create Excel file to output.\n")
        sys.exit()

    return


'''
 start of data process
'''
def ProcessAll():

    global xrio_tree, xml_tree
    global config, ktt, ktn, cur_row

    # load xml files
    try:
        xrio_tree = lxml.parse(xriofile)
        xml_tree = lxml.parse(xmlfile)
    except:
        print("Error at read XML and XRio files.\n")
        PrintSmallHelp()
        sys.exit()

    # select config section by MLFB code
    MLFBDIGSI = xml_tree.xpath('.//General/GeneralData[@Name="MLFBDIGSI"]/@ID')[0]
    for k in config.keys():
        if (MLFBDIGSI[0:len(k)]==k):
            config = config[k]
            break

    # prepare ktt and ktn values, only if primary = true
    if primary:
        v_primary = xml_tree.xpath(config['voltage_primary'])[0]
        v_primary = float(re.sub(r"[^\d+.]", "", v_primary, 0, 0))*1000 # voltage in kilovolts
        v_second = xml_tree.xpath(config['voltage_second'])[0]
        v_second = float(re.sub(r"[^\d+.]", "", v_second, 0, 0))
        ktn = v_primary / v_second
        c_primary = xml_tree.xpath(config['current_primary'])[0]
        c_primary = int(re.sub(r"[^\d+\.]", "", c_primary, 0, 0))
        c_second = xml_tree.xpath(config['current_second'])[0]
        c_second = int(re.sub(r"[^\d+\.]", "", c_second, 0, 0))
        ktt = c_primary / c_second

    # paste overview info about terminal
    # MLFB code
    sheet.merge_range(cur_row, 0, cur_row, 1, "MLFB Код", frm_name)
    sheet.merge_range(cur_row, 2, cur_row, 7, MLFBDIGSI, frm_name)
    cur_row = cur_row + 1
    # Version
    sheet.merge_range(cur_row, 0, cur_row, 1, "Версия ПО терминала", frm_name)
    sheet.merge_range(cur_row, 2, cur_row, 7, xml_tree.xpath('.//General/GeneralData[@Name="Version"]/@ID')[0], frm_name)
    cur_row = cur_row + 1
    # Topology
    sheet.merge_range(cur_row, 0, cur_row, 1, "Топология", frm_name)
    sheet.merge_range(cur_row, 2, cur_row, 7, xml_tree.xpath('.//General/GeneralData[@Name="Topology"]/@ID')[0], frm_name)
    cur_row = cur_row + 1
    # Self version (crc32)
    sheet.merge_range(cur_row, 0, cur_row, 1, "Версия SiemensPie", frm_name)
    sheet.merge_range(cur_row, 2, cur_row, 7, SelfCRC32(), frm_name)
    cur_row = cur_row + 1

    # main work
    FunctionGroups = xml_tree.findall('Settings/FunctionGroup')
    for FunctionGroup in FunctionGroups:
        ProcessFunctionGroup(FunctionGroup)

    return


'''
 insert header and page stylization 
'''
def PageSetup():

    # page margins, headers and footers
    sheet.set_margins(0.4, 0.4, 0.4, 0.4)
    sheet.set_header("", {'margin': 0.12})
    sheet.set_footer("&amp;R&amp;F\n&amp;P", {'margin': 0.12})
    sheet.set_zoom(90)
    sheet.set_landscape()
    sheet.fit_to_pages(1, 0)

    # text formats
    global frm_address, frm_name, frm_range, frm_values, frm_desc, frm_h1, frm_h2, frm_h
    frm_address = book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    frm_name = book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
    frm_range = book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
    frm_values = book.add_format({'align': 'center', 'valign': 'vcenter','text_wrap': True, 'num_format': '@'})
    frm_desc = book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
    frm_desc = book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
    frm_h1 = book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'font_size': 13, 'bg_color': '#D9D9D9'})
    frm_h2 = book.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#D9D9D9'})
    frm_h = book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True, 'bg_color': '#D9D9D9'})

    # set column width and formats
    sheet.set_column(0, 0, 6, frm_address)    # address
    sheet.set_column(1, 1, 23, frm_name)      # name
    sheet.set_column(2, 2, 30, frm_range)     # range
    sheet.set_column(3, 6, 8.3, frm_values)   # values in groups
    sheet.set_column(7, 7, 41.6, frm_desc)    # description

    return


'''
 insert chapter header
'''
def PrintH1(text):

    global cur_row, last_h1_title, frm_h1

    # check text to titles_correct
    text = config['titles_correct'].get(text, text)

    sheet.merge_range(cur_row, 0, cur_row, 7, text, frm_h1)
    cur_row = cur_row + 1

    # save last h1 header for titles_correct inn PrintH2
    last_h1_title = text

    return


'''
 insert chapter sub-header
'''
def PrintH2(text):

    if text != "" and text != last_h1_title:

        global cur_row, frm_h2

        # check text to titles_correct
        text = config['titles_correct'].get(last_h1_title+"|"+text, text)

        sheet.merge_range(cur_row, 0, cur_row, 7, text, frm_h2)
        cur_row = cur_row + 1

    return


'''
 insert header with column titles
'''
def PrintGroupHeader(at_row):

    sheet.merge_range(at_row, 0, at_row + 1, 0, "№\r\nАдрес", frm_h)
    sheet.merge_range(at_row, 1, at_row + 1, 1, "Наименование уставки", frm_h)

    if group_has_elec_values:
        sheet.write(at_row, 2, "Диапазон уставок", frm_h)
        sheet.write(at_row + 1, 2, "(вторичные величины)", frm_h)
    else:
        sheet.merge_range(at_row, 2, at_row + 1, 2, "Диапазон уставок", frm_h)

    if group_has_group_values:
        if group_has_elec_values:
            sheet.set_row(at_row, 30)
            if primary:
                sheet.merge_range(at_row, 3, at_row, 6, "Заданная уставка\r\n(первичные величины)", frm_h)
            else:
                sheet.merge_range(at_row, 3, at_row, 6, "Заданная уставка\r\n(вторичные величины)", frm_h)
        else:
            sheet.merge_range(at_row, 3, at_row, 6, "Заданная уставка", frm_h)
        sheet.write(at_row + 1, 3, "Группа A", frm_h)
        sheet.write(at_row + 1, 4, "Группа B", frm_h)
        sheet.write(at_row + 1, 5, "Группа C", frm_h)
        sheet.write(at_row + 1, 6, "Группа D", frm_h)
    else:
        if group_has_elec_values:
            if primary:
                sheet.merge_range(at_row, 3, at_row, 6, "Заданная уставка", frm_h)
                sheet.merge_range(at_row + 1, 3, at_row + 1, 6, "(первичные величины)", frm_h)
            else:
                sheet.merge_range(at_row, 3, at_row, 6, "Заданная уставка\r\n(вторичные величины)", frm_h)
                sheet.merge_range(at_row + 1, 3, at_row + 1, 6, "(вторичные величины)", frm_h)
        else:
            sheet.merge_range(at_row, 3, at_row + 1, 6, "Заданная уставка", frm_h)

    sheet.merge_range(at_row, 7, at_row + 1, 7, "Описание", frm_h)

    return


'''
 update column header if chapter have electric paramaters 
'''
def UpdateColumnHeader(RowNo, addtext_range="", addtext_value=""):

    if addtext_range != "":
        sheet.Cells(RowNo,3).Value = sheet.Cells(RowNo,3).Value + "\r\n("+addtext_range+")"
    if addtext_value != "":
        sheet.Cells(RowNo,4).Value = sheet.Cells(RowNo,4).Value + "\r\n("+addtext_value+")"
    sheet.Rows(RowNo).RowHeight = 32
    return


'''
 insert group header if needed
'''
# добавление текста групп уставок в шапку таблицы
def InsertGroupHeader(RowNo):

    global cur_row
    global group_has_elec_values

    sheet.Cells(RowNo, 1).EntireRow.Insert(1)
    sheet.Rows(RowNo).RowHeight = 15

    # велосипед на велосипеде !!!
    if group_has_elec_values:
        sheet.Rows(RowNo-1).RowHeight = 32
    else:
        sheet.Rows(RowNo - 1).RowHeight = 22

    sheet.Cells(RowNo, 4).Value = "Группа A"
    sheet.Cells(RowNo, 5).Value = "Группа B"
    sheet.Cells(RowNo, 6).Value = "Группа C"
    sheet.Cells(RowNo, 7).Value = "Группа D"

    # sheet.Rows.AutoFit()

    sheet.Range("A" + str(RowNo-1) + ":A" + str(RowNo)).Merge()
    sheet.Range("B" + str(RowNo-1) + ":B" + str(RowNo)).Merge()
    sheet.Range("C" + str(RowNo-1) + ":C" + str(RowNo)).Merge()
    sheet.Range("H" + str(RowNo-1) + ":H" + str(RowNo)).Merge()

    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).HorizontalAlignment = -4108 # win32.constants.xlCenter
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.Pattern = 1 # win32.constants.xlSolid
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.ThemeColor = 1 #win32.constants.xlThemeColorDark1
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.TintAndShade = -0.149998474074526
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Font.Bold = True

    cur_row = cur_row + 1

    return


'''
 get parameter info from XRio file
'''
def ExtractParameterName(Address):

    global xrio_tree

    ParameterName = xrio_tree.xpath("//ForeignId[text()='"+Address+"']/parent::*/Name/text()")
    if (ParameterName != None) and (len(ParameterName) > 0):
        ParameterName = str(ParameterName[0])
    else:
        ParameterName= ""

    return ParameterName


'''
 get parameter precision from XRio file
'''
def ExtractParameterPrecision(Address):

    global xrio_tree

    ParameterPrecision = xrio_tree.xpath("//ForeignId[text()='"+Address+"']/parent::*/Unit")
    if (ParameterPrecision != None) and (len(ParameterPrecision) > 0):
        return int(ParameterPrecision[0].attrib['DecimalPlaces'])
    else:
        return 0


'''
 convert electrical value to primary
'''
def ConvertToPrimary(Address, Value, Dimension, SecondaryPrecision):

    # do not convert special addresses
    if (Address in config['params_without_convert']):
        return "%g" % float(Value) + " " + Dimension

    global group_has_elec_values

    # this is a number
    rez = Value
    if re.search(r"\d+(\.|)\d*", Value, re.MULTILINE):
        Value = float(Value)
        if (Dimension == "А"):
            rez = "%g" % round(Value * ktt, SecondaryPrecision - 1) + " " + Dimension
            group_has_elec_values = True
        elif (Dimension=="В"):
            rez = "%g" % (Value * ktn / 1000) + " кВ"
            group_has_elec_values = True
        elif (Dimension=="Ом"):
            rez = "%g" % round(Value * ktn / ktt, SecondaryPrecision - 1) + " " + Dimension
            group_has_elec_values = True
        elif (Dimension=="Ом / км"):
            rez = "%g" % round(Value * ktn / ktt, SecondaryPrecision - 2) + " " + Dimension
            group_has_elec_values = True
        elif (Dimension=="ВА"):
            rez = "%g" % round(Value * ktn * ktt / 1000000, SecondaryPrecision + 1) + " МВА"
            group_has_elec_values = True
        elif (Dimension=="мкФ/км"):
            rez = "%g" % round(Value * ktt / ktn, SecondaryPrecision + 1) + " " + Dimension
            group_has_elec_values = True
        else:
            rez = "%g" % float(rez) + " " + Dimension

    return str(rez)


'''
 extract parameter values in all groups of parameters
'''
def ExtractParameterValues(Parameter):

    ParameterAddr = Parameter.attrib['DAdr']
    ParameterType = Parameter.attrib['Type']

    global group_has_group_values

    ParameterValue  = Parameter.find(r"Value")
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
        ParameterValueA = Parameter.find(r"Comment[@Number='"+ParameterValueA+"']").attrib['Name']
        ParameterValueB = Parameter.find(r"Comment[@Number='"+ParameterValueB+"']").attrib['Name']
        ParameterValueC = Parameter.find(r"Comment[@Number='"+ParameterValueC+"']").attrib['Name']
        ParameterValueD = Parameter.find(r"Comment[@Number='"+ParameterValueD+"']").attrib['Name']
    else:
        Dimension = Parameter.find('Comment[@Dimension]')
        if Dimension!=None:
            Dimension = Dimension.attrib.get('Dimension')
        else:
            Dimension = ''

        # convert to primary if needed
        if primary:
            SecondaryPrecision = ExtractParameterPrecision(ParameterAddr)
            ParameterValueA = ConvertToPrimary(ParameterAddr, ParameterValueA, Dimension, SecondaryPrecision)
            ParameterValueB = ConvertToPrimary(ParameterAddr, ParameterValueB, Dimension, SecondaryPrecision)
            ParameterValueC = ConvertToPrimary(ParameterAddr, ParameterValueC, Dimension, SecondaryPrecision)
            ParameterValueD = ConvertToPrimary(ParameterAddr, ParameterValueD, Dimension, SecondaryPrecision)
        else:
            # if value is "oo" - do not display dimension
            ParameterValueA = ParameterValueA if ParameterValueA == "oo" else ParameterValueA + " " + Dimension
            ParameterValueB = ParameterValueB if ParameterValueB == "oo" else ParameterValueB + " " + Dimension
            ParameterValueC = ParameterValueC if ParameterValueC == "oo" else ParameterValueC + " " + Dimension
            ParameterValueD = ParameterValueD if ParameterValueD == "oo" else ParameterValueD + " " + Dimension

    return [ParameterValueA, ParameterValueB, ParameterValueC, ParameterValueD]


'''
 extract parameter range
'''
def ExtractParameterRange(Parameter):

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
        if Dimension==None:
            Dimension = ''
        MinValue = Comment.attrib['MinValue']
        MaxValue = Comment.attrib['MaxValue']
        Precision = len(MinValue) - MinValue.rfind(".")
        if Precision == len(MinValue)+1:
            Precision = 0
        AdditionalValidValues = Comment.attrib.get('AdditionalValidValues')
        if AdditionalValidValues==None:
            RangeText = MinValue + " … " + MaxValue + " " + Dimension
        else:
            RangeText = MinValue + " … " + MaxValue + " " + Dimension + "; " + AdditionalValidValues

    return [RangeText, Precision]


'''
 paste parameter info to output excel sheet
'''
def PrintParameterData(ParameterData):

    global cur_row

    # write data from config then correct it by "params_correct" config section
    sheet.write(cur_row, 0, ParameterData['Address'], frm_address)
    sheet.write(cur_row, 1, ParameterData['Name'], frm_name)
    sheet.write(cur_row, 2, ParameterData['Range'], frm_range)

    ParameterValues = ParameterData['Values']

    # if values are equal merge cells
    if ParameterValues[0] == ParameterValues[1] == ParameterValues[2] == ParameterValues[3]:
        sheet.merge_range(cur_row, 3, cur_row, 6, ParameterValues[0], frm_values)
    elif ParameterValues[0] == ParameterValues[1] == ParameterValues[2]:
        sheet.merge_range(cur_row, 3, cur_row, 5, ParameterValues[0], frm_values)
        sheet.write(cur_row, 6, ParameterValues[3], frm_values)
    elif ParameterValues[0] == ParameterValues[1]:
        sheet.merge_range(cur_row, 3, cur_row, 4, ParameterValues[0], frm_values)
        if ParameterValues[2] == ParameterValues[3]:
            sheet.merge_range(cur_row, 5, cur_row, 6, ParameterValues[2], frm_values)
        else:
            sheet.write(cur_row, 5, ParameterValues[2], frm_values)
            sheet.write(cur_row, 6, ParameterValues[3], frm_values)
    elif ParameterValues[1] == ParameterValues[2] == ParameterValues[3]:
        sheet.write(cur_row, 3, ParameterValues[0], frm_values)
        sheet.merge_range(cur_row, 4, cur_row, 6, ParameterValues[1], frm_values)
    elif ParameterValues[2] == ParameterValues[3]:
        sheet.write(cur_row, 3, ParameterValues[0], frm_values)
        sheet.write(cur_row, 4, ParameterValues[1], frm_values)
        sheet.merge_range(cur_row, 5, cur_row, 6, ParameterValues[2], frm_values)
    elif ParameterValues[1] == ParameterValues[2]:
        sheet.write(cur_row, 3, ParameterValues[0], frm_values)
        sheet.merge_range(cur_row, 4, cur_row, 5, ParameterValues[1], frm_values)
        sheet.write(cur_row, 6, ParameterValues[3], frm_values)
    else:
        sheet.write(cur_row, 3, ParameterValues[0], frm_values)
        sheet.write(cur_row, 4, ParameterValues[1], frm_values)
        sheet.write(cur_row, 5, ParameterValues[2], frm_values)
        sheet.write(cur_row, 6, ParameterValues[3], frm_values)

    sheet.write(cur_row, 7, ParameterData['Description'], frm_desc)

    # and correct (or add new)
    addr = ParameterData['Address']
    need_correct = config["params_correct"].get(addr, None)
    if need_correct != None:
        col_no = config["params_correct"].get(addr, None)[0]
        col_val = config["params_correct"].get(addr, None)[1]
        sheet.write(cur_row, int(col_no), col_val)

    cur_row = cur_row + 1

    # insert formula with comments
    # !!!
    #sheet.Cells(currow, 9).FormulaR1C1 = '=IFERROR(IF(TRIM(RC[-8])<>"",INDEX(\'\\\\Prim-fs-serv\\rdu\СРЗА\\Уставки\\РАСЧЕТЫ УСТАВОК\\[!!!Siemens, общие комментарии.xlsx]7SD\'!C1:C2,MATCH(RC[-8],\'\\\\Prim-fs-serv\\rdu\\СРЗА\\Уставки\\РАСЧЕТЫ УСТАВОК\\[!!!Siemens, общие комментарии.xlsx]7SD\'!C1,0),2),""),"")'
    #sheet.Cells(currow, 9).VerticalAlignment = -4108 # xlCenter

    pass


'''
 stash parameters for rearrange, push
'''
def StashParametersPush(ParameterData):

    global stash

    PopAfter = config['params_to_rearrange'].get(ParameterData['Address'], 0)
    if (PopAfter == 0):
        return False
    else:
        stash.append({
            'PopAfter': PopAfter,
            'ParameterData': ParameterData
        })
        return True

'''
 stash parameters for rearrange, pop
'''
def StashParametersPop(ParameterAddress):

    global stash

    for i in range(len(stash)):
        if (stash[i]['PopAfter'] == ParameterAddress):
            ParameterData = stash[i]['ParameterData']
            stash.pop(i)
            PrintParameterData(ParameterData)
            return ParameterData['Address']

    return False


'''
 precess one parameter / address
'''
def ProcessParameter(Parameter):

    ParameterAddress = Parameter.attrib['DAdr']

    print(ParameterAddress)

    ParameterName = ExtractParameterName(ParameterAddress)
    ParameterDescription = Parameter.attrib['Name']

    Range = ExtractParameterRange(Parameter)
    ParameterRange = Range[0]
    ParameterValues = ExtractParameterValues(Parameter)

    ParameterData = {
        'Address': ParameterAddress,
        'Name': ParameterName,
        'Range': ParameterRange,
        'Values': ParameterValues,
        'Description': ParameterDescription,
    }

    # адреса, которые нужно переместить, если параметр не был спрятан - выведем его в Excel
    if (StashParametersPush(ParameterData) == False):
        PrintParameterData(ParameterData)
        # вставка сохраненных параметров
        while (ParameterAddress != False):
            ParameterAddress = StashParametersPop(ParameterAddress)

    return


'''
 process settings page
'''
def ProcessSettingPage(SettingPage):

    global cur_row, group_has_elec_values, group_has_group_values

    SettingPageName = SettingPage.attrib['Name']
    PrintH2(SettingPageName)

    Parameters = SettingPage.findall("Parameter")

    group_has_elec_values = False
    group_has_group_values = False
    header_row = cur_row
    cur_row = cur_row + 2

    for Parameter in Parameters:
        ProcessParameter(Parameter)

    PrintGroupHeader(header_row)

    return


'''
 process function group
'''
def ProcessFunctionGroup(FunctionGroup):

    FunctionGroupName = FunctionGroup.attrib['Name']
    PrintH1(FunctionGroupName)
    SettingPages = FunctionGroup.findall("SettingPage")
    for SettingPage in SettingPages:
        ProcessSettingPage(SettingPage)

    return


'''
 register .xrio extention for siemens py 
'''
def RegisterXrioExt():

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

    return

RegisterXrioExt()
ProcessCommandLine()
CreateOutputFile()
PageSetup()
ProcessAll()
book.close()
time.sleep(5)
