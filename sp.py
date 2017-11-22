#!/usr/bin/python
# -*- coding: utf-8 -*-
import argparse
import codecs
import json
import os
import re
import sys

import binascii
import lxml
from lxml import etree as lxml
import win32com.client as win32

__author__ = 'prim_miv'

xriofile = ""
xmlfile = ""
primary = False
config = {}

xrio_tree = None
xml_tree = None
ktt = 1
ktn = 1
excel = None
book = None
sheet = None
currow = 1
group_has_elec_values = False
group_has_group_values = False
stash = []
last_title = ""


#TODO: уставка 7137 в 7SJ не выводится в Excel, хотя в конфиге, если посмотреть в Digsi, есть.
from win32com.client import gencache
if gencache.is_readonly:
    gencache.is_readonly = False
    gencache.Rebuild() #create gen_py folder if needed


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

    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--primary', nargs='?', default='false', help='Convert values to a prymary')
    parser.add_argument('-c', '--config', nargs='?', default=os.path.dirname(sys.argv[0])+'/config.json', help='Json file with parameters')
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

    config_path = 'config.json'
    if (namespace.config != None):
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
 start of data process
'''
def ProcessAll():

    global xrio_tree, xml_tree

    # load xml files
    try:
        xrio_tree = lxml.parse(xriofile)
        xml_tree = lxml.parse(xmlfile)
    except:
        print("Error at read XML and XRio files.\n")
        PrintSmallHelp()
        sys.exit()

    global excel, book, sheet

    try:
        #excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = True
        book = excel.Workbooks.Add()
        sheet = book.Worksheets.Add()
    except:
        print("Error at create Exel file to output.\n")
        sys.exit()

    PageSetup()

    global config, ktt, ktn, currow

    # select config section by MLFB code
    MLFBDIGSI = xml_tree.xpath('.//General/GeneralData[@Name="MLFBDIGSI"]/@ID')[0]
    for k in config.keys():
        if (MLFBDIGSI[0:len(k)]==k):
            config = config[k]
            break

    # prepare ktt and ktn values, only if primary = true
    if (primary):
        v_primary = xml_tree.xpath(config['voltage_primary'])[0]
        v_primary = int(re.sub(r"[^\d+]", "", v_primary, 0, 0))*1000 # voltage in kilovolts
        v_second = xml_tree.xpath(config['voltage_second'])[0]
        v_second = int(re.sub(r"[^\d+]", "", v_second, 0, 0))
        ktn = v_primary / v_second
        c_primary = xml_tree.xpath(config['current_primary'])[0]
        c_primary = int(re.sub(r"[^\d+]", "", c_primary, 0, 0))
        c_second = xml_tree.xpath(config['current_second'])[0]
        c_second = int(re.sub(r"[^\d+]", "", c_second, 0, 0))
        ktt = c_primary / c_second

    # paste overview info about terminal
    # MLFB code
    sheet.Range("A" + str(currow) + ":B" + str(currow)).Merge()
    sheet.Range("C" + str(currow) + ":H" + str(currow)).Merge()
    sheet.Cells(currow, 1).Value = "MLFB Код"
    sheet.Cells(currow, 3).Value = MLFBDIGSI
    sheet.Cells(currow, 1).HorizontalAlignment = win32.constants.xlLeft
    currow = currow + 1
    # Version
    sheet.Range("A" + str(currow) + ":B" + str(currow)).Merge()
    sheet.Range("C" + str(currow) + ":H" + str(currow)).Merge()
    Version = xml_tree.xpath('.//General/GeneralData[@Name="Version"]/@ID')[0]
    sheet.Cells(currow, 1).Value = "Версия ПО терминала"
    sheet.Cells(currow, 3).Value = Version
    sheet.Cells(currow, 1).HorizontalAlignment = win32.constants.xlLeft
    currow = currow + 1
    # Topology
    sheet.Range("A" + str(currow) + ":B" + str(currow)).Merge()
    sheet.Range("C" + str(currow) + ":H" + str(currow)).Merge()
    Topology = xml_tree.xpath('.//General/GeneralData[@Name="Topology"]/@ID')[0]
    sheet.Cells(currow, 1).Value = "Топология"
    sheet.Cells(currow, 3).Value = Topology
    sheet.Cells(currow, 1).HorizontalAlignment = win32.constants.xlLeft
    currow = currow + 1
    # Self version (crc32)
    sheet.Range("A" + str(currow) + ":B" + str(currow)).Merge()
    sheet.Range("C" + str(currow) + ":H" + str(currow)).Merge()
    Topology = xml_tree.xpath('.//General/GeneralData[@Name="Topology"]/@ID')[0]
    sheet.Cells(currow, 1).Value = "Версия SiemensPie"
    sheet.Cells(currow, 3).Value = SelfCRC32()
    sheet.Cells(currow, 1).HorizontalAlignment = win32.constants.xlLeft
    currow = currow + 1

    # main work
    FunctionGroups = xml_tree.findall('Settings/FunctionGroup')
    for FunctionGroup in FunctionGroups:
        ProcessFunctionGroup(FunctionGroup)

    return


'''
 insert header and page stylization 
'''
def PageSetup():

    # page margins
    #sheet.PageSetup.LeftMargin = excel.InchesToPoints(0.393700787401575)
    #sheet.PageSetup.RightMargin = excel.InchesToPoints(0.393700787401575)
    #sheet.PageSetup.TopMargin = excel.InchesToPoints(0.393700787401575)
    #sheet.PageSetup.BottomMargin = excel.InchesToPoints(0.393700787401575)
    #sheet.PageSetup.HeaderMargin = excel.InchesToPoints(0.118110236220472)
    #sheet.PageSetup.FooterMargin = excel.InchesToPoints(0.118110236220472)
    #sheet.PageSetup.FitToPagesWide = 1
    #excel.ActiveWindow.Zoom = 90

    # text alignment
    sheet.Columns("A:H").VerticalAlignment = win32.constants.xlCenter
    sheet.Columns("A:H").WrapText = True
    sheet.Columns("A:H").NumberFormat = "@"

    # column width
    sheet.Columns("A:A").ColumnWidth = 6                                # address
    sheet.Columns("A:A").HorizontalAlignment = win32.constants.xlCenter
    sheet.Columns("B:B").ColumnWidth = 23                               # name
    sheet.Columns("C:C").ColumnWidth = 30                               # range
    sheet.Columns("D:G").ColumnWidth = 8.3                              # values in groups
    sheet.Columns("D:G").HorizontalAlignment = win32.constants.xlCenter
    sheet.Columns("H:H").ColumnWidth = 41.6                             # description

    # page margins
    sheet.PageSetup.LeftMargin = excel.InchesToPoints(0.433070866141732)
    sheet.PageSetup.RightMargin = excel.InchesToPoints(0.433070866141732)
    sheet.PageSetup.TopMargin = excel.InchesToPoints(0.393700787401575)
    sheet.PageSetup.BottomMargin = excel.InchesToPoints(0.590551181102362)
    sheet.PageSetup.HeaderMargin = excel.InchesToPoints(0.196850393700787)
    sheet.PageSetup.FooterMargin = excel.InchesToPoints(0.118110236220472)
    sheet.PageSetup.FitToPagesWide = 1
    sheet.PageSetup.Orientation = win32.constants.xlLandscape
    excel.ActiveWindow.Zoom = 90

    # footer
    sheet.PageSetup.RightFooter = "&F\n&P"

    return


'''
 insert chapter header
'''
def PrintHeader(text):

    global currow, last_title

    # check text to titles_correct
    text = config['titles_correct'].get(text, text)

    if (currow > 1) and (sheet.Cells(currow, 2).Text != ""):
        currow = currow +1
    sheet.Range("A"+str(currow)+":H"+str(currow)).Merge()
    sheet.Cells(currow,1).Value = text
    sheet.Cells(currow,1).HorizontalAlignment = win32.constants.xlLeft
    sheet.Cells(currow,1).Font.Bold = True
    sheet.Cells(currow,1).Font.Size = 13
    sheet.Cells(currow,1).Interior.Pattern = win32.constants.xlSolid
    sheet.Cells(currow,1).Interior.ThemeColor = win32.constants.xlThemeColorDark1
    sheet.Cells(currow,1).Interior.TintAndShade = -0.149998474074526
    currow = currow +1
    last_title = text

    return

'''
 insert chapter sub-header
'''
def PrintSubHeader(text):

    if text != "":

        global currow

        # check text to titles_correct
        text = config['titles_correct'].get(last_title+"|"+text, text)

        if (currow > 1) and (sheet.Cells(currow, 2).Text != ""):
            currow = currow +1
        sheet.Range("A"+str(currow)+":H"+str(currow)).Merge()
        sheet.Cells(currow,1).Value = text
        sheet.Cells(currow,1).HorizontalAlignment = win32.constants.xlLeft
        sheet.Cells(currow,1).Font.Bold = True
        sheet.Cells(currow,1).Font.Size = 11
        sheet.Cells(currow,1).Interior.Pattern = win32.constants.xlSolid
        sheet.Cells(currow,1).Interior.ThemeColor = win32.constants.xlThemeColorDark1
        sheet.Cells(currow,1).Interior.TintAndShade = -0.149998474074526
        currow = currow +1

        return

'''
 insert header with column titles
'''
def PrintColumnHeader():

    global currow

    sheet.Cells(currow,1).Value = "№\r\nАдрес"
    sheet.Cells(currow,2).Value = "Наименование уставки"
    sheet.Cells(currow,3).Value = "Диапазон уставок"
    sheet.Cells(currow,4).Value = "Заданная уставка"
    sheet.Cells(currow,8).Value = "Описание"
    sheet.Range("D"+str(currow)+":G"+str(currow)).Merge()

    sheet.Range("A" + str(currow) + ":H" + str(currow)).HorizontalAlignment = win32.constants.xlCenter
    sheet.Range("A" + str(currow) + ":H" + str(currow)).Interior.Pattern = win32.constants.xlSolid
    sheet.Range("A" + str(currow) + ":H" + str(currow)).Interior.ThemeColor = win32.constants.xlThemeColorDark1
    sheet.Range("A" + str(currow) + ":H" + str(currow)).Interior.TintAndShade = -0.149998474074526
    sheet.Range("A" + str(currow) + ":H" + str(currow)).Font.Bold = True
    sheet.Rows(currow).RowHeight = 32
    currow = currow + 1

    return currow - 1


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

    global currow
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

    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).HorizontalAlignment = win32.constants.xlCenter
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.Pattern = win32.constants.xlSolid
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.ThemeColor = win32.constants.xlThemeColorDark1
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Interior.TintAndShade = -0.149998474074526
    sheet.Range("A" + str(RowNo - 1) + ":H" + str(RowNo)).Font.Bold = True

    currow = currow + 1

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

    # correct some parameter names by config
    return config['params_correct'].get(Address, ParameterName)


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
        if (Dimension=="А"):
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

    ParameterValue  = Parameter.find(r"Value")                 # !!!
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

    global currow

    sheet.Cells(currow, 1).Value = ParameterData['Address']
    sheet.Cells(currow, 2).Value = ParameterData['Name']
    sheet.Cells(currow, 3).Value = ParameterData['Range']

    ParameterValues = ParameterData['Values']

    # if values are equal merge cells
    if (ParameterValues[0] == ParameterValues[1] == ParameterValues[2] == ParameterValues[3]):
        sheet.Range("D" + str(currow) + ":G" + str(currow)).Merge()
        sheet.Cells(currow, 4).Value = ParameterValues[0]
    elif (ParameterValues[0] == ParameterValues[1] == ParameterValues[2]):
        sheet.Range("D" + str(currow) + ":F" + str(currow)).Merge()
        sheet.Cells(currow, 4).Value = ParameterValues[0]
        sheet.Cells(currow, 7).Value = ParameterValues[3]
    elif (ParameterValues[0] == ParameterValues[1]):
        sheet.Range("D" + str(currow) + ":E" + str(currow)).Merge()
        sheet.Cells(currow, 4).Value = ParameterValues[0]
        if (ParameterValues[2] == ParameterValues[3]):
            sheet.Range("F" + str(currow) + ":G" + str(currow)).Merge()
            sheet.Cells(currow, 6).Value = ParameterValues[2]
        else:
            sheet.Cells(currow, 6).Value = ParameterValues[2]
            sheet.Cells(currow, 7).Value = ParameterValues[3]
    elif (ParameterValues[1] == ParameterValues[2] == ParameterValues[3]):
        sheet.Cells(currow, 4).Value = ParameterValues[0]
        sheet.Range("E" + str(currow) + ":G" + str(currow)).Merge()
        sheet.Cells(currow, 5).Value = ParameterValues[1]
    elif (ParameterValues[2] == ParameterValues[3]):
        sheet.Cells(currow, 4).Value = ParameterValues[0]
        sheet.Cells(currow, 5).Value = ParameterValues[1]
        sheet.Range("F" + str(currow) + ":G" + str(currow)).Merge()
        sheet.Cells(currow, 6).Value = ParameterValues[2]
    elif (ParameterValues[1] == ParameterValues[2]):
        sheet.Cells(currow, 4).Value = ParameterValues[0]
        sheet.Range("E" + str(currow) + ":F" + str(currow)).Merge()
        sheet.Cells(currow, 5).Value = ParameterValues[1]
        sheet.Cells(currow, 7).Value = ParameterValues[3]
    else:
        sheet.Cells(currow, 4).Value = ParameterValues[0]
        sheet.Cells(currow, 5).Value = ParameterValues[1]
        sheet.Cells(currow, 6).Value = ParameterValues[2]
        sheet.Cells(currow, 7).Value = ParameterValues[3]

    sheet.Cells(currow, 8).Value = ParameterData['Description']
    currow = currow + 1

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

    ParameterType = Parameter.attrib['Type']

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

    global group_has_elec_values, group_has_group_values

    SettingPageName = SettingPage.attrib['Name']
    PrintSubHeader(SettingPageName)
    header_row = PrintColumnHeader()
    Parameters = SettingPage.findall("Parameter")

    group_has_elec_values = False
    group_has_group_values = False

    for Parameter in Parameters:
        ProcessParameter(Parameter)
    if group_has_elec_values:
        if primary:
            UpdateColumnHeader(header_row, "вторичные величины", "первичные величины")
        else:
            UpdateColumnHeader(header_row, "вторичные величины", "вторичные величины")
    if group_has_group_values:
        InsertGroupHeader(header_row+1)

    return


'''
 process function group
'''
def ProcessFunctionGroup(FunctionGroup):

    FunctionGroupName = FunctionGroup.attrib['Name']
    PrintHeader(FunctionGroupName)
    SettingPages = FunctionGroup.findall("SettingPage")
    for SettingPage in SettingPages:
        ProcessSettingPage(SettingPage)

    return

'''
 set ptintable area and other print settings
'''
def PrintSetup():

    excel.PrintCommunication = False
    sheet.PageSetup.PrintArea = "$A$1:$H$" + str(currow)
    sheet.PageSetup.Zoom = False
    sheet.PageSetup.FitToPagesWide = 1
    sheet.PageSetup.FitToPagesTall = 0
    excel.PrintCommunication = True

    return


ProcessCommandLine()
ProcessAll()
PrintSetup()
