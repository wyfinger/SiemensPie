#!/usr/bin/python
# -*- coding: utf-8 -*-
__author__ = 'prim_miv'

#TODO: уставка 7137 в 7SJ не выводится в Excel, хотя в конфиге, если посмотреть в Digsi, есть.

import re
from string import ascii_lowercase
import xml.etree.ElementTree as ET
import win32com.client as win32
from lxml import etree
import sys
import argparse
import os
from win32com.client import gencache
if gencache.is_readonly:
    gencache.is_readonly = False
    gencache.Rebuild() #create gen_py folder if needed

xriofile = ""
xmlfile  = ""

# параметры преобразования
convert_to_primary = True # преобразование диапазонов или значений во вторичные величины
                          # если сделать True - скрипт спросит Ктт и Ктн

global group_has_elec_values   # признак того, что среди группы параметров есть хотябы одна электрическая величина
group_has_elec_values = False  # чтобы добавить подпись (первичные/вторичные величины) в заголовок таблицы

global group_has_group_values   # признак того, что среди группы параметров есть хотябы одна имеющая группы уставок
group_has_group_values = False  # чтобы добавить подпись строку с наименованиями групп в шапку таблицы


parser = argparse.ArgumentParser()
parser.add_argument('-p', '--primary', nargs='?', default='false', help='Конвертировать значение по вторичные')
parser.add_argument('file1', nargs='?')
parser.add_argument('file2', nargs='?')
namespace = parser.parse_args()

if (namespace.primary != None):
    if (namespace.primary.lower() == 'true'):
        convert_to_primary = True

# ищем файлы в коммандной строки
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
    xriofile = os.path.splitext(xmlfile)[0]+'.xrio'
    if (os.path.isfile(xriofile) == False):
        xriofile = ""
if (xriofile != "") & (xmlfile == ""):
    xmlfile = os.path.splitext(xriofile)[0]+'.xml'
    if (os.path.isfile(xmlfile) == False):
        xmlfile = ""

if (xmlfile == "") | (xriofile == ""):
    print("Error. XML or XRio file is not exists.")
    print("Use: sp.exe [-p] [xml or xrio file] [xml or xrio file]")
    print("  -p  - convert electrical values to primary")
    print("  set one (xml or xrio) file if they have the same name")
    sys.exit()

print("XRio File: " + xriofile)
print("XML File: " + xmlfile)

# создаем файл excel
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = True
#excel.DisplayAlerts = False

wb = excel.Workbooks.Add()
sheet = wb.Worksheets.Add()

# читаем xrio и xml
xriod = etree.parse(xriofile)
xmld = ET.parse(xmlfile)

# ktt и ktn попытаемя автоматически определить в PrintSpec()
ktt = 1
ktn = 1

# общие переменные
currow = 1


# вставка шапки и оформление столбцов
def PageSetup():
    # поля
    sheet.PageSetup.LeftMargin = excel.InchesToPoints(0.393700787401575)
    sheet.PageSetup.RightMargin = excel.InchesToPoints(0.393700787401575)
    sheet.PageSetup.TopMargin = excel.InchesToPoints(0.393700787401575)
    sheet.PageSetup.BottomMargin = excel.InchesToPoints(0.393700787401575)
    sheet.PageSetup.HeaderMargin = excel.InchesToPoints(0.118110236220472)
    sheet.PageSetup.FooterMargin = excel.InchesToPoints(0.118110236220472)
    sheet.PageSetup.FitToPagesWide = 1
    excel.ActiveWindow.Zoom = 90

    # все содержимое ячеек выравнивается по высоте посередине и текст переносится
    sheet.Columns("A:H").VerticalAlignment = win32.constants.xlCenter
    sheet.Columns("A:H").WrapText = True
    sheet.Columns("A:H").NumberFormat = "@"

    # ширины столбцов
    sheet.Columns("A:A").ColumnWidth = 6      # номер уставки
    sheet.Columns("A:A").HorizontalAlignment = win32.constants.xlCenter
    sheet.Columns("B:B").ColumnWidth = 23     # наименование уставки
    sheet.Columns("C:C").ColumnWidth = 30     # диапазон возможных значений
    sheet.Columns("D:G").ColumnWidth = 8.3    # значения по группам уставок
    sheet.Columns("D:G").HorizontalAlignment = win32.constants.xlCenter
    sheet.Columns("H:H").ColumnWidth = 41.6   # описание

    return True

# вставить заголовок раздела
def PrintHeader(text):
    global currow
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

    return True

# вставить заголовок подраздела
def PrintSubHeader(text):

    if text != "":
        global currow
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

    pass

# вставить шапку с наименованиями столбцов
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
    currow=currow + 1

    return currow - 1

# обновление заголовкой подгруппы, если в группе есть электрические параметры (первичные или вторичные)
def UpdateColumnHeader(RowNo, addtext_range="", addtext_value=""):

    if addtext_range != "":
        sheet.Cells(RowNo,3).Value = sheet.Cells(RowNo,3).Value + "\r\n("+addtext_range+")"
    if addtext_value != "":
        sheet.Cells(RowNo,4).Value = sheet.Cells(RowNo,4).Value + "\r\n("+addtext_value+")"
    sheet.Rows(RowNo).RowHeight = 32
    return

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

# получение короткого наименования уставки из XRio (в XML файле этих данных нет)
def ExtractParameterName(Address, XRio):

    ParameterName = XRio.xpath("//ForeignId[text()='"+Address+"']/parent::*/Name/text()")
    if (ParameterName!=None) and (len(ParameterName) > 0):
        ParameterName=str(ParameterName[0])
    else:
        ParameterName= ""

    return ParameterName

# получение количества знаков после запятой для величины из XRio (в XML файле этих данных нет)
def ExtractParameterPrecision(Address, XRio):

    ParameterPrecision = XRio.xpath("//ForeignId[text()='"+Address+"']/parent::*/Unit")
    if (ParameterPrecision!=None) and (len(ParameterPrecision) > 0):
        return int(ParameterPrecision[0].attrib['DecimalPlaces'])
    else:
        return 0

# преобразование вторичной величины в первичную
def ConvertToPrimary(Address, Value, Dimension, SecondaryPrecision):

    # адреса, которые задают Ктт и Ктн не нужно ни во что переводить
    spec_addresses = [
        '0203',  # Первичное номинальное напряжение, 7SA
        '0204',  # Вторичное номинальное напряжение, 7SA
        '0205',  # Первичный номинальный ток ТТ, 7SA
        '0206',  # Вторичный номинальный ток ТТ, 7SA
        '0217',  # Номин. ток IЕ ТТ, первичный, 7SJ
        '1101',  # Измерения: 100% шкалы напряжения, 7SJ
        '1102',  # Измерения: 100% шкалы тока, 7SJ
        '1103',  # Измерения: 100% шкалы напряжения, 7SA
        '1104'   # Измерения: 100% шкалы тока, 7SA
    ]

    if (Address in spec_addresses):
        return "%g" % float(Value) + " " + Dimension

    global group_has_elec_values

    # это число?
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

# получение значений параметра во всех группах уставок (массив)
def ExtractParameterValues(Parameter):

    ParameterAddr = Parameter.attrib['DAdr']
    ParameterType = Parameter.attrib['Type']
    #ParameterValue = Parameter.find(r"Value[not(@SettingGroup)]").text   # эта долбанная ElementTree не поддерживает
                                                                          # not() синтаксис XPath
    global group_has_group_values

    ParameterValue  = Parameter.find(r"Value")                 # !!!
    ParameterValueA = Parameter.find(r"Value[@SettingGroup='A']")
    ParameterValueB = Parameter.find(r"Value[@SettingGroup='B']")
    ParameterValueC = Parameter.find(r"Value[@SettingGroup='C']")
    ParameterValueD = Parameter.find(r"Value[@SettingGroup='D']")

    if ParameterValueA==None:
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

        # преобразование едениц измерения, если нужно
        if convert_to_primary:
            SecondaryPrecision = ExtractParameterPrecision(ParameterAddr, xriod)
            ParameterValueA = ConvertToPrimary(ParameterAddr, ParameterValueA, Dimension, SecondaryPrecision)
            ParameterValueB = ConvertToPrimary(ParameterAddr, ParameterValueB, Dimension, SecondaryPrecision)
            ParameterValueC = ConvertToPrimary(ParameterAddr, ParameterValueC, Dimension, SecondaryPrecision)
            ParameterValueD = ConvertToPrimary(ParameterAddr, ParameterValueD, Dimension, SecondaryPrecision)
        else:
            # если значение = oo - не отображаем еденицу измерения
            ParameterValueA = ParameterValueA if ParameterValueA == "oo" else ParameterValueA + " " + Dimension
            ParameterValueB = ParameterValueB if ParameterValueB == "oo" else ParameterValueB + " " + Dimension
            ParameterValueC = ParameterValueC if ParameterValueC == "oo" else ParameterValueC + " " + Dimension
            ParameterValueD = ParameterValueD if ParameterValueD == "oo" else ParameterValueD + " " + Dimension

    return [ParameterValueA, ParameterValueB, ParameterValueC, ParameterValueD]

# получение диапазона возможных значений параметра
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

# обработка одного параметра
def ProcessParameter(Parameter):

    ParameterAddress = Parameter.attrib['DAdr']

    print(ParameterAddress)

    ParameterName = ExtractParameterName(ParameterAddress, xriod)
    ParameterDescription = Parameter.attrib['Name']

    ParameterType = Parameter.attrib['Type']

    Range = ExtractParameterRange(Parameter)
    ParameterRange = Range[0]
    ParameterValues = ExtractParameterValues(Parameter)

    global currow
    sheet.Cells(currow,1).Value = ParameterAddress
    sheet.Cells(currow,2).Value = ParameterName
    sheet.Cells(currow,3).Value = ParameterRange

    #sheet.Cells(currow,4).Value = ParameterValues[0]
    #sheet.Cells(currow,5).Value = ParameterValues[1]
    #sheet.Cells(currow,6).Value = ParameterValues[2]
    #sheet.Cells(currow,7).Value = ParameterValues[3]

    # объединяем уставки, одинаковые в соседних группах
    if (ParameterValues[0]==ParameterValues[1]==ParameterValues[2]==ParameterValues[3]):
            sheet.Range("D"+str(currow)+":G"+str(currow)).Merge()
            sheet.Cells(currow,4).Value = ParameterValues[0]
    elif (ParameterValues[0]==ParameterValues[1]==ParameterValues[2]):
            sheet.Range("D"+str(currow)+":F"+str(currow)).Merge()
            sheet.Cells(currow,4).Value = ParameterValues[0]
            sheet.Cells(currow,7).Value = ParameterValues[3]
    elif (ParameterValues[0]==ParameterValues[1]):
        sheet.Range("D" + str(currow) + ":E" + str(currow)).Merge()
        sheet.Cells(currow, 4).Value = ParameterValues[0]
        if (ParameterValues[2]==ParameterValues[3]):
            sheet.Range("F" + str(currow) + ":G" + str(currow)).Merge()
            sheet.Cells(currow, 6).Value = ParameterValues[2]
        else:
            sheet.Cells(currow,6).Value = ParameterValues[2]
            sheet.Cells(currow,7).Value = ParameterValues[3]
    elif (ParameterValues[1]==ParameterValues[2]==ParameterValues[3]):
        sheet.Cells(currow,4).Value = ParameterValues[0]
        sheet.Range("E" + str(currow) + ":G" + str(currow)).Merge()
        sheet.Cells(currow,5).Value = ParameterValues[1]
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

    sheet.Cells(currow,8).Value = ParameterDescription

    currow = currow +1

    pass

# обработка подгруппы параметров
def ProcessSettingPage(SettingPage):

    SettingPageName = SettingPage.attrib['Name']
    PrintSubHeader(SettingPageName)
    header_row = PrintColumnHeader()
    Parameters = SettingPage.findall("Parameter")

    global group_has_elec_values
    group_has_elec_values = False

    global group_has_group_values
    group_has_group_values = False

    for Parameter in Parameters:
        ProcessParameter(Parameter)
    if group_has_elec_values:
        if convert_to_primary:
            UpdateColumnHeader(header_row, "вторичные величины", "первичные величины")
        else:
            UpdateColumnHeader(header_row, "вторичные величины", "вторичные величины")
    if group_has_group_values:
        InsertGroupHeader(header_row+1)

    pass

# обработка функциональной группы параметров
def ProcessFunctionGroup(FunctionGroup):

    FunctionGroupName = FunctionGroup.attrib['Name']
    header_row = PrintHeader(FunctionGroupName)
    SettingPages = FunctionGroup.findall("SettingPage")
    for SettingPage in SettingPages:
        ProcessSettingPage(SettingPage)

    pass

# вставка в выходной документ служебной информации по устройству
def PrintSpec():

    global currow
    global ktt
    global ktn

    # MLFBDIGSI
    sheet.Range("A" + str(currow) + ":H" + str(currow)).Merge()
    MLFBDIGSI = xmld.find('.//General/GeneralData[@Name="MLFBDIGSI"]')
    MLFBDIGSI = MLFBDIGSI.attrib.get('ID')
    sheet.Cells(currow, 1).Value = MLFBDIGSI
    sheet.Cells(currow, 1).HorizontalAlignment = win32.constants.xlLeft
    currow = currow + 1

    # Version
    sheet.Range("A" + str(currow) + ":H" + str(currow)).Merge()
    Version = xmld.find('.//General/GeneralData[@Name="Version"]')
    Version = Version.attrib.get('ID')
    sheet.Cells(currow, 1).Value = Version
    sheet.Cells(currow, 1).HorizontalAlignment = win32.constants.xlLeft
    currow = currow + 1

    # Topology
    sheet.Range("A" + str(currow) + ":H" + str(currow)).Merge()
    Topology = xmld.find('.//General/GeneralData[@Name="Topology"]')
    Topology = Topology.attrib.get('ID')
    sheet.Cells(currow, 1).Value = Topology
    sheet.Cells(currow, 1).HorizontalAlignment = win32.constants.xlLeft
    currow = currow + 1

    if ((MLFBDIGSI[0:6]=="7SA522") or (MLFBDIGSI[0:6]=="7SD522")):
        P0203 = float(xmld.find('.//FunctionGroup/SettingPage/Parameter[@DAdr="0203"]/Value').text) # Первичное номинальное напряжение
        P0204 = float(xmld.find('.//FunctionGroup/SettingPage/Parameter[@DAdr="0204"]/Value').text) # Вторичное номинальное напряжение
        ktn = round((P0203 * 1000) / P0204)
        P0205 = float(xmld.find('.//FunctionGroup/SettingPage/Parameter[@DAdr="0205"]/Value').text)  # Первичный номинальный ток ТТ
        P0206 = xmld.find('.//FunctionGroup/SettingPage/Parameter[@DAdr="0206"]/Value').text  # Вторичный номинальный ток ТТ
        P0206 = xmld.find('.//FunctionGroup/SettingPage/Parameter[@DAdr="0206"]/Comment[@Number="'+P0206+'"]')
        P0206 = int(P0206.attrib['Name'][0:1])
        ktt = round(P0205 / P0206)
        print("Ktt= " + str(ktt))
        print("Ktn= " + str(ktn))
    #elif (MLFBDIGSI[0:6]=="7SD522"):
        #
    else:
        if (convert_to_primary == True):
            ktt = float(input("Укажите Ктт: "))
            ktn = float(input("Укажите Ктн: "))

    pass

# оформление листа
PageSetup()
# спецификация устройства
PrintSpec()
# разбор XML
FunctionGroups = xmld.findall('Settings/FunctionGroup')
for FunctionGroup in FunctionGroups:
    ProcessFunctionGroup(FunctionGroup)

#excel.Visible = True
#excel.DisplayAlerts = True


# block = root.find("CUSTOM/Block[@Id='SETTINGS']/Block[@Id='DC']")
# PrintHeader(block.find("Description").text)
# PrintColumnHeader()
# parameters = block.findall("Parameter")
# for parameter in parameters:
#     sheet.Cells(currow,1).Value = parameter.attrib["Id"]
#     sheet.Cells(currow,2).Value = parameter.find("ForeignId").text
#     sheet.Cells(currow,3).Value = parameter.find("Name").text
#     sheet.Cells(currow,4).Value = ExtractRange(parameter)
#     sheet.Cells(currow,5).Value = ExtractValue(parameter)
#     sheet.Range("E"+str(currow)+":H"+str(currow)).Merge()
#     sheet.Cells(currow,5).HorizontalAlignment = win32.constants.xlCenter
#     sheet.Cells(currow,9).Value = parameter.find("Description").text
#     currow = currow +1
#
# # данные энергосистемы 1 (общие для всех групп уставок)
# block = root.find("CUSTOM/Block[@Id='SETTINGS']/Block[@Id='FG2']")
# PrintHeader(block.find("Description").text)
# # пройдемся по подразделам блока Данные энергосистемы 1
# subblocks = block.findall("Block")
# for subblock in subblocks:
#     PrintSubHeader(subblock.find("Description").text)
#     PrintColumnHeader()
#     parameters = subblock.findall("Parameter")
#     for parameter in parameters:
#         sheet.Cells(currow,1).Value = parameter.attrib["Id"]
#         sheet.Cells(currow,2).Value = parameter.find("ForeignId").text
#         sheet.Cells(currow,3).Value = parameter.find("Name").text
#         sheet.Cells(currow,4).Value = ExtractRange(parameter)
#         sheet.Cells(currow,5).Value = ExtractValue(parameter)
#         sheet.Range("E"+str(currow)+":H"+str(currow)).Merge()
#         sheet.Cells(currow,5).HorizontalAlignment = win32.constants.xlCenter
#         sheet.Cells(currow,9).Value = parameter.find("Description").text
#         currow = currow +1
#
# # Регистрация аварийных режимов
# block = root.find("CUSTOM/Block[@Id='SETTINGS']/Block[@Id='FG4']")
# PrintHeader(block.find("Description").text)
# # пройдемся по подразделам блока Данные энергосистемы 1
# subblocks = block.findall("Block")
# for subblock in subblocks:
#     #PrintSubHeader(subblock.find("Description").text)
#     #PrintColumnHeader()
#     parameters = subblock.findall("Parameter")
#     for parameter in parameters:
#         sheet.Cells(currow,1).Value = parameter.attrib["Id"]
#         sheet.Cells(currow,2).Value = parameter.find("ForeignId").text
#         sheet.Cells(currow,3).Value = parameter.find("Name").text
#         sheet.Cells(currow,4).Value = ExtractRange(parameter)
#         sheet.Cells(currow,5).Value = ExtractValue(parameter)
#         sheet.Range("E"+str(currow)+":H"+str(currow)).Merge()
#         sheet.Cells(currow,5).HorizontalAlignment = win32.constants.xlCenter
#         sheet.Cells(currow,9).Value = parameter.find("Description").text
#         currow = currow +1
#
# excel.DisplayAlerts = True

#wb.SaveAs(xlsxfile)
#excel.Application.Quit()

