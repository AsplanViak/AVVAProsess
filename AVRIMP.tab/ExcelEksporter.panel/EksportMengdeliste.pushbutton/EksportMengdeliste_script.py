# -*- coding: utf-8 -*-
#  Headeren over må du ha om scriptet inneholder æøå.

"""
Du trenger:
IronPython innstallert på din PC.
pyCharm(python editor).
Fungerende autocomplete i pycharm.
    Under pyCharm settings:
    Project Interpreter:
    Sett opp pycharm med IronPython som din project enterpreter(ikke virtuel)
    Sett disse mappene til Source
    Project Structure:
    Sett disse mappene til sources: pyrevitlib, site-packages, dinlokalesharepoint\Asplan Viak\AVTools - Dokumenter\AVToolsRammeverk\AutocompleteStubs

"""

# Start MÅ ha
__title__ = 'Eksport mengdeliste'  # Denne overstyrer navnet på scriptfilen
__author__ = 'Asplan Viak'  # Dette kommer opp som navnet på utvikler av knappen
__doc__ = "Klikk for å eksportere en mengdeliste rør og rørdeler til excel."  # Dette blir hjelp teksten som kommer opp når man holder musepekeren over knappen.
# End MÅ ha

# Kan sløyfes
__cleanengine__ = True  # Dette forteller tolkeren at den skal sette opp en ny ironpython motor for denne knappen, slik at den ikke kommer i konflikt med andre funksjoner, settes nesten alltid til FALSE, TRUE når du jobber med knappen.
__fullframeengine__ = False  # Denne er nødvendig for å få tilgang til noen moduler, denne gjør knappen vesentrlig tregere i oppstart hvis den står som TRUE
# __context__ = "zerodoc"  # Denne forteller tolkeren at knappen skal kunne brukes selv om et Revit dokument ikke er åpent.
# __helpurl__ = "google.no"  # Hjelp URL når bruker trykker F1 over knapp.
__min_revit_ver__ = 2015  # knapp deaktivert hvis klient bruker lavere versjon
__max_revit_ver__ = 2032  # Skjønner?
__beta__ = False  # Knapp deaktivert hos brukere som ikke har spesifikt aktivert betaknapper

# Finn flere variabler her:
# https://pyrevit.readthedocs.io/en/latest/articles/scriptanatomy.html

import math

import clr

from pyrevit import HOST_APP
doc = HOST_APP.doc
uidoc = HOST_APP.uidoc
app = HOST_APP.app
#app = uidoc.Application
#uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

#doc = DocumentManager.Instance.CurrentDBDocument
#uiapp = DocumentManager.Instance.CurrentUIApplication
#uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
#app = uiapp.Application


clr.AddReference("RevitNodes")

from Autodesk.Revit import UI, DB

from Autodesk.Revit.DB import *

from Autodesk.Revit.DB import Plumbing, IFamilyLoadOptions

import os


clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

from Microsoft.Office.Interop import Excel
from Microsoft.Office.Interop.Excel import XlListObjectSourceType, Worksheet, Range, XlYesNoGuess

from System.Collections.Generic import List
from System.Collections.Generic import *
from System import Guid
from System import Array

import datetime

#import sys

#sys.path.append("C:\Program Files (x86)\IronPython 2.7\Lib")

xl = Excel.ApplicationClass()
xl.Visible = True
xl.DisplayAlerts = True

def SaveListToExcel(filePath, exportData):
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets[1]
        #ws.title = 'komp'
        rows = len(exportData)
        #print(rows)
        cols = max(len(i) for i in exportData)
        a = Array.CreateInstance(object, rows, cols) #row and column
        for r in range(rows):
            for c in range (cols):
            #for c in range (4):
                try:
                    a[r,c] = exportData[r][c]
                except:
                    a[r,c] = ''
            
        xlrange = ws.Range["A1", chr(ord('@')+cols) + str(rows)]
        xlrange.Value2 = a
        for r in range(rows):
            #if ws.Cells[r,5].Value == 1 or r == 1:
            if exportData[r][4] == 1 or r == 0:
                bold_range = ws.Range[ws.Cells[r+1, 1], ws.Cells[r+1, 4]]  # Columns A to D
                bold_range.Cells.Font.Bold = True        
        ws.Columns[5].Delete()
        wb.SaveAs(filePath)
        return True
    except:
        return False

#button = UI.TaskDialogCommonButtons.None
#result = UI.TaskDialogResult.Ok
#UI.TaskDialog.Show('Connector elementer lagt til og oppdater', report_tekst, button)

def convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(value_in_feet):
    DN = (value_in_feet * 304.8 * 2)
    nearest_DN = round(DN/5)*5
    if abs(DN - nearest_DN) <= 1 and DN>35:
        DN = nearest_DN
    return 'DN' + str(int(DN))

def keyn(k):
    nums = set(list("0123456789"))
    chars = set(list(k))
    chars = chars - nums
    for i in range(len(k)):
        for c in chars:
            k = k.replace(c + "0", c)
    l = list(k)
    base = 10
    j = 0
    for i in range(len(l) - 1, -1, -1):
        try:
            l[i] = int(l[i]) * base ** j
            j += 1
        except:
            j = 0
    l = tuple(l)
    # print l
    return l

# Place your code below this line

PA = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()
PI = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
PF = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()

main_list = []

PA_systemtype = []
PA_family = []
PA_DN = []
PI_systemtype = []
PI_DN = []
names = []

for i in PA:
    # Finn family
    try:
        family = i.Symbol.FamilyName
        # Sjekk om flens
        if 'øsflens' in family or 'veiseflens' in family:
            # Finn system type
            try:
                # systemtype = (i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()).ToDSType(True).Name
                systemtype = (i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString())
            except:
                systemtype = 'udefinert system type PA'
            # Finn DN
            try:
                connectors = i.MEPModel.ConnectorManager.Connectors
                # DN = int(connectors[0]).Radius*2
                for kk in connectors:
                    #DN = 'DN' + str(int(kk.Radius * 304.8 * 2))
                    DN = convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(kk.Radius)
                    break
            except:
                try:
                    connectors = i.ConnectorManager.Connectors
                    for kk in connectors:
                        #DN = 'DN' + str(int(kk.Radius * 304.8 * 2))
                        DN = convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(kk.Radius)
                        break
                except:
                    DN = 'udefinert DN PA'
            main_list.append(list([systemtype, family, DN, 1]))
        else:
            continue
    except:
        continue

# names.append(i.Name)    # returnerer "Standard", "DN150", eller "rustfritt rør"

for i in PI:
    # Finn system type
    try:
        # PI_systemtype.append(i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()).ToDSType(True).Name
        # systemtypeparam = (i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString())
        # systemtype = systemtypeparam.ToDSType(True).Name
        systemtype = (i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString())
    except:
        systemtype = 'udefinert system type PI'
    # Finn rørlengde
    try:
        lengde = (i.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()) * 304.8 / 1000
    except:
        lengde = 0
    # Finn family
    if 'ustfri' in i.Name:
        family = "Rette rørlengder rustfritt/syrefast"
    else:
        family = "Rette rørlengder " + i.Name
    # Finn DN
    try:
        DN = 'DN' + (i.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsValueString()[:-2])
    except:
        DN = 'udefinert DN PI'

    main_list.append(list([systemtype, family, DN, lengde]))

for i in PF:
    # Finn system type
    try:
        # PI_systemtype.append(i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()).ToDSType(True).Name
        # systemtypeparam = (i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString())
        # systemtype = systemtypeparam.ToDSType(True).Name
        systemtype = (i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString())
    except:
        systemtype = 'udefinert system type PF'
    # Finn family

    try:
        family = i.Symbol.FamilyName
        if 'magi_sewer_bend_NOR' in family:
            family = 'Bend'
        if 'magi_sewer_tee_wye_NOR' in family:
            family = 'T-rør'
        if 'magi_sewer_transition_NOR' in family:
            family = 'Overgang'
        if 'M_Pipe Straight Coupling' in family:
            continue
        if 'Bend ' in family:
            vinkel = i.LookupParameter('Angle').AsDouble() / math.pi * 180
            vinkel = int(vinkel)
            if vinkel > 85 and vinkel < 92:
                vinkel = 90
            if vinkel > 40 and vinkel < 47:
                vinkel = 45
            family = family + ' ' + str(vinkel) + '°'

    except:
        family = 'udefinert familie PF'

    # Finn DN
    try:

        connectors = i.MEPModel.ConnectorManager.Connectors
        DNs = []
        for kk in connectors:
            #DNs.append('DN' + str(int(kk.Radius * 304.8 * 2)))
            DNs.append(convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(kk.Radius))
        DN_list = list(set(DNs))
        DN = DN_list[0]
        if len(DN_list) == 2:
            DN = DN + '-' + DN_list[1]
        elif len(DN_list) > 2:
            DN = range(1, len(DN_list) - 1)
            for j in range(1, len(DN_list) - 1):
                DN = DN + '-' + DN_list[j]

    except:
        try:
            connectors = i.ConnectorManager.Connectors
            DN = []
            for kk in connectors:
                #DNs.append('DN' + str(int(kk.Radius * 304.8 * 2)))
                DNs.append(convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(kk.Radius))
            DN_list = list(set(DNs))
            DN = DN_list[0]
            if len(DN_list) > 1:
                for j in range(1, len(DN_list) - 1):
                    DN = DN + '-' + DN_list[j]
        except:
            DN = 'udefinert DN PF'

    main_list.append(list([systemtype, family, DN, 1]))

# sortering og komprimmering av list

main_list = sorted(main_list, key=lambda x: (x[0], x[1], x[2]))

main_list_compressed = []
main_list_compressed.append(main_list[0])  # første rad
for n in range(1, len(main_list)):
    if main_list[n][0:3] == main_list[n - 1][0:3]:
        b = len(main_list_compressed) - 1
        main_list_compressed[b][3] = main_list_compressed[b][3] + main_list[n][3]
    else:
        main_list_compressed.append(main_list[n])

for m in range(0, len(main_list_compressed)):
    if not isinstance(main_list_compressed[m][3], int):
        # rund av til en desimal
        main_list_compressed[m][3] = round(float(main_list_compressed[m][3]), 1)

# names.append(i.Name)	 # returnerer "Standard", "DN150", eller "rustfritt rør"


# legg til kolonne for postnr
main_list_compressed = [x + ['x9'] for x in main_list_compressed]

# legg til postnr i denne kolonne
# if b1:
#	if b1[0] != 'no_table_found':

PS = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsElementType().ToElements()
b1 = []
for s in PS:

    ps_name = Element.Name.GetValue(s)
    try:
        if s.LookupParameter('Gprog_postnr').AsString():
            ps_postnr = s.LookupParameter('Gprog_postnr').AsString()
        else:
            if s.LookupParameter('Abbreviation').AsString():
                ps_postnr = s.LookupParameter('Abbreviation').AsString()
            else:
                ps_postnr = ''
    except:
        if s.LookupParameter('Abbreviation').AsString():
            ps_postnr = s.LookupParameter('Abbreviation').AsString()
        else:
            ps_postnr = ''

    b1.append(list([ps_name, ps_postnr]))

for g in range(len(main_list_compressed)):
    for h in range(len(b1)):
        if b1[h][0] == main_list_compressed[g][0]:
            main_list_compressed[g][4] = b1[h][1]

# sorter ut fra postnr
a1 = sorted(main_list_compressed, key=lambda x: keyn(x[4]))

a2 = []
a2.append(['System type', 'Komponent', 'Dimensjon', 'Lengde (m) / \n Antall (stk)', 'sort'])

for i in range(len(a1)):
    if i == 0 or a1[i][0] != a1[i - 1][0]:
        a2.append(['', '', '', '', ''])
        p = ''
        for k in range(len(b1)):
            if b1[k][0] == a1[i][0]:
                p = b1[k][1]
        a2.append([a1[i][0], 'Post: ' + p, '', '', 1])

    a2.append([a1[i][0], a1[i][1], a1[i][2], a1[i][3], ''])

# Assign your output to the OUT variable.
#OUT = a2

#Finne sti til mine dokumenter
mydoc = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents')

x = datetime.datetime.now()
timestamp = x.strftime("%Y_%m_%d %H_%M")
#timestamp = x.strftime("%Y %H-%M")
#debug.append(timestamp)

filename = mydoc+'\mengdeliste_'+timestamp+'.xlsx'

SaveListToExcel(filename, a2)

#xl.DisplayAlerts = True
