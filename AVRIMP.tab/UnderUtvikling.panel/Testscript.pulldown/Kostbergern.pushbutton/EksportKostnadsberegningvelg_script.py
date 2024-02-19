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
__title__ = 'Eksport kostnadsberegning'  # Denne overstyrer navnet på scriptfilen
__author__ = 'Asplan Viak'  # Dette kommer opp som navnet på utvikler av knappen
__doc__ = "Klikk for å eksportere en kostnadsberegning for rør, rørdeler og ventiler til excel."  # Dette blir hjelp teksten som kommer opp når man holder musepekeren over knappen.
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
from Autodesk.Revit.UI.Selection import ObjectType

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

import csv


#import sys

#sys.path.append("C:\Program Files (x86)\IronPython 2.7\Lib")

melding = 'Dette scriptet er ikke kvalitetsikret, og bruk av scriptet er på eget ansvar.\n\nFaktorer for prisjustering er ikke offisielle AV-tall. Bruker er selv ansvarlig for å gjøre gode anslag på sine prosjekter.\n\nEnhetsprisene er i virkeligheten avhengige av en rekke forhold, som det ikke tas hensyn til i dette scriptet, som f.eks. veggtykkelse på rør, PN, krav til ventiler etc. Bruker er selv ansvarlig for å kontrollere om prisene som bruker i script vil være representative for sitt eget prosjekt.\n\nFor å skille mellom forskjellige entrepriser, og mellom eksisterende og nye rør, kan entreprise legges inn som parameter "Entreprise_piping_systems", eller parameter "Description" for filer uten den første parameteren tilgjengelig. Dette gjøres ved å endre "type properties" for "piping system" som er brukt på aktuelle rør.\n\nFor rette rørlengder vil kun rustfrie rør inkluderes, men enkelte rørdeler (bend, påstikk) av andre materialer vil desverre få samme priser som tilsvarende deler for rustfritt! Det anbefales derfor å modellere andre rør (PP, PVC, sveisemuffer, pressfittings etc) som egne piping systems og skille disse ut som egen entreprise (f.eks. "E61 plastrør"), som vist over, slik at disse kan prises manuelt.\n\nTips for kopiering over til andre regneark: Flytt først celler med prisjusteringsfaktor slik at disse ligger på samme celler i regneark du skal kopiere celle fra og regneark du skal kopiere til. Velg deretter funksjon "vis formler" og kopier over.'
button = UI.TaskDialogCommonButtons.None
result = UI.TaskDialogResult.Ok
UI.TaskDialog.Show('Viktig!',melding,button)

def SortedPoints(fittingspoints, ductStartPoint):
    sortedpoints = sorted(fittingspoints, key=lambda x: measure(ductStartPoint, x))
    return sortedpoints


def convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(value_in_feet):
    DN = (value_in_feet * 304.8 * 2)
    nearest_DN = round(DN/5)*5
    if abs(DN - nearest_DN) <= 1 and DN>35:
        DN = nearest_DN
    return 'DN' + str(int(DN))

xl = Excel.ApplicationClass()
xl.Visible = True
xl.DisplayAlerts = True

def SaveListToExcel(filePath, exportData):
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets[1]
        #ws.title = 'komp'
        rows = len(exportData)
        cols = max(len(i) for i in exportData)
        a = Array.CreateInstance(object, rows, cols) #row and column
        for r in range(rows):
            for c in range (cols):
                try:
                    a[r,c] = exportData[r][c]
                except:
                    a[r,c] = ''
        xlrange = ws.Range["A1", chr(ord('@')+cols) + str(rows)]
        xlrange.Value2 = a
        wb.SaveAs(filePath)
        return True
    except:
        return False


#button = UI.TaskDialogCommonButtons.None
#result = UI.TaskDialogResult.Ok
#UI.TaskDialog.Show('Connector elementer lagt til og oppdater', report_tekst, button)

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

def unicode_csv_reader(utf8_data, dialect=csv.excel):
    csv_reader = csv.reader(utf8_data, dialect=dialect, delimiter =";")
    for row in csv_reader:
        yield [unicode(cell, 'utf-8') for cell in row]


########################
##############Last inn prisbank fra excel

#prisbank = list(csv.reader(open("S:\Felles\_AVstandard\Revit\Dynamo\VA-prosess\Prisbank.csv"), delimiter  =";"))
prisbank = list(unicode_csv_reader(open("S:\Felles\_AVstandard\Revit\Dynamo\VA-prosess\Prisbank.csv")))
#prisbank_prosjekter = list(csv.reader(open("S:\Felles\_AVstandard\Revit\Dynamo\VA-prosess\Prisbank_prosjekter.csv"), delimiter  =";"))
prisbank_prosjekter = list(unicode_csv_reader(open("S:\Felles\_AVstandard\Revit\Dynamo\VA-prosess\Prisbank_prosjekter.csv")))


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



# make selection in UI for selecting pipe accessories and mech eq ++
cat_list = ['Pipe Accessories', 'Mechanical Equipment', 'Generic Model']
picked = []
try:
    picked = uidoc.Selection.PickObjects(ObjectType.Element)
except:
    pass
EQ = []
for k in picked:
    el = doc.GetElement(k.ElementId)
    if el.Category.Name in cat_list:
        EQ.append(el)



for i in PA:
    #Legger inn parameter for sortering. Default er PA1. I Praksis vil denne brukes mest til instrumentering. Mengdemålere vil ha PA0.
    #Ventiler får PA2. Flenser PA3.
    sortering = 'PA1'
    #sjekk om markert
    if not i.ElementID in EQ:
        continue
    # Finn family
    try:
        family = i.Symbol.FamilyName
        # Generaliser tekst
        if 'Krage-Løsflens' in family:
            family = 'Krage-løsflens'
            #Plassere disse sist ved sortering
            sortering = 'PA3'
        if 'Sveiseflens' in family:
            family = 'Sveiseflens'
            #Plassere disse sist ved sortering
            sortering = 'PA3'
        if 'Sluseventil' in family:
            family = 'Sluseventil'
        if 'Mengdemaaler' in family:
            family = 'Mengdemåler'
            #Plassere disse først ved sortering
            sortering = 'PA0'
        if 'Dreiespjeldsventil innspent manuell' in family:
            family = 'Dreiespjeldsventil innspent manuell'
        if  'ventil' in family:
            sortering = 'PA22'
        if 'uleventil' in family:
            sortering = 'PA21'
        if 'PF-stykke' in family:
            family = 'PF-stykke'

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
                #DN = 'DN' + str(int(math.ceil(kk.Radius * 304.8 * 2)))
                DN = convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(kk.Radius)
                break
        except:
            try:
                connectors = i.ConnectorManager.Connectors
                for kk in connectors:
                    #DN = 'DN' + str(int(math.ceil(kk.Radius * 304.8 * 2)))
                    DN = convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(kk.Radius)
                    break
            except:
                #DN = 'udefinert DN PA'
                DN = ''

        #Fjern DN fra familynavn dersom den er identisk med DN
        if family.startswith(DN) and DN <>'':
            family = family[len(DN)+1:]

        main_list.append(list([systemtype, family, DN, 1, sortering]))

    except:
        continue

# names.append(i.Name)    # returnerer "Standard", "DN150", eller "rustfritt rør"

for i in PI:
    #sjekk om markert
    if not i.ElementID in EQ:
        continue
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
        DN = 'DN' + (i.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsValueString()[:-3])
    except:
        DN = 'udefinert DN PI'

    main_list.append(list([systemtype, family, DN, lengde, 'PI']))

for i in PF:
    #sjekk om markert
    if not i.ElementID in EQ:
        continue
    #Resett parameter for å telle deler av bend.
    n_PF = 1
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
        if 'Bend ' in family:
            vinkel = i.LookupParameter('Angle').AsDouble() / math.pi * 180
            vinkel = int(vinkel)
            #if vinkel > 85 and vinkel < 92:
            #    vinkel = 90
            #if vinkel > 40 and vinkel < 47:
            #    vinkel = 45
            #family = family + ' ' + str(vinkel) + '°'
            if vinkel < 45:
                n_PF = 0.5
            else:
                n_PF = 1

    except:
        family = 'udefinert familie PF'

    # Finn DN
    try:

        connectors = i.MEPModel.ConnectorManager.Connectors
        DNs = []
        for kk in connectors:
            #DNs.append('DN' + str(int(math.ceil(kk.Radius * 304.8 * 2))))
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
                #DNs.append('DN' + str(int(math.ceil(kk.Radius * 304.8 * 2))))
                DNs.append(convert_radius_in_feet_to_DN_and_round_to_nearest_five_if_off_by_one(kk.Radius))
            DN_list = list(set(DNs))
            DN = DN_list[0]
            if len(DN_list) > 1:
                for j in range(1, len(DN_list) - 1):
                    DN = DN + '-' + DN_list[j]
        except:
            DN = 'udefinert DN PF'

    main_list.append(list([systemtype, family, DN, n_PF, 'PF']))

# legg til kolonne for entreprise
main_list = [x + ['Udefinert'] for x in main_list]

PS = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsElementType().ToElements()
entrepriseliste = []
for s in PS:

    ps_name = Element.Name.GetValue(s)
    try:
        if s.LookupParameter('Entreprise').AsString():
            entreprise = s.LookupParameter('Entreprise').AsString()
        else:
            if s.LookupParameter('Description').AsString():
                entreprise = s.LookupParameter('Description').AsString()
            else:
                entreprise = 'Udefinert'
    except:
        if s.LookupParameter('Description').AsString():
            entreprise = s.LookupParameter('Description').AsString()
        else:
            entreprise = 'Udefinert'

    entrepriseliste.append(list([ps_name, entreprise]))


#Legg til entreprise
for g in range(len(main_list)):
    for h in range(len(entrepriseliste)):
        if entrepriseliste[h][0] == main_list[g][0]:
            main_list[g][5] = entrepriseliste[h][1]
            break

#########################
# Sortering og komprimmering av list.
###################
# Sorterer etter entreprise, category, family, DN.
main_list = sorted(main_list, key=lambda x: (x[5], x[4], x[1], keyn(x[2])))
#a1 = sorted(main_list_compressed, key=lambda x: keyn(x[4]))

#Sletter systemtype-kolonne og setter entreprisekolonne først
for item in main_list:
	item[0], item[1], item[2], item[3], item[4] = item[5], item[1], item[2], item[3], item[4]



#print main_list[1]

#print main_list[2]

#Komprimerer liste
# slår sammen rader dersom entreprise, DN, family etc er likt.
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


#for item in main_list_compressed:
#    print item
# names.append(i.Name)	 # returnerer "Standard", "DN150", eller "rustfritt rør"
#print(main_list_compressed)

# sorter ut fra entreprise, family,
#a1 = sorted(main_list_compressed, key=lambda x: keyn(x[4]))
a1 = main_list_compressed

a2 = []             #hoveddataset til output. Med overskrifter etc.
a_entreprise = []   #subdataset entreprise
#a2.append(['System type', 'Komponent', 'Dimensjon', 'Lengde (m) / \n Antall (stk)', 'sort'])
#a_entreprise.append([a1[i][0], a1[i][1], a1[i][2], a1[i][3], ''])

entreprise_index = 0
entrepriser = [a1[0][5]]        #Legger til første entreprise i entrepriser-list

for i in range(len(a1)):
    #Sjekk ny entreprise
    #if i == 0 or a1[i][0] != a1[i - 1][0]:
    if i == 0 or a1[i][5] != a1[i - 1][5]:

        #Legg subdataset for entreprise til hoved-dataset
        if i != 0:
            a2.append(a_entreprise)
            entrepriser.append(a1[i][5])
            entreprise_index = entreprise_index + 1

        a_entreprise = []
        for k in range(1, len(prisbank_prosjekter)):
            a_entreprise.append(['Prisstigningsfaktor ' + prisbank_prosjekter[k][1], '', prisbank_prosjekter[k][2], '', '', ''])
        a_entreprise.append(['', '', '', '', '', ''])
        a_entreprise.append(['Entreprise: ', a1[i][5], '','','',''])
        a_entreprise.append(['','','','','',''])
        a_entreprise.append(['Beskrivelse', '', '', '', 'Enhet', 'Mengde', 'enhetspris', 'kostnad', '', '',
                   'Pris fra prosjekt', 'Opprinnelig enhetspris'])
        #a_entreprise.append(['Beskrivelse', '', '', '', 'Enhet', 'Mengde'])

    pr = 0

    # finn rad i prisbank som stemmer
    for j in range(len(prisbank)):
        # a1[i][0]:         entreprise
        # a1[i][1]:         family
        # a1[i][2]:         DN?
        # a1[i][3]:         mengde
        # a1[i][4]:         entreprise
        # prisbank[j][0]:   family
        # prisbank[j][1]:   DN
        # prisbank[j][2]:   enhet
        # prisbank[j][3]:   prosjekt index
        # prisbank[j][4]:   prosjekt
        # prisbank[j][5]:   enhetspris uregulert

        if a1[i][1] == prisbank[j][0] and a1[i][2] == prisbank[j][1]:
            # senere: if a1[i][1] == b1[j][0] and a1[i][2] == b1[j][1] and materiale = materiale:
            pr = j
            break
    if pr == 0 or prisbank[pr][3] == '' or prisbank[pr][4] == '' or prisbank[pr][5] == '':
        # ledd 2, 3 og 4 i OR-setning over er for å forhindre at script krasjer dersom det er tomme kolonner i prisbank
        #[a1[i][1] + ' ' + a1[i][2], a1[i][4], '', '', '', a1[i][3], '', '', '', a1[i][0], '', '', ''])
        #DN             + Beskrivelse, vinkel bend, -, -, -, MEngde, -, -, -, Entreprise(utgår), -, -, -,
        if a1[i][2] <> '':
            a_entreprise.append([a1[i][2] + ' ' + a1[i][1], '', '', '', '', a1[i][3],'', '=RC[-2]*RC[-1]'])
        else:
            a_entreprise.append([a1[i][1], '', '', '', '', a1[i][3], '', '=RC[-2]*RC[-1]'])
    else:
        #rad under krasjet tidligere dersom kolonne D, dvs. "valgt prosjekt", er tom i prisbank.
        a_entreprise.append([a1[i][2] + ' ' + a1[i][1], '', '', '', prisbank[pr][2], a1[i][3],
                   '=' + str(prisbank[pr][5]) + '*R' + str(int(prisbank[pr][3])) + 'C3', '=RC[-2]*RC[-1]', '', '', prisbank[pr][4],
                   prisbank[pr][5]])
        # DN             + Beskrivelse, vinkel bend, -, -, Enhet, MEngde, enhetspris, kostnad, -, Entreprise(utgår), -, pris fra prosjekt, Enhetspris uregulert
        #a_entreprise.append([a1[i][4], a1[i][1], a1[i][2], a1[i][3], ''])

# Legg subdataset for entreprise til hoved-dataset for siste entreprise
a2.append(a_entreprise)

# Assign your output to the OUT variable.
#OUT = a2

#Finne sti til mine dokumenter
mydoc = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents')

x = datetime.datetime.now()
timestamp = x.strftime("%Y_%m_%d %H_%M")
#timestamp = x.strftime("%Y %H-%M")
#debug.append(timestamp)

antall_entrepriser = entreprise_index +1


if antall_entrepriser == 1:
    filename = mydoc + '\kostnadsberegning_' + timestamp + '.xlsx'
    SaveListToExcel(filename, a2[0])
else:
    for i in range(antall_entrepriser):
        filename = mydoc + '\kostnadsberegning_'+ entrepriser[i] + '_' + timestamp + '.xlsx'
        SaveListToExcel(filename, a2[i])


xl.DisplayAlerts = True