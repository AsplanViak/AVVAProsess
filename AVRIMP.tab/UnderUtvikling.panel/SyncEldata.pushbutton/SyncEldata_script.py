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
__title__ = 'SyncEldata'  # Denne overstyrer navnet på scriptfilen
__author__ = 'Asplan Viak'  # Dette kommer opp som navnet på utvikler av knappen
__doc__ = "Klikk for å legge til flenser i prosjektet."  # Dette blir hjelp teksten som kommer opp når man holder musepekeren over knappen.
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

from System.Collections.Generic import List

from Autodesk.Revit.DB import Plumbing, IFamilyLoadOptions

import sys
import math

#clr.AddReference('ProtoGeometry')
#from Autodesk.DesignScript.Geometry import *

clr.AddReferenceByName(
    'Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel
from Microsoft.Office.Interop.Excel import XlListObjectSourceType, Worksheet, Range, XlYesNoGuess

#clr.AddReference("RevitServices")
#import RevitServices
#from RevitServices.Persistence import DocumentManager
#from RevitServices.Transactions import TransactionManager

#clr.AddReference("RevitNodes")
#import Revit

#clr.ImportExtensions(Revit.Elements)
#clr.ImportExtensions(Revit.GeometryConversion)

#clr.AddReference("RevitAPI")
import Autodesk

from System.Collections.Generic import List
#from System.Collections.Generic import *
from System import Guid
from System import Array


# Import RevitAPI
#clr.AddReference("RevitAPI")
import Autodesk
#from Autodesk.Revit.DB import *
#from Autodesk.Revit.DB.Analysis import *

#clr.AddReference('RevitAPIUI')

from Autodesk.Revit.UI import TaskDialog

#import System

import datetime

#pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
#sys.path.append(pyt_path)

debug_mode = 1
summary_mode = 1

def DebugPrint(tekst):
    if debug_mode == 1:
        print(tekst)
    return 1

def DetailedDebugPrint(tekst):
    if summary_mode == 1:
        print(tekst)
    return 1



def TryGetRoom(room, phase):
    try:
        inRoom = room.Room[phase]
    except:
        inRoom = None
        pass
    return inRoom


xl = Excel.ApplicationClass()
xl.Visible = False
xl.DisplayAlerts = False


def SaveListToExcel(filePath, exportData):
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets[1]
        # ws.title = 'komp'
        rows = len(exportData)
        cols = max(len(i) for i in exportData)
        a = Array.CreateInstance(object, rows, cols)  # row and column
        for r in range(rows):
            for c in range(cols):
                try:
                    a[r, c] = exportData[r][c]
                except:
                    a[r, c] = ''
        xlrange = ws.Range["A1", chr(ord('@') + cols) + str(rows)]
        xlrange.Value2 = a
        wb.SaveAs(filePath)
        #wb.Close(False)
        wb.Close()
        return True
    except:
        print('Feil med lagring av excel-eksport')
        return False

output_message = []
stat = []
errorReport = []

# Finn IO liste og andre parametre i ark "Kobling mot IO-liste"
har_funnet_IO-liste_ark_revit = 0
for sheet in FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements():
    if sheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString() == 'Kobling mot IO-liste':
        har_funnet_IO - liste_ark_revit = 1
        DebugPrint('Fant riktig ark med IO-liste info')
        IO_liste_filplassering = sheet.LookupParameter('IO-liste_filplassering').AsString()
        DebugPrint('IO-liste_filplassering :' + IO_liste_filplassering)
        Datablader_filplassering = sheet.LookupParameter('Datablader_filplassering').AsString()
        DebugPrint('Datablader_filplassering :' + Datablader_filplassering)
        tag_param = sheet.LookupParameter('AV_Tag-parameter_sync').AsString()
        DebugPrint('tag_param: ' + tag_param)
        sync_guid = sheet.LookupParameter('AV_sync_mot_GUID').AsInteger()
        DebugPrint('Sync mot GUID: ' + str(sync_guid))
        try:
            AV_room_link_str = sheet.LookupParameter('AV_room_link_str').AsString()
        except:
            AV_room_link_str = 'RIB'
        DebugPrint('AV_room_link_str: ' + AV_room_link_str)
        break

if (har_funnet_IO-liste_ark_revit==0) :
    button = UI.TaskDialogCommonButtons.None
    result = UI.TaskDialogResult.Ok
    UI.TaskDialog.Show('Synkronisering avbrutt', 'Fant ikke noe ark med navn "Kobling mot IO-liste" som angir filplassering IO-liste', button)
    #Sjekk at rad under ikke lukker revit!!!
    sys.exit(0)

# Filename for excel exports
filename = Document.PathName.GetValue(doc)
if "BIM 360://" in filename:
    sentralfil = filename
else:
    sentralfil = BasicFileInfo.Extract(doc.PathName).CentralPath
if not sentralfil:
    prosjektnavn = 'testprosjekt'
else:
    sentralfil = sentralfil.replace('\\', '/').replace('.', '/')
    sentralfil = sentralfil.split('/')
    prosjektnavn = sentralfil[-2]

### Finn riktig mappe for revit-sync
delt_filbane = IO_liste_filplassering.Split('\\')
s = '\\'
mappe = s.join(delt_filbane[:-1])
DebugPrint('Mappe :' + mappe)

x = datetime.datetime.now()
timestamp = x.strftime("%Y %H-%M")
DebugPrint(timestamp)

excel_filenames = [mappe + '\Komponenter_' + prosjektnavn + '.xlsx',
                   mappe + '\Komponenter_skjema_' + prosjektnavn + '.xlsx',
                   mappe + '\Backup_presync\\' + prosjektnavn + '_3D - ' + timestamp + '.xlsx',
                   mappe + '\Backup_presync\\' + prosjektnavn + '_skjema - ' + timestamp + '.xlsx']

komp_3d = []
komp_3d.append(['element_id', 'Family', 'FamilyType', 'TAG', 'Rom', 'System Type', 'X', 'Y', 'Z', 'Tegning'])
komp_skjema = []
komp_skjema.append(['element_id', 'TAG', 'Family', 'FamilyType', 'Tegning'])
presync_3d = []
presync_skjema = []

msg = TaskDialog

# Finn linker, for å lage liste over alle rom
linkCollector = Autodesk.Revit.DB.FilteredElementCollector(doc)
linkInstances = linkCollector.OfClass(Autodesk.Revit.DB.RevitLinkInstance)

# linkDoc, linkName, linkPath = [], [], []
linkname_with_rooms = AV_room_link_str
link_highscore = 0
DebugPrint("str(linkname_with_rooms) :" + str(linkname_with_rooms))
for i in linkInstances:
    link_score = 5
    if str(linkname_with_rooms) == 'null' or str(linkname_with_rooms) == '' or str(linkname_with_rooms) is None:
        DebugPrint("undefined linkname_with_rooms")
        if 'RIB' in i.Name or 'Bygg' in i.Name:
            link_score += (i.Name.count('RIB') + i.Name.count('Bygg'))
            if 'RIBr' in i.Name:
                link_score -= 1
            if link_score > link_highscore:
                link_highscore = link_score
                DebugPrint("udef. AV_room_link_str. Fant link med RIB/ARK" + i.Name)
                DebugPrint("link_score = " + str(link_score))
                linkDoc = i.GetLinkDocument()
    elif linkname_with_rooms in i.Name:
        DebugPrint("defined linkname_with_rooms string")
        link_score += i.Name.count(linkname_with_rooms)
        if 'RIBr' in i.Name:
            link_score -= (i.Name.count('RIBr'))
        if 'LARK' in i.Name:
            link_score -= (i.Name.count('LARK'))
        if link_score > link_highscore:
            link_highscore = link_score
            DebugPrint("fant link med RIB/ARK: " + i.Name)
            DebugPrint("link_score: " + str(link_score))
            linkDoc = i.GetLinkDocument()

cat_list = [BuiltInCategory.OST_Rooms]
typed_list = List[BuiltInCategory](cat_list)
filter = ElementMulticategoryFilter(typed_list)

try:
    rooms = FilteredElementCollector(linkDoc).WherePasses(filter).WhereElementIsNotElementType().ToElements()

    if len(rooms) == 0:
        DebugPrint("Feil: Mangler definesjoner på rom i link. Kan ikke sende info om romplassering til database.")

except:
    # if error accurs anywhere in the process catch it
    import traceback

    errorReport.append(traceback.format_exc())
    rooms = []
    DebugPrint(
        "Feil : Kan ikke lese romdata fra link. Kan skyldes at ARK/RIB link ikke er lastet inn, eller er i lukket workset.")


#IOliste = IN[0]
# LES INN IO-LISTE FRA EXCEL
try:
    wb_IO_liste = xl.Workbooks.Open(IO_liste_filplassering)
except:
    button = UI.TaskDialogCommonButtons.None
    result = UI.TaskDialogResult.Ok
    UI.TaskDialog.Show('Synkronisering avbrutt', 'Fant ikke excel-dokument med IO-liste på plassering angitt i ', button)

#Finn sheet med IO-liste. Bruker sheet1 dersom ingen treff på navn
try:
    ws_IO_liste = wb_IO_liste.Worksheets('IO-liste')
except:
    try:
        ws_IO_liste = wb_IO_liste.Worksheets('IOliste')
    except:
        ws_IO_liste = wb_IO_liste.Worksheets[1]]
#used = ws_IO_liste.UsedRange
cols = ws_IO_liste.UsedRange.Columns.Count
rows = ws_IO_liste.UsedRange.Rows.Count
#print('cols: '+ str(cols))
#print('rows: '+ str(rows))
IOliste =[]
for i in range(1,rows+1):
    rad = []
    for j in range(1,cols+1):

        rad.append(ws_IO_liste.Range[chr(ord('@') + j) + str(i)].Text)

    IOliste.append(rad)
#IO_liste_range = ws_IO_liste.Range["A1", chr(ord('@') + cols) + str(rows)]
#IO_liste_range = ws_IO_liste.Range["A1", "A" + str(rows)]
#IO_liste_range = ws_IO_liste.Range["A1", "A4"]
#rad under gir kanskje problemer med array
#IOliste = IO_liste_range.Value2
#IOliste = IO_liste_range.Text
#IOliste = IO_liste_range.Value

#print('default øæøå')
print(IOliste)
#print('dearray')
#print(IOliste[0])


if tag_param is None or tag_param == '':
    tag_param = 'TAG'

if sync_guid is None or sync_guid == 0:
    sync_guid = False

DebugPrint('Tag parameter: ' + str(tag_param))
DebugPrint('sync_guid: ' + str(sync_guid))

parametre_shared_name = ['Entreprise', 'Tekstlinje 1', 'Tekstlinje 2', 'Driftsform', 'AV_MMI', 'BUS-kommunikasjon',
                         'Sikkerhetsbryter', 'Kommentar', 'Fabrikat', 'Modell', 'Merkespenning', 'Merkeeffekt',
                         'Merkestrøm', 'Status', 'Access_TagType', 'Access_TagType beskrivelse', 'IO-er', 'Datablad']

print('parametre_shared_name')
print(parametre_shared_name)


print('parametre_shared_name[12]')
print(parametre_shared_name[12])

print('IOliste[0][2]')
print(IOliste[0][2])

if(IOliste[0][2] == parametre_shared_name[12]):
    print('samme')

parametre_guid = ['2c78b93c-2c2d-4bf5-a4cd-b5ab37d40b3f', '88e7a061-da67-44a8-bdab-19fd0e111277',
                  '8bd618c5-4b04-4089-8c1c-d3329c359af7', 'da7bed97-096f-4949-840f-3125bdf40605',
                  'fd766c10-96b3-470b-aa1e-3b1b5b572492', '18f9c00a-61cf-4429-b022-311b2ce6a667',
                  'daf45c3a-35fc-4994-a29a-180483305de1', '3df6ebff-09dd-475f-8ca7-6eb31e697fd5',
                  'c4a831f8-4d5f-46a3-a40c-41f987c910f6', 'e854697d-4f77-4882-b709-72bcc27ee040',
                  'c6dde6a0-ab66-4102-9e09-30ab645e56bb', 'ed0e58dd-3954-47f7-b7eb-f561f2b55ff9',
                  '55e65789-9ba8-40fe-9795-287755774934', '57e99ba0-ecc1-43aa-8de5-60155ae1e99d',
                  '78d27e5c-b652-4ef5-b19c-767de086fe46', '2603489f-989b-4df2-b87f-0cbfbe9d4f5b',
                  '1c7f04a2-ed1a-412e-a7c1-170dde0c203c', 'bf7410c5-c78b-4463-8b39-fefedd6b4ac7']
for i, g in enumerate(parametre_guid):
    parametre_guid[i] = Guid(g)

parametre_shared_name_lc = [x.lower() for x in parametre_shared_name]
# DebugPrint('parametre_shared_name_lc ')
# DebugPrint(parametre_shared_name_lc)

parametre_signalinfo_lc = ['tekst', 'signaltag', 'type', 'signalkilde', 'spenning', 'tilleggstekst']
parametre_ikke_sync_lc = ['sortering', 'tag', 'guid', ' tfm11fksamlet']

# Plan etterhvert: laste inn parametre_project fra kobling_IO_liste_sheet i revit
# En enda bedre plan: Last inn alle ubrukte parametre fra IO liste her. Slik at de ikke trenger å defineres noe sted.
# parametre_project_name = ['Byggtype', 'MMI']
parametre_project_name = []

TAG_guid = '141d33b4-0f91-4dd8-a8b6-be1fa232d39f'
TFM11FkSamlet_guid = '6b52eb8b-6935-45f9-a509-bb76724ba272'

if tag_param == 'TAG':
    tguid = Guid(TAG_guid)
elif tag_param == 'TFM11FkSamlet':
    tguid = Guid(TFM11FkSamlet_guid)
else:
    tguid = -1

# finn kolonne i Io liste med tag/tfm, og finn project parametre
tag_kol = -1


for j, celle in enumerate(IOliste[0]):
    try:
        # if celle.lower() == tag_param.lower():
        if celle.lower() == 'tag':  # Bruker tag her og ikke tag_param.lower(), siden det alltid er TAG som brukes i eksport fra database.
            # Kan evt. splitte dette opp i to forskjellige parametre i ark med parametre.

            tag_kol = j
        if celle.lower() == 'guid':
            GUID_kol = j
        if celle.lower() not in parametre_shared_name_lc:
            if celle.lower() not in parametre_signalinfo_lc and celle.lower() not in parametre_ikke_sync_lc:
                parametre_project_name.append(celle)
                DebugPrint('project param lagt til: ' + celle.lower())

        DebugPrint('j, celle, try: ' + str(j) + ' ' + celle)

    except:
        DebugPrint('j, celle, continue: ' + str(j) + ' ' + celle)
        # Tidligere stod det break her. Tror continue er bedre.
        continue

# sjekk om det ble funnet kolonne med tag/tfm
if tag_kol == (-1):
    # tag_kol not found
    output_message.append('Fant ikke kolonne ved navn ' + tag_param + ' i IO liste excel-ark. Sync avbrutt')
    ## må legge til en avbryt-funksjon her

# Finn alle parametre som brukes i IO_liste
parametre_shared_IO_liste_kolonne = []
parametre_shared_IO_liste_name = []
parametre_shared_IO_liste_guid = []
parametre_project_IO_liste_kolonne = []
parametre_project_IO_liste_name = []

# shared
for i, navn in enumerate(parametre_shared_name):
    for j, celle in enumerate(IOliste[0]):
        try:
            if celle.lower() == navn.lower():
                parametre_shared_IO_liste_kolonne.append(j)
                parametre_shared_IO_liste_name.append(navn)
                parametre_shared_IO_liste_guid.append(parametre_guid[i])
                break
        except:
            break
# project
for i, navn in enumerate(parametre_project_name):
    for j, celle in enumerate(IOliste[0]):
        try:
            if celle.lower() == navn.lower():
                parametre_project_IO_liste_kolonne.append(j)
                parametre_project_IO_liste_name.append(navn)
                break
        except:
            break

cat_list = [BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_GenericModel, BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_Sprinklers,
            BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_DetailComponents]

# typed_list = List[BuiltInCategory](cat_list)
# filter = ElementMulticategoryFilter(typed_list)
# EQ = FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements()


transaction = DB.Transaction(doc)
transaction.Start("Sync eldata")



# loop all categories
for cat in cat_list:
    DebugPrint(cat)
    DetailedDebugPrint(cat)

    # sjekk om tag/tfm parameter finnes, og om den er shared, og om den er definert med samme GUID som den riktige shared parameteren
    tag_cat_status = (-1)

    # lag lister over parametre som finnes for category
    p_s_IO_cat_kol = []  # parametre_shared_IO_liste_category_kolonne
    p_s_IO_cat_name = []
    p_s_IO_cat_guid = []
    p_r_IO_cat_kol = []
    p_r_IO_cat_name = []

    # må kanskje bruke try: her for å unngå feil for categorier som ikke er i bruk, og dermed ikke har firstElement
    # find all shared parameters defined for the category
    try:
        for param in FilteredElementCollector(doc).OfCategory(
                cat).WhereElementIsNotElementType().FirstElement().Parameters:
            if param.IsShared == True:
                for i, guid in enumerate(parametre_shared_IO_liste_guid):
                    if param.GUID == guid:
                        p_s_IO_cat_kol.append(parametre_shared_IO_liste_kolonne[i])
                        p_s_IO_cat_name.append(parametre_shared_IO_liste_name[i])
                        p_s_IO_cat_guid.append(parametre_shared_IO_liste_guid[i])
                    if param.GUID == tguid:
                        tag_cat_status = 1  # status 1 betyr at den finnes som shared parameter, og at definisjonen stemmer overens med offisiel AV standard.
    except:
        # bør legge inn en advarsel her
        DebugPrint('failed parameter testing')
        continue

    DebugPrint('Size p_s_IO_cat_name: ' + str(len(p_s_IO_cat_name)))
    # DebugPrint('p_s_IO_cat_name: ')
    # DebugPrint(p_s_IO_cat_name)
    # sorting params

    # Z1 = [x for _,x in sorted(zip(p_s_IO_cat_kol,p_s_IO_cat_name))]
    p_s_IO_cat_name = [x for (y, x) in sorted(zip(p_s_IO_cat_kol, p_s_IO_cat_name), key=lambda pair: pair[0])]
    p_s_IO_cat_guid = [x for (y, x) in sorted(zip(p_s_IO_cat_kol, p_s_IO_cat_guid), key=lambda pair: pair[0])]
    p_s_IO_cat_kol = sorted(p_s_IO_cat_kol)
    # DebugPrint(p_s_IO_cat_name)
    # DebugPrint(p_s_IO_cat_kol)
    # DebugPrint(p_s_IO_cat_guid)

    # find all shared parameters not defined for category
    p_s_IO_cat_unused_kol = list(set(parametre_shared_IO_liste_kolonne) - set(p_s_IO_cat_kol))
    # p_s_IO_cat_unused_kol = [item for item in parametre_shared_IO_liste_kolonne if item not in p_s_IO_cat_kol]     #gjør samme som linje over
    p_s_IO_cat_unused_name = list(set(parametre_shared_IO_liste_name) - set(p_s_IO_cat_name))
    # p_s_IO_cat_unused_name = [item for item in parametre_shared_IO_liste_name if item not in p_s_IO_cat_name]		#gjør samme som linje over

    DebugPrint('Size p_s_IO_cat_unused_kol : ' + str(len(p_s_IO_cat_unused_kol)))
    # DebugPrint(p_s_IO_cat_unused_kol)
    DebugPrint('p_s_IO_cat_unused_name : ')
    DebugPrint(p_s_IO_cat_unused_name)

    # merge project params and unused shared params
    p_IO_cat_uandp_kol = p_s_IO_cat_unused_kol + parametre_project_IO_liste_kolonne  # unused and project parameters
    p_IO_cat_uandp_name = p_s_IO_cat_unused_name + parametre_project_IO_liste_name  # unused and project parameters

    DebugPrint('parametre_project_IO_liste_name')
    DebugPrint(parametre_project_IO_liste_name)

    DebugPrint('Size p_IO_cat_uandp_name: ' + str(len(p_IO_cat_uandp_name)))

    # find remaining parameters by name
    for param in FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().FirstElement().Parameters:
        #DetailedDebugPrint('param.Definition.Name : ' + param.Definition.Name)
        for i, name in enumerate(p_IO_cat_uandp_name):
            #DetailedDebugPrint('name in p_IO_cat_uandp_name: ' + name)
            if param.Definition.Name == name:
                p_r_IO_cat_kol.append(p_IO_cat_uandp_kol[i])  # r = remaining
                p_r_IO_cat_name.append(p_IO_cat_uandp_name[i])
            if param.Definition.Name == tag_param:
                if tag_cat_status == -1:
                    tag_cat_status = 0  # status 0 betyr at den finnes, men at man ikke kan bruke GUID (enten fordi tag_param er ulik både tfm og tag, eller fordi det ikke er brukt shared parameter for denne)

    if tag_cat_status == -1 and cat == BuiltInCategory.OST_DetailComponents:
        if tag_param == 'TFM11FkSamlet':
            tag_cat_status = 2  # Betyr at man bruker denne som tag: SystemVar + '-' + TFM11FkKompGruppe + TFM11FkKompLNR

    DebugPrint('Size p_r_IO_cat_name: ' + str(len(p_r_IO_cat_name)))
    DebugPrint(p_r_IO_cat_name)

    DebugPrint('tag_cat_status: ' + str(tag_cat_status))
    if tag_cat_status == (-1):
        # category mangler definert tag-parameter
        # går til neste category
        continue

    # loop elements in category
    EQ = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
    n_elements = 0
    
    for k in EQ:
        # Tag reset
        tag = ''

        # sjekk om tag/tfm er blank, og finn tekstverdi på tag i revit
        if tag_cat_status == 1:
            if k.get_Parameter(tguid).AsString() == 'null' or k.get_Parameter(
                    tguid).AsString() == '=-' or k.get_Parameter(tguid).AsString() == '' or k.get_Parameter(
                tguid).AsString() is None:
                # gå til neste element dersom blank tag/tfm
                DetailedDebugPrint('blank tag/tfm (shared)')
                continue
            tag = k.get_Parameter(tguid).AsString()
            # DebugPrint('k.get_Parameter(tguid).AsString() : ' + k.get_Parameter(tguid).AsString())
            # DebugPrint('k.LookupParameter(TAG).AsString() : ' + k.LookupParameter('TAG').AsString())
        elif tag_cat_status == 0:
            if k.LookupParameter(tag_param).AsString() == 'null' or k.LookupParameter(
                    tag_param).AsString() == '=-' or k.LookupParameter(tag_param).AsString() == '' or k.LookupParameter(
                tag_param).AsString() is None:
                # gå til neste element dersom blank tag/tfm
                DetailedDebugPrint('blank tag/tfm (string)')
                continue
            tag = k.LookupParameter(tag_param).AsString()
        elif tag_cat_status == 2:
            # skjema TFM
            try:
                SystemVar = i.LookupParameter('SystemVar').AsString()
                TFM11FkKompLNR = i.LookupParameter('TFM11FkKompLNR').AsString()
                TFM11FkKompGruppe = i.Symbol.LookupParameter('TFM11FkKompGruppe').AsString()
                tag = r"'" + '=' + SystemVar + '-' + TFM11FkKompGruppe + TFM11FkKompLNR
            except:
                DetailedDebugPrint('feil ved sammenslåing/avlesing av parametre som inngår i TFM for skjema')
                continue
        DebugPrint('Tag: ' + tag)

        n_elements += 1

        #############################################################################################################
        # SYNC DATA TIL REVIT. Inkl. eksport presync
        #############################################################################################################
        # guid-syncrom_eksport.append(Element.Name.GetValue(r))
        # break
        IO_liste_row = (-1)
        if sync_guid == True and cat <> BuiltInCategory.OST_DetailComponents:  # Kan ikke bruke guid-sync mot skjema. Kun tag/TFM.
            TAG_match_id = (-1)
            for b in range(1, len(IOliste)):
                if tag == IOliste[b][tag_kol]:
                    TAG_match_id = b
                if str(k.UniqueId) == IOliste[l][GUID_kol]:
                    IO_liste_row = b
                    break
            # resten av code under kjøres dersom ikke match på GUID:
            if TAG_match_id != (-1):
                IO_liste_row = TAG_match_id
        else:
            # tag/tfm-sync
            for b in range(1, len(IOliste)):
                # if k.LookupParameter('TAG').AsString() == IOliste[l][tag_kol]:
                if tag == IOliste[b][tag_kol]:
                    IO_liste_row = b
                    break
        DetailedDebugPrint('IO_liste_row :' + str(IO_liste_row))

        # Presync data
        # lag tom header list for denne category (antall parametre kan variere mellom categories)
        if n_elements == 1:
            presync_top_row = ['TAG']
        # lag tom list til presync-data for dette elementet
        if cat <> BuiltInCategory.OST_DetailComponents:
            presync_3d_row = [tag]
        else:
            presync_skjema_row = [tag]
        # oppdater_eldata(IO_liste_row, k)


        # loop shared params
        for i, kol in enumerate(p_s_IO_cat_kol):
            DebugPrint('Looping shared params')
            # presync header
            if n_elements == 1:
                DebugPrint('presync header')
                presync_top_row.append(p_s_IO_cat_name[i])
            # presync data
            try:
                DebugPrint('Check if detail item')
                if cat <> BuiltInCategory.OST_DetailComponents:
                    DebugPrint('Is detail item')
                    presync_3d_row.append(k.get_Parameter(p_s_IO_cat_guid[i]).AsString())
                else:
                    DebugPrint('Is not detail item')
                    presync_skjema_row.append(k.get_Parameter(p_s_IO_cat_guid[i]).AsString())
            except:
                DebugPrint('Something went wrong with check if detail item')
                if cat <> BuiltInCategory.OST_DetailComponents:
                    presync_3d_row.append('')
                else:
                    presync_skjema_row.append('')
            if IO_liste_row != (-1):
                # sync
                DebugPrint('Syncing')
                IOliste_tekst = IOliste[IO_liste_row][kol]
                if IOliste_tekst is None:
                    IOliste_tekst = ' '
                DebugPrint('IOliste_tekst: ' + str(IOliste_tekst))
                try:
                    res = k.get_Parameter(p_s_IO_cat_guid[i]).Set(IOliste_tekst)
                    if (res):
                        DebugPrint('shared parameter ' + p_s_IO_cat_name[i] + ': ok')
                    else:
                        DebugPrint('shared parameter ' + p_s_IO_cat_name[i] + ': feil')
                except Exception, ex:
                    DebugPrint('shared parameter ' + p_s_IO_cat_name[i] + ': exception')
                    DebugPrint('k.get_Parameter(' + str(p_s_IO_cat_guid[i]) + ').Set(' + str(IOliste_tekst) + ')')
                    DebugPrint(ex.message)

        # loop remaining params
        for i, kol in enumerate(p_r_IO_cat_kol):
            # presync header
            if n_elements == 1:
                presync_top_row.append(p_r_IO_cat_name[i])
            # presync data
            try:
                if cat <> BuiltInCategory.OST_DetailComponents:
                    presync_3d_row.append(k.LookupParameter(p_r_IO_cat_name[i]).AsString())
                else:
                    presync_skjema_row.append(k.LookupParameter(p_r_IO_cat_name[i]).AsString())
            except:
                if cat <> BuiltInCategory.OST_DetailComponents:
                    presync_3d_row.append('')
                else:
                    presync_skjema_row.append('')
            if IO_liste_row != (-1):
                # sync
                IOliste_tekst = IOliste[IO_liste_row][kol]
                if IOliste_tekst is None:
                    IOliste_tekst = ' '
                DebugPrint('IOliste_tekst: ' + str(IOliste_tekst))
                try:
                    res = k.LookupParameter(p_r_IO_cat_name[i]).Set(IOliste_tekst)
                    if (res):
                        DebugPrint('remaining parameter ' + p_r_IO_cat_name[i] + ': ok')
                    else:
                        DebugPrint('remaining parameter ' + p_r_IO_cat_name[i] + ': feil')
                except Exception, ex:
                    DebugPrint('remaining parameter ' + p_r_IO_cat_name[i] + ': exception')
                    DebugPrint('k.LookupParameter(' + str(p_r_IO_cat_name[i]) + ').Set(' + str(IOliste_tekst) + ')')
                    DebugPrint(ex.message)

            # Update TAG if GUID sync????
            # Funksjon må legges til her!!



        if cat <> BuiltInCategory.OST_DetailComponents:
            if n_elements == 1:
                presync_3d.append(presync_top_row)
            presync_3d.append(presync_3d_row)
        else:
            if n_elements == 1:
                presync_skjema.append(presync_top_row)
            presync_skjema.append(presync_skjema_row)

        #############################################################################################################
        # LAG DATASET TIL EXCEL EKSPORT
        #############################################################################################################

        # komp_3d
        if cat <> BuiltInCategory.OST_DetailComponents:

            # finn ut hvilket rom den ligger i
            pt = k.Location.Point
            room_hit = 0
            for r in rooms:
                if r.IsPointInRoom(pt):
                    rom = Element.Name.GetValue(r)
                    room_hit = 1
                    break
                # else:
            #	for phase in doc.Phases:
            # inRoom = TryGetRoom(family, phase)
            #		inRoom = TryGetRoom(k, phase)
            #		if inRoom != None and inRoom.ToDSType(True).Name == _room.ToDSType(True).Name:
            #			rom_eksport.append(Element.Name.GetValue(r))
            #			break
            if room_hit == 0:
                rom = ''

            # Tegning
            # kjører loop tegninger lengre ned som egen loop

            # System type
            try:
                systemtype = k.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
            except:
                systemtype = 'udefinert system type'
            # Finn family
            try:
                # family = k.Symbol.FamilyName
                family = k.Symbol.FamilyName
            except:
                family = 'udefinert familie'
            # Finn familytype
            try:
                familytype = k.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            except:
                familytype = 'udefinert familietype'
            # Finn XYZ koord
            try:
                # alt 1
                # bb = k.get_BoundingBox(None)
                # if not bb is None:
                #    centre = bb.Min + (bb.Max - bb.Min) / 2
                #    x = centre.X*304.8
                #    y = centre.Y*304.8
                # level = k.ReferenceLevel
                # z = centre.Z - level.Elevation
                #    z = centre.Z*304.8
                # alt 2
                centre = k.Location.Point
                x = centre.X * 304.8
                y = centre.Y * 304.8
                # level = k.ReferenceLevel
                # z = centre.Z - level.Elevation
                z = centre.Z * 304.8
            except:
                x = ''
                y = ''
                z = ''

            # header: komp_3d.append(['l__element__l', 'Family','FamilyType', 'TAG','Rom','System Type','X','Y','Z', 'Tegning'])
            komp_3d.append([k.Id, family, familytype, tag, rom, systemtype, x, y, z, ''])
            # DebugPrint('komp_3d.append() ' + tag + ' , ' + family])

        # komp_skjema
        else:
            # Finn family
            DebugPrint('skjemaelement')
            DebugPrint(tag)
            # Finn family
            try:
                # family = k.Symbol.FamilyName
                family = k.Symbol.FamilyName
            except:
                family = 'udefinert familie'
            # Finn familytype
            try:
                familytype = k.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            except:
                familytype = 'udefinert familietype'
                # header : komp_skjema.append(['element_id', 'TAG', 'Family', 'FamilyType', 'Tegning'])
            komp_skjema.append([k.Id, tag, family, familytype, ''])

    DebugPrint('n_elements: ' + str(n_elements))

#########################################################################################################################################
# Loop tegninger
##############################################################################################################################################

# Liste/collector over alle viewports
paramvalue = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER))
ruleval = DB.FilterStringGreater()
filterrule = DB.FilterStringRule(paramvalue, ruleval, "", False)
paramfi = DB.ElementParameterFilter(filterrule)
viewscollector = DB.FilteredElementCollector(doc).OfClass(DB.View).WherePasses(paramfi)

for v in viewscollector:
    viewelemcollector = DB.FilteredElementCollector(doc, v.Id).ToElementIds()
    vparam = v.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER)  # type: DB.Parameter
    # loop komp 3d
    DebugPrint('Tegning: ' + vparam.AsString())
    for n, e in enumerate(komp_3d):
        if n == 0:
            # må hoppe over tabell headers
            continue
        # DebugPrint('Tag : ' + e[2])
        if e[0] in viewelemcollector:
            DebugPrint('Treff ' + vparam.AsString() + ' - ' + e[2])
            if komp_3d[n][9] == '':
                komp_3d[n][9] = vparam.AsString()
            else:
                komp_3d[n][9] += ', ' + vparam.AsString()
    # loop komp skjema
    for n, e in enumerate(komp_skjema):
        if n == 0:
            # må hoppe over tabell headers
            continue
        # DebugPrint('Tag : ' + e[1])
        if e[0] in viewelemcollector:
            DebugPrint('Treff ' + vparam.AsString() + ' - ' + e[1])
            if komp_skjema[n][4] == '':
                komp_skjema[n][4] = vparam.AsString()
            else:
                komp_skjema[n][4] += ', ' + vparam.AsString()

# Assign your output to the OUT variable.
# OUT = debug

#####################################################
############Eksporter til excel
#####################################################

excel_eksport = [komp_3d, komp_skjema, presync_3d, presync_skjema]

for i in range(4):
#if 0:
    # def SaveListToExcel(filePath, exportData)
    if SaveListToExcel(excel_filenames[i], excel_eksport[i]):
        DetailedDebugPrint('Fileksport success ' + str(i))
    else:
        DetailedDebugPrint('Fileksport failed ' + str(i))
        # Lag melding her: 'Kunne ikke eksportere fil ' + excel_filenames[i] + '. Sjekk om fil allerede er åpen, og lukk den, og prøv på ny.'

#OUT = [excel_eksport, excel_filenames, debug, debug_details, debug_summary, errorReport]
# OUT = [parameter, parameter_kolonne]
# OUT = parameter_kolonne

wb_IO_liste.Close()


xl.DisplayAlerts = True
xl.Visible = True

transaction.Commit()

#wb_IO_liste.close()



button = UI.TaskDialogCommonButtons.None
result = UI.TaskDialogResult.Ok
UI.TaskDialog.Show('Sync eldata ferdig', 'report_tekst', button)
