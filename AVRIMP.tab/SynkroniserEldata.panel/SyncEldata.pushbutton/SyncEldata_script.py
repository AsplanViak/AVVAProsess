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
__title__ = 'Synkroniser eldata'  # Denne overstyrer navnet på scriptfilen
__author__ = 'Asplan Viak'  # Dette kommer opp som navnet på utvikler av knappen
__doc__ = "Klikk for å synkronisere eldata med IO-liste/database."  # Dette blir hjelp teksten som kommer opp når man holder musepekeren over knappen.
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

#clr.AddReference("RevitNodes")

from Autodesk.Revit import UI, DB

from Autodesk.Revit.DB import *

from System.Collections.Generic import List

import sys
import math
import os
import time
import csv

clr.AddReferenceByName(
    'Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

from System.Collections.Generic import List
from System import Guid
from System import Array

import Autodesk
from Autodesk.Revit.UI import TaskDialog
import datetime
import string


#printer en del meldinger til terminal i revit dersom man setter debug_mode til 1
def DebugPrint(tekst):
    if debug_mode == 1:
        print(tekst)
    return 1

#printer en del meldinger til terminal i revit dersom man setter summary_mode til 1
def SummaryPrint(tekst):
    if summary_mode == 1:
        print(tekst)
    return 1

try:
    debug_mode_id = GlobalParametersManager.FindByName(doc, "debug_mode")
    debug_mode = doc.GetElement(debug_mode_id).GetValue().Value
except:
    debug_mode = 0

try:
    summary_mode_id = GlobalParametersManager.FindByName(doc, "summary_mode")
    summary_mode = doc.GetElement(summary_mode_id).GetValue().Value
except:
    summary_mode = 0


def TryGetRoom(room, phase):
    try:
        inRoom = room.Room[phase]
    except:
        inRoom = None
        pass
    return inRoom

#definer excel-applikasjon
xl = Excel.ApplicationClass()
xl.Visible = False
xl.DisplayAlerts = False

#funksjon for å returnere kolonnenavn for excel.
def n2a(n,b=string.ascii_uppercase):
   d, m = divmod(n,len(b))
   return n2a(d-1,b)+b[m] if d else b[m]

#funksjon for å lagre dataset (2d-lister) til excel, og lukke dem etterpå
def SaveListToExcel(filePath, exportData):
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets[1]
        ws.name = 'komp'
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
        wb.Close()
        return True
    except:
        DebugPrint('Feil med lagring av excel-eksport')
        return False

def OppdaterEldata(cat, IO_liste_row, k, n_elements, p_s_IO_cat_kol, p_s_IO_cat_guid, p_s_IO_cat_name, p_r_IO_cat_name, p_r_IO_cat_kol):
    # loop shared params
    global presync_top_row
    global presync_3d_row
    global presync_skjema_row
    global IOliste

    for i, kol in enumerate(p_s_IO_cat_kol):
        # DebugPrint('Looping shared params')
        if n_elements == 1:
            DebugPrint('i: ' + str(i) + ', kol: ' + str(kol))
        # presync header
        if n_elements == 1:
            # DebugPrint('presync header')
            presync_top_row.append(p_s_IO_cat_name[i])
        # presync data
        try:
            # DebugPrint('Check if detail item')
            if cat <> BuiltInCategory.OST_DetailComponents:
                # DebugPrint('Is not detail item')
                presync_3d_row.append(k.get_Parameter(p_s_IO_cat_guid[i]).AsString())
            else:
                # DebugPrint('Is detail item')
                presync_skjema_row.append(k.get_Parameter(p_s_IO_cat_guid[i]).AsString())
        except:
            if cat <> BuiltInCategory.OST_DetailComponents:
                presync_3d_row.append('')
            else:
                presync_skjema_row.append('')
        if IO_liste_row != (-1):
            # sync
            # DebugPrint('Syncing')
            if p_s_IO_cat_name[i] == 'TrengerSignalKabel' or p_s_IO_cat_name[i] == 'TrengerStrømKabel':

                if IOliste[IO_liste_row][kol] == 'Yes':
                    IOliste_tekst = 1
                elif IOliste[IO_liste_row][kol] == 'No':
                    IOliste_tekst = 0
            else:
                IOliste_tekst = IOliste[IO_liste_row][kol]

            if IOliste_tekst is None:
                IOliste_tekst = ' '
            # DebugPrint('IOliste_tekst: ' + str(IOliste_tekst))
            try:

                #rad under er den som utfører selve endringen av parameter-verdi i Revit
                res = k.get_Parameter(p_s_IO_cat_guid[i]).Set(IOliste_tekst)
                if (res):
                    SummaryPrint('shared parameter ' + p_s_IO_cat_name[i] + ': ok')
                else:
                    SummaryPrint('shared parameter ' + p_s_IO_cat_name[i] + ': feil')
            except Exception, ex:
                DebugPrint('shared parameter ' + p_s_IO_cat_name[i] + ': exception')
                DebugPrint('k.get_Parameter(' + str(p_s_IO_cat_guid[i]) + ').Set(' + str(IOliste_tekst) + ')')
                DebugPrint(ex.message)

    # loop remaining params
    for j, kol2 in enumerate(p_r_IO_cat_kol):
        # presync header
        if n_elements == 1:
            presync_top_row.append(p_r_IO_cat_name[j])
        # presync data
        try:
            if cat <> BuiltInCategory.OST_DetailComponents:
                presync_3d_row.append(k.LookupParameter(p_r_IO_cat_name[j]).AsString())
            else:
                presync_skjema_row.append(k.LookupParameter(p_r_IO_cat_name[j]).AsString())
        except:
            if cat <> BuiltInCategory.OST_DetailComponents:
                presync_3d_row.append('')
            else:
                presync_skjema_row.append('')
        if IO_liste_row != (-1):
            # sync
            if p_r_IO_cat_name[j] == 'TrengerSignalKabel' or p_r_IO_cat_name[j] == 'TrengerStrømKabel':

                if IOliste[IO_liste_row][kol2] == 'Yes':
                    IOliste_tekst = 1
                elif IOliste[IO_liste_row][kol2] == 'No':
                    IOliste_tekst = 0
            else:
                IOliste_tekst = IOliste[IO_liste_row][kol2]
            if IOliste_tekst is None:
                IOliste_tekst = ' '
            # DebugPrint('IOliste_tekst: ' + str(IOliste_tekst))
            try:
                #rad under er den som utfører selve endringen av parameter-verdi i Revit
                ############SKAL TROLIG STÅ kol ISTEDET FOR j PÅ RAD UNDER!!!!!!!!!!!!!!!!!!!!!
                res = k.LookupParameter(p_r_IO_cat_name[j]).Set(IOliste_tekst)
                if (res):
                    SummaryPrint('remaining parameter ' + p_r_IO_cat_name[j] + ': ok')
                else:
                    SummaryPrint('remaining parameter ' + p_r_IO_cat_name[j] + ': feil')
            except Exception, ex:
                DebugPrint('remaining parameter ' + p_r_IO_cat_name[j] + ': exception')
                DebugPrint('k.LookupParameter(' + str(p_r_IO_cat_name[j]) + ').Set(' + str(IOliste_tekst) + ')')
                DebugPrint(ex.message)

    # Update TAG if GUID sync????
    # Funksjon må legges til her!!
    return

def MainFunction():

    global presync_top_row
    global presync_3d_row
    global presync_skjema_row
    global IOliste

    #kritiske feil:
    errorReport = ""
    #ikke-kritiske feil og generell info:
    summaryReport = ""
    start = time.time()

    # Finn IO liste og andre parametre i ark "Kobling mot IO-liste"
    har_funnet_IOliste_ark_revit = 0
    #collector for alle sheets
    for sheet in FilteredElementCollector(doc).OfCategory(
            BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements():
        if sheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString() == 'Kobling mot IO-liste':
            har_funnet_IOliste_ark_revit = 1
            DebugPrint('Fant riktig ark med IO-liste info')
            IO_liste_filplassering = sheet.LookupParameter('IO-liste_filplassering').AsString()
            DebugPrint('IO-liste_filplassering :' + IO_liste_filplassering)
            Datablader_filplassering = sheet.LookupParameter('Datablader_filplassering').AsString()
            DebugPrint('Datablader_filplassering :' + Datablader_filplassering)
            try:
                tag_param = sheet.LookupParameter('AV_Tag-parameter_sync').AsString()
                DebugPrint('tag_param: ' + tag_param)
            except:
                tag_param = 'TAG'
                DebugPrint('Fant ikke tag_parameter  i sheet. Satt som default til: ' + tag_param)
            try:
                sync_guid = sheet.LookupParameter('AV_sync_mot_GUID').AsInteger()
                DebugPrint('Sync mot GUID: ' + str(sync_guid))
            except:
                sync_guid = False
                DebugPrint('Fant ikke parameter for sync mot GUID. Satt som default til: ' + str(sync_guid))
            try:
                AV_room_link_str = sheet.LookupParameter('AV_room_link_str').AsString()
            except:
                AV_room_link_str = ''
            DebugPrint('AV_room_link_str: ' + AV_room_link_str)
            break

    DebugPrint('Parametre fra sheet ' + str(time.time() - start))

    #Avbryt script dersom ikke IO-liste regneark er definert.
    if (har_funnet_IOliste_ark_revit == 0):
        errorReport += 'Fant ikke noe ark med navn "Kobling mot IO-liste" som angir filplassering IO-liste'
        button = UI.TaskDialogCommonButtons.None
        result = UI.TaskDialogResult.Ok
        UI.TaskDialog.Show('Synkronisering avbrutt',
                           errorReport, button)
        return

    # Filename for excel exports
    filename = Document.PathName.GetValue(doc)
    if "BIM 360://" in filename:
        sentralfil = filename
    else:
        try:
            sentralfil = BasicFileInfo.Extract(doc.PathName).CentralPath
        except:
            sentralfil = filename
    if not sentralfil:
        prosjektnavn = 'testprosjekt'
    else:
        sentralfil = sentralfil.replace('_Sentralfil','')
        sentralfil = sentralfil.replace('_sentralfil', '')
        sentralfil = sentralfil.replace(' sentralfil', '')
        sentralfil = sentralfil.replace(' Sentralfil', '')
        sentralfil = sentralfil.replace('\\', '/').replace('.', '/')
        sentralfil = sentralfil.split('/')
        prosjektnavn = sentralfil[-2]

    ### Finn riktig mappe for revit-sync
    delt_filbane = IO_liste_filplassering.Split('\\')
    s = '\\'
    mappe = s.join(delt_filbane[:-1])
    DebugPrint('Mappe :' + mappe)

    x = datetime.datetime.now()
    timestamp = x.strftime("%y%m%d %H-%M")
    DebugPrint(timestamp)
    DebugPrint(str(time.time() - start))

    #Sjekk om mappe 'Bakcup-presync' finnes
    mappe_finnes = os.path.isdir(mappe + '\Backup_presync\\')
    #lag mappe dersom ikke finnes
    if mappe_finnes == False :
        os.mkdir(mappe + '\Backup_presync\\')

    DebugPrint('Lag mappe ' + str(time.time() - start))

    #klargjør filnavn for excel-eksport. Selve eksporten kjøres helt mot slutten av script
    prosjektnavn_uten_nr = ''.join([i for i in prosjektnavn if not i.isdigit()])
    prosjektnavn_uten_nr = prosjektnavn_uten_nr[2:]
    excel_filenames = [mappe + '\Komponenter_' + prosjektnavn + '.xlsx',
                       mappe + '\Komponenter_skjema_' + prosjektnavn + '.xlsx',
                       mappe + '\Backup_presync\\' + prosjektnavn_uten_nr + '_3D_' + timestamp + '.xlsx',
                       mappe + '\Backup_presync\\' + prosjektnavn_uten_nr + '_skjema_' + timestamp + '.xlsx']

    #Liste over alle komponenter i 3d. Skal eksporters til excel og importeres til database.
    komp_3d = []
    #Første rad inneholder headers
    komp_3d.append(['element_id', 'Family', 'FamilyType', 'TAG', 'Rom', 'System Type', 'X', 'Y', 'Z', 'Tegning'])
    # Liste over alle komponenter i skjema. Skal eksporters til excel og importeres til database.
    komp_skjema = []
    #Første rad inneholder headers
    komp_skjema.append(['element_id', 'TAG', 'Family', 'FamilyType', 'Tegning', 'Komponent', 'Funksjon'])

    #"Presync" er lister over eldata slik det så ut før synkronisering. Blir sjelden brukt, men fungerer som backup dersom noen av en eller
    #annen grunn har fyllt ut masse eldata direkte i modell. Dette blir jo overskrevet av script.
    #For 3d
    presync_3d = []
    #For skjema
    presync_skjema = []

    #msg = TaskDialog

    # Finn linker, for å lage liste over alle rom
    linkCollector = Autodesk.Revit.DB.FilteredElementCollector(doc)
    linkInstances = linkCollector.OfClass(Autodesk.Revit.DB.RevitLinkInstance)

    # linkDoc, linkName, linkPath = [], [], []
    #Kode under identifiserer link som er ønsket som link som inneholder romdefinisjon (typisk ARK eller RIB)
    #Siden det er en risiko for at selve prosjektnavnet inneholder bokstavene ARK eller RIB, telles forekomst av streng
    #Link med høyest forekomst av denne strengen "vinner"
    #Eksempel der bruker har sagt at det skal brukes romdefinisjoner fra link ARK:
        #Knarkdal skole RIB: link_score: 1
        #Knarkdal skole ARK: link_score: 2
        #Knarkdal skole RIE: link_score: 1'
            # --> ARK modellen "vinner"

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
        else:
            DebugPrint("No match found for link "  + i.Name)

    DebugPrint('Loope linker ' + str(time.time() - start))
    cat_list = [BuiltInCategory.OST_Rooms]
    typed_list = List[BuiltInCategory](cat_list)
    filter = ElementMulticategoryFilter(typed_list)

    try:
        rooms = FilteredElementCollector(linkDoc).WherePasses(filter).WhereElementIsNotElementType().ToElements()
        if len(rooms) == 0:
            summaryReport += "Feil: Mangler definesjoner på rom i link. Kan ikke sende info om romplassering til database."

    except:
        # if error accurs anywhere in the process catch it
        import traceback
        #summaryReport += traceback.format_exc()
        rooms = []
        summaryReport += "Feil : Kan ikke lese romdata fra link. Kan skyldes at ARK/RIB link ikke er lastet inn, eller er i lukket workset."

    DebugPrint(' Lese inn romdata ' +  str(time.time() - start))

    #################################################
    # LES INN IO-LISTE FRA EXCEL
    #################################################
    if 'xls' in IO_liste_filplassering:
        try:
            wb_IO_liste = xl.Workbooks.Open(IO_liste_filplassering)

            # Finn sheet med IO-liste. Bruker sheet1 dersom ingen treff på navn
            try:
                ws_IO_liste = wb_IO_liste.Worksheets('IO-liste')
            except:
                try:
                    ws_IO_liste = wb_IO_liste.Worksheets('IOliste')
                except:
                    ws_IO_liste = wb_IO_liste.Worksheets[1]
                    # used = ws_IO_liste.UsedRange
            cols = ws_IO_liste.UsedRange.Columns.Count
            rows = ws_IO_liste.UsedRange.Rows.Count

            DebugPrint(' Finne excel-fil og riktig worksheet ' + str(time.time() - start))


            for i in range(1, rows+1):
                rad = []
                for j in range(0, cols):

                    try:
                        rad.append(ws_IO_liste.Range[n2a(j) + str(i)].Text)
                        #linje under fungerer kun til og med kolonne z
                        #rad.append(ws_IO_liste.Range[chr(ord('@') + j) + str(i)].Text)
                    except:
                        DebugPrint("Feil ved innlesing av IO-liste rad " + str(i) + " og kolonne " + str(j))

                IOliste.append(rad)
            #DebugPrint(IOliste)

            wb_IO_liste.Close()

            DebugPrint('Lese inn IO liste fra excel ' + str(time.time() - start))
        except:
            errorReport += "Fant ikke excel-dokument med IO-liste på plassering angitt i ark 'kobling mot IO-liste' \n Sync avbrutt"
            button = UI.TaskDialogCommonButtons.None
            result = UI.TaskDialogResult.Ok
            UI.TaskDialog.Show('Synkronisering avbrutt', errorReport, button)
            return

    elif 'csv' in IO_liste_filplassering:
        try:
            IOliste = list(csv.reader(open(IO_liste_filplassering), delimiter  =";"))
            DebugPrint('Lese inn IO liste fra csv fil ' + str(time.time() - start))
        except:
            summaryReport += "Fant ikke csv-fil med IO-liste på plassering angitt i ark 'kobling mot IO-liste' \n Sync avbrutt."
            button = UI.TaskDialogCommonButtons.None
            result = UI.TaskDialogResult.Ok
            UI.TaskDialog.Show('Synkronisering avbrutt. Fant ikke csv-fil med IO-liste på plassering angitt i ark "kobling mot IO-liste"', errorReport, button)
            return

    elif IO_liste_filplassering =="":
        errorReport +=  "Ingen fil med IO-liste angitt i ark 'kobling mot IO-liste'. \n Sync avbrutt."
        button = UI.TaskDialogCommonButtons.None
        result = UI.TaskDialogResult.Ok
        UI.TaskDialog.Show('Synkronisering avbrutt. Ingen fil med IO-liste angitt i ark "kobling mot IO-liste"', errorReport, button)
        return

    else:
        errorReport +=  "Ingen gyldig fil med IO-liste angitt i ark 'kobling mot IO-liste'. Mangler kanskje filtype i filnavn. \n Sync avbrutt"
        button = UI.TaskDialogCommonButtons.None
        result = UI.TaskDialogResult.Ok
        UI.TaskDialog.Show('Synkronisering avbrutt. Ingen gyldig fil med IO-liste angitt i ark "kobling mot IO-liste". Mangler kanskje filtype i filnavn', errorReport, button)
        return


    DebugPrint('Tag parameter: ' + str(tag_param))
    DebugPrint('sync_guid: ' + str(sync_guid))

    parametre_shared_name = ['Entreprise', 'Tekstlinje 1', 'Tekstlinje 2', 'Driftsform', 'AV_MMI',
                             'BUS-kommunikasjon',
                             'Sikkerhetsbryter', 'Kommentar', 'Fabrikat', 'Modell', 'Merkespenning', 'Merkeeffekt',
                             'Merkestrøm', 'Status', 'Access_TagType', 'Access_TagType beskrivelse', 'IO-er', 'Datablad', 'TrengerSignalKabel', 'TrengerStrømKabel']

    parametre_guid = ['2c78b93c-2c2d-4bf5-a4cd-b5ab37d40b3f', '88e7a061-da67-44a8-bdab-19fd0e111277',
                      '8bd618c5-4b04-4089-8c1c-d3329c359af7', 'da7bed97-096f-4949-840f-3125bdf40605',
                      'fd766c10-96b3-470b-aa1e-3b1b5b572492', '18f9c00a-61cf-4429-b022-311b2ce6a667',
                      'daf45c3a-35fc-4994-a29a-180483305de1', '3df6ebff-09dd-475f-8ca7-6eb31e697fd5',
                      'c4a831f8-4d5f-46a3-a40c-41f987c910f6', 'e854697d-4f77-4882-b709-72bcc27ee040',
                      'c6dde6a0-ab66-4102-9e09-30ab645e56bb', 'ed0e58dd-3954-47f7-b7eb-f561f2b55ff9',
                      '55e65789-9ba8-40fe-9795-287755774934', '57e99ba0-ecc1-43aa-8de5-60155ae1e99d',
                      '78d27e5c-b652-4ef5-b19c-767de086fe46', '2603489f-989b-4df2-b87f-0cbfbe9d4f5b',
                      '1c7f04a2-ed1a-412e-a7c1-170dde0c203c', 'bf7410c5-c78b-4463-8b39-fefedd6b4ac7',
                      '483d6e8f-fed7-4ffd-a6af-62c288325fbb', 'ee034c8e-6555-40ba-af35-0aa6244a86a3']
    for i, g in enumerate(parametre_guid):
        parametre_guid[i] = Guid(g)

    parametre_shared_name_lc = [x.lower() for x in parametre_shared_name]
    # DebugPrint('parametre_shared_name_lc ')
    # DebugPrint(parametre_shared_name_lc)

    parametre_signalinfo_lc = ['tekst', 'signaltag', 'type', 'signalkilde', 'spenning', 'tilleggstekst', 'TypeAnalogtSignal']
    parametre_ikke_sync_lc = ['sortering', 'tag', 'guid', ' tfm11fksamlet']

    # Last inn alle ubrukte parametre fra IO liste i list. Slik at de ikke trenger å defineres noe sted.
    parametre_project_name = []
    parametre_project_IO_liste_kolonne = []
    parametre_project_guid = []

    TAG_guid = '141d33b4-0f91-4dd8-a8b6-be1fa232d39f'
    TFM11FkSamlet_guid = '6b52eb8b-6935-45f9-a509-bb76724ba272'

    if tag_param.upper() == 'TAG':
        tguid = Guid(TAG_guid)
    elif tag_param == 'TFM11FkSamlet':
        tguid = Guid(TFM11FkSamlet_guid)
    elif tag_param.upper() == 'TFM':
        tguid = Guid(TFM11FkSamlet_guid)
    else:
        tguid = -1

    # finn kolonne i Io liste med tag/tfm, og finn project parametre
    tag_kol = -1

    DebugPrint(IOliste[0])

    for j, celle in enumerate(IOliste[0]):
        DebugPrint('j: ' + str(j) + ', ' + celle )
        if(1):
            # if celle.lower() == tag_param.lower():
            if celle.lower() == 'tag':  # Bruker tag her og ikke tag_param.lower(), siden det alltid er TAG som brukes i eksport fra database.
                # Kan evt. splitte dette opp i to forskjellige parametre i ark med parametre.
                tag_kol = j
            if celle.lower() == 'guid':
                GUID_kol = j
            if celle.lower() not in parametre_signalinfo_lc and celle.lower() not in parametre_ikke_sync_lc:
                parametre_project_name.append(celle)
                parametre_project_IO_liste_kolonne.append(j)
                if celle.lower() in parametre_shared_name_lc:
                    p_index = parametre_shared_name_lc.index(celle.lower())
                    parametre_project_guid.append(parametre_guid[p_index])
                else:
                    parametre_project_guid.append(None)
                DebugPrint('project param lagt til: ' + celle + '. Kolonne: ' + str(j))
                #DebugPrint('IO liste headers, j, celle, try: ' + str(j) + ' ' + celle)

        #except:
        if(0):
            DebugPrint('IO liste headers, j, celle, continue: ' + str(j) + ' ' + celle)
            # Tidligere stod det break her. Tror continue er bedre.
            continue
    #Resultat av kode over: tre lister for parametre i prosjektet.
    # parametre_project_name
    # parametre_project_IO_liste_kolonne
    # parametre_project_guid    (denne er satt til "None") for parametre som ikke er med i standard parameteromfang)
    # Disse blir brukt ved looping av categories

    # sjekk om det ble funnet kolonne med tag/tfm
    if tag_kol == (-1):
        # tag_kol not found
        errorReport += 'Fant ikke kolonne ved navn ' + tag_param + ' i IO liste excel-ark. Sync avbrutt'
        ## må legge til en avbryt-funksjon her
        button = UI.TaskDialogCommonButtons.None
        result = UI.TaskDialogResult.Ok
        UI.TaskDialog.Show('Synkronisering avbrutt. Fant ikke kolonne med TAG', errorReport, button)
        return
    DebugPrint('tag_kol: ' + str(tag_kol))
    DebugPrint('Behandle parameternavn i første rad IO liste '+ str(time.time() - start))

    cat_list = [BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_GenericModel, BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_DetailComponents, BuiltInCategory.OST_SecurityDevices]

    transaction = DB.Transaction(doc)
    transaction.Start("Sync eldata")

    # loop all categories
    for cat in cat_list:
        DebugPrint(str(time.time() - start))
        DebugPrint(cat)

        # sjekk om tag/tfm parameter finnes, og om den er shared, og om den er definert med samme GUID som den riktige shared parameteren
        # OBS: Finnes en svakhet med script. Kan risikere at det er litt tilfeldig om finner parametre siden det går ann å ha parametre som er definert i family og ikke som category parameter
        # Dette gjelder spesielt for Detail Items der parameter TAG tidligere kun var definert i family og ikke som category parameter. Dette er riktig i nyere malfiler.
        tag_cat_status = (-1)

        # tag_cat_status -1 betyr at mangler tag/tfm parameter for category. Skipper category
        # tag_cat_status 0 betyr at den finnes, men at man ikke kan bruke GUID (enten fordi tag_param er ulik både tfm og tag, eller fordi det ikke er brukt shared parameter for denne)
        # tag_cat_status 1 betyr at den finnes som shared parameter, og at definisjonen stemmer overens med offisiel AV standard.
        # tag_cat_status 2 betyr at aktuell komponent er skjemakomponent, og man bruker TFM: SystemVar + '-' + TFM11FkKompGruppe + TFM11FkKompLNR

        # lag lister over parametre som finnes for category
        p_s_IO_cat_kol = []  # parametre_shared_IO_liste_category_kolonne
        p_s_IO_cat_name = []
        p_s_IO_cat_guid = []
        p_r_IO_cat_kol = []
        p_r_IO_cat_name = []

        # bruker try her for å unngå feil for categorier som ikke er i bruk, og dermed ikke har firstElement
        # find all parameters defined for the category
        try:
            catel = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().FirstElement()
            if catel is not None:
                for param in catel.Parameters:
                    if param.Definition.Name in parametre_project_name:
                        i = parametre_project_name.index(param.Definition.Name)
                        if param.IsShared == True and param.GUID == parametre_project_guid[i]:
                            p_s_IO_cat_kol.append(parametre_project_IO_liste_kolonne[i])    # s = shared
                            p_s_IO_cat_name.append(parametre_project_name[i])
                            p_s_IO_cat_guid.append(parametre_project_guid[i])
                        else:
                            p_r_IO_cat_kol.append(parametre_project_IO_liste_kolonne[i])  # r = remaining (dvs. not shared)
                            p_r_IO_cat_name.append(parametre_project_name[i])
                    if param.IsShared == True and param.GUID == tguid:
                        tag_cat_status = 1  # status 1 betyr at den finnes som shared parameter, og at definisjonen stemmer overens med offisiel AV standard.
                    elif param.Definition.Name == tag_param:
                        if tag_cat_status == -1:
                            tag_cat_status = 0  # status 0 betyr at den finnes, men at man ikke kan bruke GUID (enten fordi tag_param er ulik både tfm og tag, eller fordi det ikke er brukt shared parameter for denne)
            else:
                DebugPrint('Ingen elementer i cat. Ingen parameter testing')
                continue
        except:
            # bør legge inn en advarsel her
            DebugPrint('failed parameter testing')
            continue

        DebugPrint('Size p_s_IO_cat_name: ' + str(len(p_s_IO_cat_name)))
        DebugPrint(p_s_IO_cat_name)
        DebugPrint(p_s_IO_cat_kol)
        DebugPrint('Size p_r_IO_cat_name: ' + str(len(p_r_IO_cat_name)))
        DebugPrint(p_r_IO_cat_name)
        DebugPrint(p_r_IO_cat_kol)
        # sorting params
        #p_s_IO_cat_name = [x for (y, x) in sorted(zip(p_s_IO_cat_kol, p_s_IO_cat_name), key=lambda pair: pair[0])]
        #p_s_IO_cat_guid = [x for (y, x) in sorted(zip(p_s_IO_cat_kol, p_s_IO_cat_guid), key=lambda pair: pair[0])]
        #p_s_IO_cat_kol = sorted(p_s_IO_cat_kol)

        #if tag_cat_status == -1 and cat == BuiltInCategory.OST_DetailComponents:
        if cat == BuiltInCategory.OST_DetailComponents:
            if tag_param == 'TFM11FkSamlet' or tag_param == 'TFM':
                 # summaryReport += "Advarsel. Skyldes trolig at tag-parameter ikke er lagt til som project parameter for Detail Item. \n"
                tag_cat_status = 2  # Betyr at man bruker sammenslått streng som tag/TFM

        DebugPrint('tag_cat_status: ' + str(tag_cat_status))
        if tag_cat_status == (-1):
            # category mangler definert tag-parameter
            DebugPrint('category mangler definert tag-parameter')
            # går til neste category
            continue

        if cat == BuiltInCategory.OST_DetailComponents:
            #Lag liste over alle detail item tags. Brukes senere for å kryssjekke om detail item er tagget på tegning
            c_list = [DB.BuiltInCategory.OST_DetailComponentTags]
            t_list = List[DB.BuiltInCategory](c_list)
            filt = DB.ElementMulticategoryFilter(t_list)
            DItags= DB.FilteredElementCollector(doc).WherePasses(filt).WhereElementIsNotElementType().ToElements()
            DetailItemTags = []
            DetailTtemTagObjects = []
            for i in DItags:
                try:
                    DetailItemTags.append(i.TagText)
                    DetailTtemTagObjects.append(i)
                except :
                    pass

        # loop elements in category
        EQ = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
        n_elements = 0

        DebugPrint('parameteromfang klarert')
        DebugPrint(str(time.time() - start))

        for k in EQ:
            # Tag reset
            tag = ''

            # sjekk om tag/tfm er blank, og finn tekstverdi på tag i revit
            if tag_cat_status == 1:
                try:
                    if k.get_Parameter(tguid).AsString() == 'null' or k.get_Parameter(
                            tguid).AsString() == '=-' or k.get_Parameter(tguid).AsString() == '' or k.get_Parameter(
                        tguid).AsString() is None:
                        # gå til neste element dersom blank tag/tfm
                        SummaryPrint('blank tag/tfm (shared param):')
                        continue
                    tag = k.get_Parameter(tguid).AsString()
                    # DebugPrint('k.get_Parameter(tguid).AsString() : ' + k.get_Parameter(tguid).AsString())
                    # DebugPrint('k.LookupParameter(TAG).AsString() : ' + k.LookupParameter('TAG').AsString())
                except:

                    SummaryPrint("Feil. Skyldes trolig at parameter TAG ikke er lagt til som project parameter for Detail Item. Skipper element")
                    continue
            elif tag_cat_status == 0:
                if k.LookupParameter(tag_param).AsString() == 'null' or k.LookupParameter(
                        tag_param).AsString() == '=-' or k.LookupParameter(tag_param).AsString() == '' or k.LookupParameter(
                    tag_param).AsString() is None:
                    # gå til neste element dersom blank tag/tfm
                    SummaryPrint('blank tag/tfm (project param) for:')
                    continue
                tag = k.LookupParameter(tag_param).AsString()
            elif tag_cat_status == 2:
                #DebugPrint('tag_cat_status 2 start ')
                # skjema TFM
                try:
                    #SystemVar = i.LookupParameter('SystemVar').AsString()
                    #TFM11FkKompLNR = i.LookupParameter('TFM11FkKompLNR').AsString()
                    #TFM11FkKompGruppe = i.Symbol.LookupParameter('TFM11FkKompGruppe').AsString()
                    #tag = r"'" + '=' + SystemVar + '-' + TFM11FkKompGruppe + TFM11FkKompLNR
                    #DebugPrint('tag_cat_status 2 try start')
                    TFMkode = k.LookupParameter('TFM-kode').AsString()
                    #DebugPrint('tag_cat_status 2 try 2')
                    Systemnummer = k.LookupParameter('Systemnummer').AsString()
                    #Bør trolig legge inn sjekk her mot tomme verdier
                    #DebugPrint('tag_cat_status 2 ')
                    if TFMkode !='' and Systemnummer != '' :
                        tag = Systemnummer + '-' + TFMkode
                        #DebugPrint('Sammenslått tag: ' + tag)
                    else:
                        #DebugPrint('Klarte ikke lage sammenslått tag')
                        continue
                except:
                #else :
                    SummaryPrint('feil ved sammenslåing/avlesing av parametre som inngår i TFM for skjema')
                    DebugPrint('feil ved sammenslåing/avlesing av parametre som inngår i TFM for skjema')
                    errorReport += 'feil ved sammenslåing/avlesing av parametre som inngår i TFM for skjema'
                    continue
            DebugPrint('Tag: ' + tag)

            n_elements += 1

            #############################################################################################################
            # SYNC DATA TIL REVIT. Inkl. eksport presync
            #############################################################################################################

            #####Finn matchende rad i IO-liste
            IO_liste_row = (-1)
            #Sjekk om sync mot GUID
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
            DebugPrint('IO_liste_row :' + str(IO_liste_row))

            # Presync data
            # lag tom header list for denne category (antall parametre kan variere mellom categories)
            if n_elements == 1:
                presync_top_row = ['TAG']
            # lag tom list til presync-data for dette elementet
            if cat <> BuiltInCategory.OST_DetailComponents:
                SummaryPrint('Ny rad presync - 3d. Tag: ' + tag)
                presync_3d_row = [tag]
            else:
                SummaryPrint('Ny rad presync - skjema. Tag: ' + tag)
                presync_skjema_row = [tag]

            # oppdater_eldata(IO_liste_row, k)
            try:
                DebugPrint('presync_top_row')
                DebugPrint(presync_top_row)
                DebugPrint('IO_liste_row: ' + str(IO_liste_row) + ', n_elements: ' + str(n_elements))
                OppdaterEldata(cat, IO_liste_row, k, n_elements, p_s_IO_cat_kol, p_s_IO_cat_guid, p_s_IO_cat_name,
                               p_r_IO_cat_name, p_r_IO_cat_kol)
            except:
                DebugPrint("feil for rad:" + str(IO_liste_row))

            if cat <> BuiltInCategory.OST_DetailComponents:
                if n_elements == 1:
                    presync_3d.append(presync_top_row)
                presync_3d.append(presync_3d_row)
            else:
                DebugPrint('skjemaleement presync')
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

            # komp_skjema
            else:
                # Finn family
                DebugPrint('skjemaelement')
                #DebugPrint(tag)
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
                # Finn verdier av felter "Komponent" og "Funksjon". Tekstfelter med metadata som VVS bruker
                try:
                    funksjon = k.LookupParameter('Funksjon').AsString()
                except:
                    funksjon = ''
                try:
                    komponent = k.LookupParameter('Komponent').AsString()
                except:
                    komponent = ''

                # Sjekk om tag_label er synlig. Kan være vist enten som label innebygget i detail item, eller som tag by catebory
                # Dersom ikke synlig tag: Stor sannsynlighet for at feil tag. Fjernes derfor fra eksport.
                try:
                    tag_label = k.LookupParameter('Tag label').AsInteger()
                    DebugPrint('tag_label: ' + str(tag_label))
                except:
                    DebugPrint('tag_label undefined')
                    tag_label = 0
                finnes_tagget = 0
                #Sjekker om detail item er vist på tegning som plottes, dvs tegning med tegningsnummer. Dersom ikke, sannsynligvis kok, eller uferdig. Tas ikke med på eksport.
                #try:
                if(1):

                    sheetelem = doc.GetElement(k.OwnerViewId)
                    sheetparameter = sheetelem.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER)
                    skjemanr = sheetparameter.AsString()
                    #DetailTtemTagObjects

                    DebugPrint('skjema: ' + skjemanr)

                    # summaryReport = summaryReport + (' \n tag_label: ' + str(tag_label))
                    if tag_label == 1 or tag in DetailItemTags:
                        komp_skjema.append([k.Id, tag, family, familytype, '', komponent, funksjon])
                        DebugPrint ('Lagt til fordi full tag label')
                    elif TFMkode in DetailItemTags:
                        taglabelindex = DetailItemTags.index(TFMkode)
                        taglabelindex = [j for j, y in enumerate(DetailItemTags) if y == TFMkode]
                        DebugPrint('Treff på komponentkode:')
                        DebugPrint(taglabelindex)
                        for m in taglabelindex:
                            sheetelemtag = doc.GetElement(DetailTtemTagObjects[m].OwnerViewId)
                            sheetparametertag = sheetelemtag.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER)
                            skjemanrtag = sheetparametertag.AsString()
                            #  antar at dersom detail item er tagget på samme tegning som detail item er vist, og TFM-kode er lik så er det det objektet som er tagget.
                            #  Kan ikke vær e100% sikker siden systemnr ikke er vist på tegning
                            DebugPrint('skjemanrtag :' +skjemanrtag)
                            if skjemanrtag == skjemanr:
                                komp_skjema.append([k.Id, tag, family, familytype, '', komponent, funksjon])
                                break
                #except:
                else:
                    # Blir ikke med på eksport siden ikke vist på tegning
                    DebugPrint('detail item ikke med på eksport siden ikke vist på tegning')

        DebugPrint('n_elements: ' + str(n_elements))

    #########################################################################################################################################
    # Loop tegninger
    ##############################################################################################################################################

    # Liste/collector over alle viewports
    paramvalue = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER))
    ruleval = DB.FilterStringGreater()
    #Old code. Last parameter is case-sensitivity:
    #filterrule = DB.FilterStringRule(paramvalue, ruleval, "", False)
    #New code, 2023 compatible
    try:
        filterrule = DB.FilterStringRule(paramvalue, ruleval, "")
    except:
        filterrule = DB.FilterStringRule(paramvalue, ruleval, "", False)

    paramfi = DB.ElementParameterFilter(filterrule)
    viewscollector = DB.FilteredElementCollector(doc).OfClass(DB.View).WherePasses(paramfi)

    for v in viewscollector:
        viewelemcollector = DB.FilteredElementCollector(doc, v.Id).ToElementIds()
        vparam = v.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER)  # type: DB.Parameter
        # loop komp 3d
        #DebugPrint('Tegning: ' + vparam.AsString())
        for n, e in enumerate(komp_3d):
            if n == 0:
                # må hoppe over tabell headers
                continue
            # DebugPrint('Tag : ' + e[2])
            if e[0] in viewelemcollector:
                #DebugPrint('Treff ' + vparam.AsString() + ' - ' + e[2])
                if komp_3d[n][9] == '':
                    komp_3d[n][9] = vparam.AsString()
                else:
                    komp_3d[n][9] += ', ' + vparam.AsString()
        # loop komp skjema
        for n, e in enumerate(komp_skjema):
            if n == 0:
                # må hoppe over tabell headers
                continue
            if e[0] in viewelemcollector:
                #DebugPrint('Treff ' + vparam.AsString() + ' - ' + e[1])
                if komp_skjema[n][4] == '':
                    komp_skjema[n][4] = vparam.AsString()
                else:
                    komp_skjema[n][4] += ', ' + vparam.AsString()
    #####
    #Slå samme duplikater flytskjemasymboler
    #####

    finnes = []
    komp_skjema_ny = []
    for n, e in enumerate(komp_skjema):
        if n == 0:
            # tabell headers
            finnes.append(komp_skjema[n][1])
            komp_skjema_ny.append(e)
            continue
        if e[1] in finnes:
            i = finnes.index(e[1])
            #komp_skjema.append([k.Id, tag, family, familytype, '', komponent, funksjon])
            #komp_skjema_ny[i][0] += ', ' + e[0]      #element_id
            komp_skjema_ny[i][2] += ', ' + e[2]       #Family
            komp_skjema_ny[i][3] += ', ' + e[3]       #FamilyType
            komp_skjema_ny[i][4] += ', ' + e[4]       #Tegning            DETTE ER DEN VIKTIGSTE HER
            komp_skjema_ny[i][5] += ', ' + e[5]       #Komponent
            komp_skjema_ny[i][6] += ', ' + e[6]       #Funksjon

        else:
            finnes.append(komp_skjema[n][1])
            komp_skjema_ny.append(e)

    komp_skjema = komp_skjema_ny

    #####################################################
    ############Eksporter til excel
    #####################################################

    excel_eksport = [komp_3d, komp_skjema, presync_3d, presync_skjema]

    DebugPrint('presync skjema')
    DebugPrint(presync_skjema)
    DebugPrint(' Før excel-eksport ' +str(time.time() - start))

    for i in range(4):
        if excel_eksport[i] <> []:
            if SaveListToExcel(excel_filenames[i], excel_eksport[i]):
                DebugPrint('Fileksport success ' + str(i))
            else:
                DebugPrint('Fileksport failed ' + str(i))
                DebugPrint(excel_filenames[i])
                summaryReport += 'Kunne ikke eksportere fil ' + excel_filenames[i] + '. Sjekk om fil allerede er åpen, og lukk den, og prøv på ny.'

    xl.DisplayAlerts = True
    xl.Visible = True
    DebugPrint(' Etter excel-eksport, før transaction ' + str(time.time() - start))
    transaction.Commit()
    DebugPrint(' Etter transaction ' + str(time.time() - start))
    button = UI.TaskDialogCommonButtons.None
    result = UI.TaskDialogResult.Ok
    if summaryReport == "":
        summaryReport = 'Synkronisering gjennomført uten feil'
    else:
        summaryReport =  'Synkronisering ferdig.\n\n' + summaryReport
    UI.TaskDialog.Show('Ferdig', summaryReport, button)
    DebugPrint(str(time.time() - start))
    return

presync_top_row = []
presync_3d_row = []
presync_skjema_row = []
IOliste = []

MainFunction()
