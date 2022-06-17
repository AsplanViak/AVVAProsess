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
__title__ = 'Synkroniser postnr'  # Denne overstyrer navnet på scriptfilen
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
        #if n_elements == 1:
            # DebugPrint('presync header')
        #    presync_top_row.append(p_s_IO_cat_name[i])
        # presync data
        #try:
            # DebugPrint('Check if detail item')
         #   if cat <> BuiltInCategory.OST_DetailComponents:
                # DebugPrint('Is not detail item')
          #      presync_3d_row.append(k.get_Parameter(p_s_IO_cat_guid[i]).AsString())
           # else:
                # DebugPrint('Is detail item')
            #    presync_skjema_row.append(k.get_Parameter(p_s_IO_cat_guid[i]).AsString())
        #except:
         #   if cat <> BuiltInCategory.OST_DetailComponents:
          #      presync_3d_row.append('')
          #  else:
           #     presync_skjema_row.append('')
        if IO_liste_row != (-1):
            # sync
            # DebugPrint('Syncing')
            IOliste_tekst = IOliste[IO_liste_row][kol]
            if IOliste_tekst is None:
                IOliste_tekst = ' '
            # DebugPrint('IOliste_tekst: ' + str(IOliste_tekst))
            try:
                #rad under er den som utfører selve endringen av parameter-verdi i Revit
                ############SKAL TROLIG STÅ kol ISTEDET FOR i PÅ RAD UNDER!!!!!!!!!!!!!!!!!!!!!

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
                AV_room_link_str = 'RIB'
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
        sentralfil = BasicFileInfo.Extract(doc.PathName).CentralPath
    if not sentralfil:
        prosjektnavn = 'testprosjekt'
    else:
        sentralfil = sentralfil.replace('_Sentralfil','')
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
    #excel_filenames = [mappe + '\Komponenter_' + prosjektnavn + '.xlsx',
     #                  mappe + '\Komponenter_skjema_' + prosjektnavn + '.xlsx',
      #                 mappe + '\Backup_presync\\' + prosjektnavn_uten_nr + '_3D_' + timestamp + '.xlsx',
      #                 mappe + '\Backup_presync\\' + prosjektnavn_uten_nr + '_skjema_' + timestamp + '.xlsx']

    #Liste over alle komponenter i 3d. Skal eksporters til excel og importeres til database.
    komp_3d = []
    #Første rad inneholder headers
    komp_3d.append(['element_id', 'Family', 'FamilyType', 'TAG', 'Rom', 'System Type', 'X', 'Y', 'Z', 'Tegning'])
    # Liste over alle komponenter i skjema. Skal eksporters til excel og importeres til database.
    komp_skjema = []
    #Første rad inneholder headers
    komp_skjema.append(['element_id', 'TAG', 'Family', 'FamilyType', 'Tegning'])

    #"Presync" er lister over eldata slik det så ut før synkronisering. Blir sjelden brukt, men fungerer som backup dersom noen av en eller
    #annen grunn har fyllt ut masse eldata direkte i modell. Dette blir jo overskrevet av script.
    #For 3d
    presync_3d = []
    #For skjema
    presync_skjema = []

    #msg = TaskDialog

    # Finn linker, for å lage liste over alle rom
    #linkCollector = Autodesk.Revit.DB.FilteredElementCollector(doc)
    #linkInstances = linkCollector.OfClass(Autodesk.Revit.DB.RevitLinkInstance)

    # linkDoc, linkName, linkPath = [], [], []
    #Kode under identifiserer link som er ønsket som link som inneholder romdefinisjon (typisk ARK eller RIB)
    #Siden det er en risiko for at selve prosjektnavnet inneholder bokstavene ARK eller RIB, telles forekomst av streng
    #Link med høyest forekomst av denne strengen "vinner"
    #Eksempel der bruker har sagt at det skal brukes romdefinisjoner fra link ARK:
        #Knarkdal skole RIB: link_score: 1
        #Knarkdal skole ARK: link_score: 2
        #Knarkdal skole RIE: link_score: 1'
            # --> ARK modellen "vinner"


    #################################################
    # LES INN IO-LISTE FRA EXCEL
    #################################################
    if 'xls' in IO_liste_filplassering:
        try:
            wb_IO_liste = xl.Workbooks.Open(IO_liste_filplassering)
        except:
            errorReport += 'Fant ikke excel-dokument med IO-liste på plassering angitt i ark "kobling mot IO-liste" '
            button = UI.TaskDialogCommonButtons.None
            result = UI.TaskDialogResult.Ok
            UI.TaskDialog.Show('Synkronisering avbrutt', errorReport, button)

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

    elif 'csv' in IO_liste_filplassering:
        IOliste = list(csv.reader(open(IO_liste_filplassering), delimiter  =";"))
        DebugPrint('Lese inn IO liste fra csv fil ' + str(time.time() - start))


    DebugPrint('Tag parameter: ' + str(tag_param))
    DebugPrint('sync_guid: ' + str(sync_guid))

    parametre_shared_name = ['Entreprise', 'Tekstlinje 1', 'Tekstlinje 2', 'Driftsform', 'AV_MMI',
                             'BUS-kommunikasjon',
                             'Sikkerhetsbryter', 'Kommentar', 'Fabrikat', 'Modell', 'Merkespenning', 'Merkeeffekt',
                             'Merkestrøm', 'Status', 'Access_TagType', 'Access_TagType beskrivelse', 'IO-er', 'Datablad']

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
        UI.TaskDialog.Show('Synkronisering avbrutt', errorReport, button)
        return
    DebugPrint('tag_kol: ' + str(tag_kol))
    DebugPrint('Behandle parameternavn i første rad IO liste '+ str(time.time() - start))

    #cat_list = [BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment,
        #        BuiltInCategory.OST_GenericModel, BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_Sprinklers,
        #        BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_DuctTerminal,
        #        BuiltInCategory.OST_DetailComponents]
    #cat_list = [BuiltInCategory.OST_PipeFitting,BuiltInCategory.OST_PipeCurves]
    cat_list = [BuiltInCategory.OST_PipeFitting,BuiltInCategory.OST_PipeCurves,BuiltInCategory.OST_PipingSystem]

    # BuiltInCategory.OST_PipeSegments,

    transaction = DB.Transaction(doc)
    transaction.Start("Sync eldata")

    # loop all categories
    for cat in cat_list:
        DebugPrint(str(time.time() - start))
        DebugPrint(cat)

        # sjekk om tag/tfm parameter finnes, og om den er shared, og om den er definert med samme GUID som den riktige shared parameteren
        tag_cat_status = (-1)

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
                            p_s_IO_cat_kol.append(parametre_project_IO_liste_kolonne[i])  # s = shared
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
        # p_s_IO_cat_name = [x for (y, x) in sorted(zip(p_s_IO_cat_kol, p_s_IO_cat_name), key=lambda pair: pair[0])]
        # p_s_IO_cat_guid = [x for (y, x) in sorted(zip(p_s_IO_cat_kol, p_s_IO_cat_guid), key=lambda pair: pair[0])]
        # p_s_IO_cat_kol = sorted(p_s_IO_cat_kol)

        if tag_cat_status == -1 and cat == BuiltInCategory.OST_DetailComponents:
            if tag_param == 'TFM11FkSamlet' or tag_param == 'TFM':
                tag_cat_status = 2  # Betyr at man bruker denne som tag: SystemVar + '-' + TFM11FkKompGruppe + TFM11FkKompLNR

        DebugPrint('tag_cat_status: ' + str(tag_cat_status))
        if tag_cat_status == (-1):
            # category mangler definert tag-parameter
            DebugPrint('category mangler definert tag-parameter')
            # går til neste category
            #continue

        # loop elements in category
        EQ = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
        n_elements = 0

        DebugPrint('parameteromfang klarert')
        DebugPrint(str(time.time() - start))

        for k in EQ:
            # Tag reset
            #tag = k.LookupParameter("System Type").AsString()
            if 1:
            #try:
                if(cat==BuiltInCategory.OST_PipingSystem):
                    tag = k.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
                else:
                    #'llinje under på fikses-------------------------------------------------------------------------------------------'
                    #tag = k.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
                    tag = k.Name
            #except:
            else:
                DebugPrint('Finner ikke system type. Skipper til neste')
                continue
            DebugPrint('tag/systemtype: ' + tag)

            IO_liste_row = -1
            # tag/tfm-sync
            for b in range(1, len(IOliste)):
                # if k.LookupParameter('TAG').AsString() == IOliste[l][tag_kol]:
                #tag = k.LookupParameter(tag_param).AsString()

                #bruker "in" siden tag/system type har ekstra benevnelser som PN16 etc.
                if IOliste[b][tag_kol] in tag:

                    IO_liste_row = b
                    break
            #DebugPrint('IO_liste_row :' + str(IO_liste_row))
            if IO_liste_row <> -1:
                OppdaterEldata(cat, IO_liste_row, k, 3, p_s_IO_cat_kol, p_s_IO_cat_guid, p_s_IO_cat_name,
                           p_r_IO_cat_name, p_r_IO_cat_kol)
            else:
                DebugPrint('Fant ikke rad i IO-liste for system type: ' + tag)


        DebugPrint('n_elements: ' + str(n_elements))

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
    UI.TaskDialog.Show('Synkronisering eldata ferdig', summaryReport, button)
    DebugPrint(str(time.time() - start))
    return

presync_top_row = []
presync_3d_row = []
presync_skjema_row = []
IOliste = []

MainFunction()
