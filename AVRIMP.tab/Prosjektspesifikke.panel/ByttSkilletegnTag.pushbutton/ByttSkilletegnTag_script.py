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
__title__ = 'Bytt Skilletegn Tag'  # Denne overstyrer navnet på scriptfilen
__author__ = 'Asplan Viak'  # Dette kommer opp som navnet på utvikler av knappen
__doc__ = "Gir mulighet for å endre skilletegn på alle TAG i flytskjema og 3D objekter."  # Dette blir hjelp teksten som kommer opp når man holder musepekeren over knappen.
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

from Autodesk.Revit.DB.Plumbing import *

from System.Collections.Generic import List

import sys
import math
import os
import time
import csv

#clr.AddReferenceByName(
#    'Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
#from Microsoft.Office.Interop import Excel

from System.Collections.Generic import List
from System import Guid
from System import Array

import Autodesk
from Autodesk.Revit.UI import TaskDialog
import datetime
import string

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Application, Form, Label, TextBox, Button, DockStyle, DialogResult

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

'''#definer excel-applikasjon
xl = Excel.ApplicationClass()
xl.Visible = False
xl.DisplayAlerts = False
'''

#funksjon for å returnere kolonnenavn for excel.
def n2a(n,b=string.ascii_uppercase):
   d, m = divmod(n,len(b))
   return n2a(d-1,b)+b[m] if d else b[m]


class InputFormGammel(Form):
    def __init__(self):
        self.Text = "Gammelt skilletegn tag"
        self.label = Label(Text="Skriv inn gammelt skilletegn:")
        self.label.Dock = DockStyle.Top

        self.input_box = TextBox()
        self.input_box.Dock = DockStyle.Top

        self.ok_button = Button(Text="OK", DialogResult=DialogResult.OK)
        self.ok_button.Dock = DockStyle.Top
        self.ok_button.Click += self.ok_button_click

        self.cancel_button = Button(Text="Cancel", DialogResult=DialogResult.Cancel)
        self.cancel_button.Dock = DockStyle.Top

        self.Controls.Add(self.cancel_button)
        self.Controls.Add(self.ok_button)
        self.Controls.Add(self.input_box)
        self.Controls.Add(self.label)

    def ok_button_click(self, sender, event):
        self.Close()

class InputFormNytt(Form):
    def __init__(self):
        self.Text = "Nytt skilletegn tag"
        self.label = Label(Text="Skriv inn nytt skilletegn:")
        self.label.Dock = DockStyle.Top

        self.input_box = TextBox()
        self.input_box.Dock = DockStyle.Top

        self.ok_button = Button(Text="OK", DialogResult=DialogResult.OK)
        self.ok_button.Dock = DockStyle.Top
        self.ok_button.Click += self.ok_button_click

        self.cancel_button = Button(Text="Cancel", DialogResult=DialogResult.Cancel)
        self.cancel_button.Dock = DockStyle.Top

        self.Controls.Add(self.cancel_button)
        self.Controls.Add(self.ok_button)
        self.Controls.Add(self.input_box)
        self.Controls.Add(self.label)
    def ok_button_click(self, sender, event):
        self.Close()

def MainFunction():

    global IOliste
    #gammelt_skilletegn = input("Skriv inn gammelt skilletegn: ")
    #nytt_skilletegn = input("Skriv inn nytt skilletegn: ")

    # Create and run the application
    form = InputFormGammel()
    result = form.ShowDialog()

    # Check if the OK button was clicked
    if result == DialogResult.OK:
        gammelt_skilletegn = form.input_box.Text
        DebugPrint("G You entered:"+ gammelt_skilletegn)
    else:
        DebugPrint("User canceled the input.")

    # Create and run the application
    form = InputFormNytt()
    result = form.ShowDialog()

    # Check if the OK button was clicked
    if result == DialogResult.OK:
        nytt_skilletegn = form.input_box.Text
        DebugPrint("Nytt You entered:"+ nytt_skilletegn)
    else:
        DebugPrint("User canceled the input.")

    #kritiske feil:
    errorReport = ""
    #ikke-kritiske feil og generell info:
    summaryReport = ""
    start = time.time()

    TAG_guid = '141d33b4-0f91-4dd8-a8b6-be1fa232d39f'
    TFM11FkSamlet_guid = '6b52eb8b-6935-45f9-a509-bb76724ba272'


    #tag_param = input("Skriv inn navn på parameter for TAG, f.eks. TAG eller TFM11FkSamlet")
    tag_param = 'TAG'

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

    cat_list = [BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment,
               BuiltInCategory.OST_GenericModel, BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_DetailComponents]

    transaction = DB.Transaction(doc)
    transaction.Start("Bytt skilletegn")

    # loop all categories
    for cat in cat_list:
        DebugPrint(str(time.time() - start))
        DebugPrint(cat)

        # sjekk om tag/tfm parameter finnes, og om den er shared, og om den er definert med samme GUID som den riktige shared parameteren
        tag_cat_status = (-1)

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

        if tag_cat_status == -1 and cat == BuiltInCategory.OST_DetailComponents:
            if tag_param == 'TFM11FkSamlet' or tag_param == 'TFM':
                # summaryReport += "Advarsel. Skyldes trolig at tag-parameter ikke er lagt til som project parameter for Detail Item. \n"
                tag_cat_status = 2  # Betyr at man bruker denne som tag: SystemVar + '-' + TFM11FkKompGruppe + TFM11FkKompLNR

        DebugPrint('tag_cat_status: ' + str(tag_cat_status))
        if tag_cat_status == (-1):
            # category mangler definert tag-parameter
            DebugPrint('category mangler definert tag-parameter')
            # går til neste category
            continue

            # Ny kode starter her

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

                    # Kode for å endre tag her
                    res = k.get_Parameter(tguid).Set(tag.replace(gammelt_skilletegn, nytt_skilletegn))

                    # DebugPrint('k.get_Parameter(tguid).AsString() : ' + k.get_Parameter(tguid).AsString())
                    # DebugPrint('k.LookupParameter(TAG).AsString() : ' + k.LookupParameter('TAG').AsString())
                except:
                    SummaryPrint(
                        "Feil. Skyldes trolig at parameter TAG ikke er lagt til som project parameter for Detail Item. Skipper element")
                    continue
            elif tag_cat_status == 0:
                if k.LookupParameter(tag_param).AsString() == 'null' or k.LookupParameter(
                        tag_param).AsString() == '=-' or k.LookupParameter(
                    tag_param).AsString() == '' or k.LookupParameter(
                    tag_param).AsString() is None:
                    # gå til neste element dersom blank tag/tfm
                    SummaryPrint('blank tag/tfm (project param) for:')
                    continue
                tag = k.LookupParameter(tag_param).AsString()

                #Kode for å endre tag her
                res = k.LookupParameter(tag_param).Set(tag.replace(gammelt_skilletegn, nytt_skilletegn))
                if (res):
                    SummaryPrint('remaining parameter : ok')
                else:
                    SummaryPrint('remaining parameter : feil')

            elif tag_cat_status == 2:
                # skjema TFM
                try:
                    SystemVar = i.LookupParameter('SystemVar').AsString()
                    TFM11FkKompLNR = i.LookupParameter('TFM11FkKompLNR').AsString()
                    TFM11FkKompGruppe = i.Symbol.LookupParameter('TFM11FkKompGruppe').AsString()
                    tag = r"'" + '=' + SystemVar + '-' + TFM11FkKompGruppe + TFM11FkKompLNR
                    # Bør trolig legge inn sjekk her mot tomme verdier
                except:
                    SummaryPrint('feil ved sammenslåing/avlesing av parametre som inngår i TFM for skjema')
                    errorReport += 'feil ved sammenslåing/avlesing av parametre som inngår i TFM for skjema'
                    continue
            DebugPrint('Tag: ' + tag)

            n_elements += 1

    #xl.DisplayAlerts = True
    #xl.Visible = True
    DebugPrint(' Etter excel-eksport, før transaction ' + str(time.time() - start))
    transaction.Commit()
    DebugPrint(' Etter transaction ' + str(time.time() - start))
    button = UI.TaskDialogCommonButtons.None
    result = UI.TaskDialogResult.Ok

    UI.TaskDialog.Show('Bytt skilletegn ferdig', summaryReport, button)
    DebugPrint(str(time.time() - start))
    return

MainFunction()
