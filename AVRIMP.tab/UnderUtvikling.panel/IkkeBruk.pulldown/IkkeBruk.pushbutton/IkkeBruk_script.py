# -*- coding: Windows-1252 -*-
#  Headeren over m� du ha om scriptet inneholder ���.

"""
Du trenger:
IronPython innstallert p� din PC.
pyCharm(python editor).
Fungerende autocomplete i pycharm.
    Under pyCharm settings:
    Project Interpreter:
    Sett opp pycharm med IronPython som din project enterpreter(ikke virtuel)
    Sett disse mappene til Source
    Project Structure:
    Sett disse mappene til sources: pyrevitlib, site-packages, dinlokalesharepoint\Asplan Viak\AVTools - Dokumenter\AVToolsRammeverk\AutocompleteStubs

"""

# Start M� ha
__title__ = 'IkkeBruk'  # Denne overstyrer navnet p� scriptfilen
__author__ = 'Asplan Viak'  # Dette kommer opp som navnet p� utvikler av knappen
__doc__ = "Klikk for � legge til flenser i prosjektet."  # Dette blir hjelp teksten som kommer opp n�r man holder musepekeren over knappen.
# End M� ha

# Kan sl�yfes
__cleanengine__ = True  # Dette forteller tolkeren at den skal sette opp en ny ironpython motor for denne knappen, slik at den ikke kommer i konflikt med andre funksjoner, settes nesten alltid til FALSE, TRUE n�r du jobber med knappen.
__fullframeengine__ = False  # Denne er n�dvendig for � f� tilgang til noen moduler, denne gj�r knappen vesentrlig tregere i oppstart hvis den st�r som TRUE
# __context__ = "zerodoc"  # Denne forteller tolkeren at knappen skal kunne brukes selv om et Revit dokument ikke er �pent.
# __helpurl__ = "google.no"  # Hjelp URL n�r bruker trykker F1 over knapp.
__min_revit_ver__ = 2015  # knapp deaktivert hvis klient bruker lavere versjon
__max_revit_ver__ = 2032  # Skj�nner?
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



import clr

import codecs

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

clr.AddReference("RevitNodes")


xl = Excel.ApplicationClass()
xl.Visible = False
xl.DisplayAlerts = False



from Autodesk.Revit import UI, DB

from Autodesk.Revit.DB import *

from System.Collections.Generic import List

from Autodesk.Revit.DB import Plumbing, IFamilyLoadOptions

def measure(startpoint, point):
    return startpoint.DistanceTo(point)

print('gitkloning dolly')


