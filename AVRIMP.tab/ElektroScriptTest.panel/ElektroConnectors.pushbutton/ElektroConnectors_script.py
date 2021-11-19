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
__title__ = 'ElektroScript'  # Denne overstyrer navnet på scriptfilen
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
from System.Collections.Generic import List

from Autodesk.Revit.DB import Plumbing, IFamilyLoadOptions

def measure(startpoint, point):
    return startpoint.DistanceTo(point)

class FamOpt1(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        return True


def placeConnector(pump, familytype, lineDirection):
    toggle = False
    isVertical = False

    global tempfamtype
    if tempfamtype == None:
        tempfamtype = familytype
        toggle = True
    elif tempfamtype != familytype:
        toggle = True
        tempfamtype = familytype
    level = pump.ReferenceLevel
    point = pump.Origin

    # point = XYZ(point.X,point.Y,point.Z-level.Elevation)
    point = DB.XYZ(point.X, point.Y, point.Z)

    newfam = doc.Create.NewFamilyInstance(point, familytype, level, DB.Structure.StructuralType.NonStructural)
    doc.Regenerate()

    return newfam


def SortedPoints(fittingspoints, ductStartPoint):
    sortedpoints = sorted(fittingspoints, key=lambda x: measure(ductStartPoint, x))
    return sortedpoints

# class for overwriting loaded families in the project
class FamOpt1(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        return True


# Last inn connector element
path = 'S:\Felles\_AVstandard\Revit\Dynamo\Connector_Tekniske_Fag.rfa'
famDoc = app.OpenDocumentFile(path)
famDoc.LoadFamily(doc,FamOpt1())
famDoc.Close(False)

# Finn RIMP link
rvtLinks = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()
max = 0
for link in rvtLinks:
    #print(link.Name)
    if link.Name.count('RIMP') > max:
        max = link.Name.count('RIMP')
        RIMP_link = link

#link.GetLinkDocument()
#link.GetLinkDocument().PathName

print(RIMP_link.Name)


# Finn Pipe Accessories i link
PA_cat = GetCategory(doc, OST_PipeAccessory)

try:
	errorReport = None
	filter = ElementCategoryFilter(System.Enum.ToObject(BuiltInCategory, PA_cat.Id))
	result = FilteredElementCollector(RIMP_link.GetLinkDocument()).WherePasses(filter).WhereElementIsNotElementType().ToElements()
    print('success')
except:
	# if error accurs anywhere in the process catch it
	import traceback
	errorReport = traceback.format_exc()

print(len(result))


#transaction = DB.Transaction(doc)
#transaction.Start("Connectors")

#transaction.Commit()


print('done')
#button = UI.TaskDialogCommonButtons.None
#result = UI.TaskDialogResult.Ok
#UI.TaskDialog.Show('Connector elementer lagt til og oppdater', report_tekst, button)
