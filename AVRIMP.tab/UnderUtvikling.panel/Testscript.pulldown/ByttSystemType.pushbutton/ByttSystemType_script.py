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
__title__ = 'Bytt System Type på rør'  # Denne overstyrer navnet på scriptfilen
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
#uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

clr.AddReference("RevitNodes")

from Autodesk.Revit import UI, DB
from Autodesk.Revit.UI.Selection import ObjectType

from System.Collections.Generic import List

from Autodesk.Revit.DB import Plumbing, IFamilyLoadOptions

from Autodesk.Revit.DB import *

import os


clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')

from Microsoft.Office.Interop import Excel
from Microsoft.Office.Interop.Excel import XlListObjectSourceType, Worksheet, Range, XlYesNoGuess

from System.Collections.Generic import List
from System.Collections.Generic import *
from System import Guid
from System import Array

import datetime



def measure(startpoint, point):
    return startpoint.DistanceTo(point)

def copyElement(element, oldloc, loc):
    elementlist = List[DB.ElementId]()
    elementlist.Add(element.Id)
    OffsetZ = (oldloc.Z - loc.Z) * -1
    OffsetX = (oldloc.X - loc.X) * -1
    OffsetY = (oldloc.Y - loc.Y) * -1
    direction = DB.XYZ(OffsetX, OffsetY, OffsetZ)
    newelementId = DB.ElementTransformUtils.CopyElements(doc, elementlist, direction)
    newelement = doc.GetElement(newelementId[0])
    return newelement

def GetDirection(faminstance):
    for c in faminstance.MEPModel.ConnectorManager.Connectors:
        conn = c
        break
    return conn.CoordinateSystem.BasisZ

def GetClosestDirection(faminstance, lineDirection):
    conndir = None
    flat_linedir = DB.XYZ(lineDirection.X, lineDirection.Y, 0).Normalize()
    for conn in faminstance.MEPModel.ConnectorManager.Connectors:
        conndir = conn.CoordinateSystem.BasisZ
        if flat_linedir.IsAlmostEqualTo(conndir):
            return conndir
    return conndir

# global variables for rotating new families
tempfamtype = None
xAxis = DB.XYZ(1, 0, 0)

transaction = DB.Transaction(doc)
transaction.Start("Bytt System Type")

def report(duct_piping_system_type, pipe_connector, status):
    try:
        report_row = list([str(duct_piping_system_type), 'DN' + str(int(pipe_connector.Radius * 304.8 * 2)), status])
    except:
        try:
            report_row = list(['udefinert system', 'DN' + str(int(pipe_connector.Radius * 304.8 * 2)), status])
        except:
            try:
                report_row = list([str(duct_piping_system_type), 'udefinert DN', status])
            except:
                report_row =  list(['udefinert system', 'udefinert DN', status])
    return report_row

def placeFitting(duct, point, familytype, lineDirection):
    toggle = False
    isVertical = False

    global tempfamtype
    if tempfamtype == None:
        tempfamtype = familytype
        toggle = True
    elif tempfamtype != familytype:
        toggle = True
        tempfamtype = familytype
    level = duct.ReferenceLevel

    width = 4
    height = 4
    radius = 2
    round = False
    connectors = duct.ConnectorManager.Connectors
    for c in connectors:
        if c.ConnectorType != DB.ConnectorType.End:
            continue
        shape = c.Shape
        if shape == DB.ConnectorProfileType.Round:
            radius = c.Radius
            round = True
            break
        elif shape == DB.ConnectorProfileType.Rectangular or shape == DB.ConnectorProfileType.Oval:
            if abs(lineDirection.Z) == 1:
                isVertical = True
                yDir = c.CoordinateSystem.BasisY
            width = c.Width
            height = c.Height
            break
    print('point.Z :' + str(point.Z))
    #point = DB.XYZ(point.X,point.Y,point.Z-level.Elevation)
    point = DB.XYZ(point.X, point.Y, point.Z - level.ProjectElevation)
    print('ny point.Z :' + str(point.Z))
    print('level.ProjectElevation :' +str(level.ProjectElevation))
    print('level.Elevation :' + str(level.Elevation))
    #point = DB.XYZ(point.X, point.Y, point.Z)

    ## THIS LINE IS DEPENDENT ON UNITS AND PROJECT SETTINGS. LINE BELOW IS FOR PROEJCT USING MM AS UNIT, AND THERE IS NOT ADDED FLANGES FOR DN<45 MM
    if radius < 45 / 304.8 / 2:
        return 0

    newfam = doc.Create.NewFamilyInstance(point, familytype, level, DB.Structure.StructuralType.NonStructural)
    doc.Regenerate()
    transform = newfam.GetTransform()
    axis = DB.Line.CreateUnbound(transform.Origin, transform.BasisZ)
    global xAxis
    if toggle:
        xAxis = GetDirection(newfam)
    zAxis = DB.XYZ(0, 0, 1)

    if isVertical:
        angle = xAxis.AngleOnPlaneTo(yDir, zAxis)
    else:
        angle = xAxis.AngleOnPlaneTo(lineDirection, zAxis)

    DB.ElementTransformUtils.RotateElement(doc, newfam.Id, axis, angle)
    doc.Regenerate()

    if lineDirection.Z != 0:
        newAxis = GetClosestDirection(newfam, lineDirection)
        yAxis = newAxis.CrossProduct(zAxis)
        angle2 = newAxis.AngleOnPlaneTo(lineDirection, yAxis)
        axis2 = DB.Line.CreateUnbound(transform.Origin, yAxis)
        DB.ElementTransformUtils.RotateElement(doc, newfam.Id, axis2, angle2)

    newfam.LookupParameter('Nominal Radius').Set(radius)

    return newfam


def ConnectElements(duct, fitting):

    ductconns = duct.ConnectorManager.Connectors
    fittingconns = fitting.MEPModel.ConnectorManager.Connectors

    for conn in fittingconns:
        for ductconn in ductconns:
            result = ductconn.Origin.IsAlmostEqualTo(conn.Origin)
            if result:
                ductconn.ConnectTo(conn)
                break

    return result


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


# function for å endre type connector
def changecontype(con, typeid):
    # 20 is the integer-id for Domestic Cold Water
    #Hydronic supply: 7
    #Fitting: 28
    #
    if (con.get_Parameter(DB.BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).Set(typeid)):
        return True
    else:
        return False

def CheckConnectors(family, typeid):
    #typeid 20 = Domestic cold water. typeid 28 = Fitting
    famdoc = doc.EditFamily(family)
    fam_connections = DB.FilteredElementCollector(famdoc).WherePasses(
        con_filter).WhereElementIsNotElementType().ToElements()
    changed = False
    for a in fam_connections:
        try:
            if a.get_Parameter(
                    DB.BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM) != typeid:

                try:  # this might fail if the parameter exists or for some other reason
                    if (changecontype(a, typeid)):
                        # success
                        changed = True
                        #pass
                    else:
                        #feil ved forsøk på å endre con type
                        pass
                except:
                    pass
                    #print('Feil med endring av connector-type i family')
            else:
                pass
                #print('Connector har riktig type')
        except:
            pass
            #print('Feil med sjekk av connector-type i family')
    #print('changed :' + str(changed))
    if changed:
        #famdoc.LoadFamily.Overloads.Functions[3](Document=doc, IFamilyLoadOptions=FamOpt1())
        try:
            famdoc.LoadFamily.Overloads.Functions[3](doc, FamOpt1())
        except:
            print('Feil med innlasting av family med endrede connectors')
    famdoc.Close(False)

def AddFlange(pipe, valve_connector, gasket):
    pointlist = valve_connector.Origin

    if flensetype == 'kragelosflens':
        # Krage-løsflens
        if gasket:
            familytype = flange_family_type[0]
        else:
            familytype = flange_family_type[1]
    else:
        # Sveiseflens
        if gasket:
            familytype = flange_family_type[2]
        else:
            familytype = flange_family_type[3]

    if isinstance(familytype, int):
        #Flens har ikke blitt assigned. Mest sannsynlig for aktuell flens mangler i aktuell revit-fil.
        return 0

    # create duct location line
    ductline = pipe.Location.Curve
    lineDirection = ductline.Direction
    try:
        new_flange = placeFitting(pipe, pointlist, familytype, lineDirection)
    except:
        new_flange = 1

    return new_flange

#klargjør til rapportering til skjerm
output_report = []
output_report_errors = []

#lag liste over alle piping systems
#pipingSystem = DB.FilteredElementCollector(doc).OfClass(Plumbing.PipingSystemType).ToElements()
#pipingSystem = DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsElementType().ToElements()

pipingSystem = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsElementType().ToElements()

list_piping_system = []
list_piping_system_id_lesbar = []
list_piping_system_id = []
for i in pipingSystem:
    #list_piping_system.append(i)
    #list_piping_system.append(i.AsString())
    #list_piping_system.append(i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsString())
    list_piping_system.append(Element.Name.GetValue(i))
    list_piping_system_id_lesbar.append(i.Id.IntegerValue.ToString())
    list_piping_system_id.append(i.Id)
    #list_piping_system_id.append(i.Id.AsInteger())
print list_piping_system
print list_piping_system_id_lesbar
###########################################################
## start algorithm for changing systemt type
###########################################################
#cat_list = ['Pipe Accessories', 'Mechanical Equipment', 'Generic Model']
cat_list = ['Pipes']
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

# list containing all family names where connectors has been checked and potentially modified
checked_valve_families = []

for i in EQ:
    systemtype = (i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString())
    print systemtype
    systemid = (i.Id.IntegerValue.ToString())
    print systemid
    try:
        # PI_systemtype.append(i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()).ToDSType(True).Name
        # systemtypeparam = (i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString())
        # systemtype = systemtypeparam.ToDSType(True).Name

        #res = i.LookupParameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).Set(list_piping_system_id[30])
        res = i.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).Set(list_piping_system_id[30])
        #res = i.MEPSystem.GetTypeId()).SetType(new_system_type
        if res:
            print("solid")
        else:
            print("Njet")
    except:
        systemtype = 'udefinert system type'
        print ("gikk ikke så bra")


transaction.Commit()

button = UI.TaskDialogCommonButtons.None
result = UI.TaskDialogResult.Ok
UI.TaskDialog.Show('Autoflens ferdig', 'bla blad', button)