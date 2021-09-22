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
__title__ = 'Autoflens'  # Denne overstyrer navnet på scriptfilen
__author__ = 'Asplan Viak'  # Dette kommer opp som navnet på utvikler av knappen
__doc__ = "Klikk for å legge til flenser i prosjektet."  # Dette blir hjelp teksten som kommer opp når man holder musepekeren over knappen.
# End MÅ ha

# Kan sløyfes
__cleanengine__ = True  # Dette forteller tolkeren at den skal sette opp en ny ironpython motor for denne knappen, slik at den ikke kommer i konflikt med andre funksjoner, settes nesten alltid til FALSE, TRUE når du jobber med knappen.
__fullframeengine__ = False  # Denne er nødvendig for å få tilgang til noen moduler, denne gjør knappen vesentrlig tregere i oppstart hvis den står som TRUE
#__context__ = "zerodoc"  # Denne forteller tolkeren at knappen skal kunne brukes selv om et Revit dokument ikke er åpent.
#__helpurl__ = "google.no"  # Hjelp URL når bruker trykker F1 over knapp.
__min_revit_ver__ = 2015  # knapp deaktivert hvis klient bruker lavere versjon
__max_revit_ver__ = 2032  # Skjønner?
__beta__ = False  # Knapp deaktivert hos brukere som ikke har spesifikt aktivert betaknapper


# Finn flere variabler her:
# https://pyrevit.readthedocs.io/en/latest/articles/scriptanatomy.html

from pyrevit import DB, UI  # Dette er alt du trenger for å få tilgang til nesten hele Revit sin API.
from pyrevit import script, forms  # Se eksempelbruk under
from pyrevit import forms, HOST_APP, script


#from AVSnippets import Parametere as PA
import AVSnippets as AVS
from Autodesk.Revit import DB, Exceptions

#import datetime
#import AVSnippets.Excel as avsexcel

from pyrevit.forms import alert


import clr
import sys
import math
import msvcrt

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

#doc = DocumentManager.Instance.CurrentDBDocument
doc = HOST_APP.doc
#uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

clr.AddReference("RevitNodes")
import Revit

clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

from System.Collections.Generic import List

from Autodesk.Revit.DB.Plumbing import *

#if __shiftclick__:  # variabel som er True hvis bruker shiftklikker
#    klikkmetode = "Shiftklikk"
#    print "Shiftklikk"
#else:
#    klikkmetode = "Normalklikk"
#    print "Normalklikk"

#melding = "Dette er en mal/demoknapp for utviklere,\n" \
#          "alt-klikk på knapp for å lese kildekode med mer informasjon."
#forms.alert(melding, title=klikkmetode, cancel=True, yes=True, no=True, retry=True, exit=True)

pipingSystem = FilteredElementCollector(doc).OfClass(PipingSystemType).ToElements()

output_report = []
output_report_errors = []

list_piping_system = []
list_piping_system_id = []

for i in pipingSystem:
    list_piping_system.append(i)
    list_piping_system_id.append(i.Id)


def measure(startpoint, point):
    return startpoint.DistanceTo(point)

#@DB.Transaction.ensure('Copy element')
def copyElement(element, oldloc, loc):
    #TransactionManager.Instance.EnsureInTransaction(doc)
    transaction = Transaction(doc)
    transaction.Start("Copy element")
    elementlist = List[ElementId]()
    elementlist.Add(element.Id)
    OffsetZ = (oldloc.Z - loc.Z) * -1
    OffsetX = (oldloc.X - loc.X) * -1
    OffsetY = (oldloc.Y - loc.Y) * -1
    direction = XYZ(OffsetX, OffsetY, OffsetZ)
    newelementId = ElementTransformUtils.CopyElements(doc, elementlist, direction)
    newelement = doc.GetElement(newelementId[0])
    #TransactionManager.Instance.TransactionTaskDone()
    transaction.Commit()
    return newelement


def GetDirection(faminstance):
    for c in faminstance.MEPModel.ConnectorManager.Connectors:
        conn = c
        break
    return conn.CoordinateSystem.BasisZ


def GetClosestDirection(faminstance, lineDirection):
    conndir = None
    flat_linedir = XYZ(lineDirection.X, lineDirection.Y, 0).Normalize()
    for conn in faminstance.MEPModel.ConnectorManager.Connectors:
        conndir = conn.CoordinateSystem.BasisZ
        if flat_linedir.IsAlmostEqualTo(conndir):
            return conndir
    return conndir


#
debug = []

debug7 = []

# global variables for rotating new families
tempfamtype = None
xAxis = XYZ(1, 0, 0)

#@DB.Transaction.ensure('Place fitting')
def placeFitting(duct, point, familytype, lineDirection):
    transaction = Transaction(doc)
    transaction.Start("Place fitting")

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
        if c.ConnectorType != ConnectorType.End:
            continue
        shape = c.Shape
        if shape == ConnectorProfileType.Round:
            radius = c.Radius
            round = True
            break
        elif shape == ConnectorProfileType.Rectangular or shape == ConnectorProfileType.Oval:
            if abs(lineDirection.Z) == 1:
                isVertical = True
                yDir = c.CoordinateSystem.BasisY
            width = c.Width
            height = c.Height
            break

    #TransactionManager.Instance.EnsureInTransaction(doc)
    # point = XYZ(point.X,point.Y,point.Z-level.Elevation)
    point = XYZ(point.X, point.Y, point.Z)

    ## THIS LINE IS DEPENDENT ON UNITS AND PROJECT SETTINGS. LINE BELOW IS FOR PROEJCT USING MM AS UNIT, AND THERE IS NOT ADDED FLANGES FOR DN<45 MM
    if radius < 45 / 304.8 / 2:
        return 0

    newfam = doc.Create.NewFamilyInstance(point, familytype, level, Structure.StructuralType.NonStructural)
    doc.Regenerate()
    transform = newfam.GetTransform()
    axis = Line.CreateUnbound(transform.Origin, transform.BasisZ)
    global xAxis
    if toggle:
        xAxis = GetDirection(newfam)
    zAxis = XYZ(0, 0, 1)

    if isVertical:
        angle = xAxis.AngleOnPlaneTo(yDir, zAxis)
    else:
        angle = xAxis.AngleOnPlaneTo(lineDirection, zAxis)

    ElementTransformUtils.RotateElement(doc, newfam.Id, axis, angle)
    doc.Regenerate()

    if lineDirection.Z != 0:
        newAxis = GetClosestDirection(newfam, lineDirection)
        yAxis = newAxis.CrossProduct(zAxis)
        angle2 = newAxis.AngleOnPlaneTo(lineDirection, yAxis)
        axis2 = Line.CreateUnbound(transform.Origin, yAxis)
        ElementTransformUtils.RotateElement(doc, newfam.Id, axis2, angle2)

    result = {}
    connpoints = []
    # famconns = newfam.MEPModel.ConnectorManager.Connectors

    newfam.LookupParameter('Nominal Radius').Set(radius)
    ##
    # famconns = UnwrapElement(famconns)

    # if round:
    # for conn in famconns:
    ## JF
    #	for j,conn in enumerate(famconns):
    #		if IsParallel(lineDirection,conn.CoordinateSystem.BasisZ) == False:
    #			continue
    #		if conn.Shape != shape:
    #			continue
    # conn.Radius = radius
    ## JF
    #		if j == 0:
    # conn.Radius = 0.32
    #			debug.append('radius ' + str(radius))
    #			debug.append('conn.radius ' + str(conn.Radius))
    #		debug.append('anyway')
    #		connpoints.append(conn.Origin)
    # else:
    #	for conn in famconns:
    #		if IsParallel(lineDirection,conn.CoordinateSystem.BasisZ) == False:
    #			continue
    #		if conn.Shape != shape:
    #			continue
    #		conn.Width = width
    #		conn.Height = height
    #		connpoints.append(conn.Origin)
    #TransactionManager.Instance.TransactionTaskDone()
    # result[newfam] = connpoints
    transaction.Commit()
    return newfam

#@DB.Transaction.ensure('Connect elements')
def ConnectElements(duct, fitting):
    transaction = Transaction(doc)
    transaction.Start("Connect elements")

    ductconns = duct.ConnectorManager.Connectors
    fittingconns = fitting.MEPModel.ConnectorManager.Connectors

    #TransactionManager.Instance.EnsureInTransaction(doc)
    for conn in fittingconns:
        for ductconn in ductconns:
            result = ductconn.Origin.IsAlmostEqualTo(conn.Origin)
            if result:
                ductconn.ConnectTo(conn)
                break
    #TransactionManager.Instance.TransactionTaskDone()
    transaction.Commit()
    return result


def SortedPoints(fittingspoints, ductStartPoint):
    sortedpoints = sorted(fittingspoints, key=lambda x: measure(ductStartPoint, x))
    return sortedpoints


# class for overwriting loaded families in the project
class FamOpt1(IFamilyLoadOptions):
    def __init__(self): pass

    def OnFamilyFound(self, familyInUse, overwriteParameterValues): return True

    def OnSharedFamilyFound(self, familyInUse, source, overwriteParameterValues): return True


#function for å endre type connector
#@DB.Transaction.ensure('Change connection type')
def changecontype(con):
    transaction = Transaction(doc)
    transaction.Start("Change connection type")

    if(con.get_Parameter(BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).Set(20)):
        transaction.Commit()
        return true
    else:
        transaction.Commit()
        return false




###########################################################
## start algorithm for finding missing flanges
###########################################################

# COMPANY SPESIFIC CODE FOR FLANGE TYPES TO BE SELECTED
PA1 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsElementType()

flange_family_type = [0, 0, 0, 0]

for i in PA1:
    if 'Krage-Løsflens_med pakning' in i.Family.Name:
        flange_family_type[0] = i
        break

for j in PA1:
    if 'Krage-Løsflens_uten pakning' in j.Family.Name:
        flange_family_type[1] = j
        break

for k in PA1:
    if 'Sveiseflens_uten pakning' in k.Family.Name:
        flange_family_type[2] = k
        break

for l in PA1:
    if 'Sveiseflens_uten pakning' in l.Family.Name:
        flange_family_type[3] = l
        break

# prepare connection filter for later collector within family editor
con_cat_list = [BuiltInCategory.OST_ConnectorElem]
con_typed_list = List[BuiltInCategory](con_cat_list)
con_filter = ElementMulticategoryFilter(con_typed_list)

# collect all mechanical equipment and pipe accessories in project


cat_list = [BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_MechanicalEquipment]
typed_list = List[BuiltInCategory](cat_list)
filter = ElementMulticategoryFilter(typed_list)
EQ = FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements()

# Gir: Autodesk.Revit.DB.FamilyInstance
# EQ_families = FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType()
# Gir: Autodesk.Revit.DB.FamilySymbol
# EQ_families = FilteredElementCollector(doc).WherePasses(filter).WhereElementIsElementType()

debug4 = []

pipe = []
pipe_connector = []
pipe_endpoints = []
pipe_endpoint_id = []
valve_connector = []
opposite_valve_connector = []
valve_location = []
gasket = []
valve = []
valve_number_of_connectors = []

# list containing all family names where connectors has been checked and potentially modified

checked_valve_families = []

for i in EQ:
    # for i in []:
    # Filter out flanges and other parts where type-name i "Standard"

    if i.Name != 'Standard':

        # find connectorer to valve
        try:
            connectors = i.MEPModel.ConnectorManager.Connectors
        except:
            try:
                connectors = i.ConnectorManager.Connectors
            except:
                connectors = []
        # modify connectorset to be subscriptable
        cons = []
        for kk in connectors:
            cons.append(kk)

        # check what type of parts the connectors are connected to
        # from node "connector.connectedElements from MEPover
        for n, j in enumerate(cons):
            refs = None
            try:
                refs = [doc.GetElement(x.Owner.Id) for x in j.AllRefs if
                        x.Owner.Id != j.Owner.Id and x.ConnectorType != ConnectorType.Logical]
            except:
                refs = None
            if type(refs) == list:
                if refs == []:
                    reft = None
                else:
                    refs = refs[0]

            for y in j.AllRefs:
                if y.Owner.Id != j.Owner.Id and y.ConnectorType != ConnectorType.Logical:
                    pipecon = y

            # check if connected to straight pipe
            if refs is not None:
                cat_name = None
                try:
                    cat_name = refs.Category.Name
                except:
                    continue

                if cat_name == 'Pipes':
                    # checking if pipe is longer than 20mm. don't want to add flanges on very short pipes which should not be there, stuck between valves.
                    # LINE BELOW IS ONLY FOR PROJECTS USING MM AS UNIT
                    if refs.Location.Curve.GetEndPoint(0).DistanceTo(refs.Location.Curve.GetEndPoint(1)) > 20 / 304.8:

                        # Preparing lists with corresponding indexes:
                        pipe.append(refs)
                        pipe_endpoints.append([refs.Location.Curve.GetEndPoint(0), refs.Location.Curve.GetEndPoint(1)])
                        pipe_connector.append(pipecon)
                        valve_connector.append(j)
                        ##problem with this line for equipment with only 1, or 3 or more connections
                        if len(cons) == 2:
                            opposite_valve_connector.append(cons[1 - n])
                        else:
                            opposite_valve_connector.append(-1)
                        valve_number_of_connectors.append(len(cons))
                        valve_location.append(i.Location)
                        valve.append(i)
                        if refs.Location.Curve.GetEndPoint(0).DistanceTo(j.Origin) < refs.Location.Curve.GetEndPoint(
                                1).DistanceTo(j.Origin):
                            pipe_endpoint_id.append(0)
                        else:
                            pipe_endpoint_id.append(1)
                        # Company spesific code determining which type of flange to be used
                        if 'innspent' in i.Symbol.FamilyName or 'kyvespjeldsventil' in i.Symbol.FamilyName:
                            gasket.append(False)
                        else:
                            gasket.append(True)

                        # Checking the connector-types of the family
                        valve_type_id = i.GetTypeId()
                        valve_element_type = doc.GetElement(valve_type_id)
                        valve_family = valve_element_type.Family
                        valve_family_name = valve_family.Name
                        if valve_family.Name not in checked_valve_families:

                            #trans1 = TransactionManager.Instance
                            #trans1.ForceCloseTransaction()  # just to make sure everything is closed down

                            famdoc = doc.EditFamily(valve_family)

                            # Below is moved
                            # con_cat_list = [BuiltInCategory.OST_ConnectorElem]
                            # con_typed_list = List[BuiltInCategory](con_cat_list)
                            # con_filter = ElementMulticategoryFilter(con_typed_list)

                            fam_connections = FilteredElementCollector(famdoc).WherePasses(
                                con_filter).WhereElementIsNotElementType().ToElements()

                            for a in fam_connections:

                                # debug4.append(a.get_Parameter(BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).AsValueString())
                                # debug4.append(a.get_Parameter(BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).AsInteger())
                                try:
                                    if a.get_Parameter(
                                            BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).AsValueString() == 'Global':
                                        debug4.append('treff på global')
                                        try:  # this might fail if the parameter exists or for some other reason
                                            #trans1.EnsureInTransaction(famdoc)
                                            # famdoc.FamilyManager.AddParameter(par_name, par_grp, par_type, inst_or_typ)
                                            ###a.get_Parameter(BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).SetValueString('Domestic Cold Water')

                                            #a.get_Parameter(BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).Set(20)

                                            if(changecontype(a)):
                                                pass
                                            else:
                                                debug4.append('feil ved forsøk på å endre con type')


                                            # 20 is the integer-id for Domestic Cold Water

                                            #trans1.ForceCloseTransaction()
                                            famdoc.LoadFamily(doc, FamOpt1())
                                            # result.append(True)
                                            debug4.append('jabadadoo')
                                        except:  # you might want to import traceback for a more detailed error report
                                            # result.append(False)
                                            debug4.append('neeeeei')
                                            #trans1.ForceCloseTransaction()
                                except:
                                    debug4.append('mech equipment?')

                            famdoc.Close(False)

                            checked_valve_families.append(valve_family_name)

                        # TransactionManager.Instance.EnsureInTransaction(famdoc)

                        # debug4.append('treff på global')
                        # if a.get_Parameter(BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).SetValueString('Domestic Cold Water'):
                        #	debug4.append('suksess')
                        # famdoc.Regenerate()
                        # TransactionManager.Instance.TransactionTaskDone()

############################################################################
#####  Start script for adding flanges. Based on MEPover node
###########################################################################


#TransactionManager.Instance.EnsureInTransaction(doc)
transaction2 = Transaction(doc)
transaction2.Start('activate flange type')
for typ in flange_family_type:
    if typ != 0:
        if typ.IsActive == False:
            typ.Activate()
            doc.Regenerate()
transaction2.Commit()
#TransactionManager.Instance.TransactionTaskDone()

debug2 = []

for i, duct in enumerate(pipe):

    #########################
    ### START MAIN TRY: WHICH IS INDIVIDUAL PER FLANGE
    ########################

    try:
        pointlist = valve_connector[i].Origin

        # Krage-løsflens
        if gasket[i]:
            familytype = flange_family_type[0]
        else:
            familytype = flange_family_type[1]

        # Sveiseflens
        # if gasket[i]:
        #	familytype = flange_family_type[2]
        # else:
        #	familytype = flange_family_type[3]

        # create duct location line
        ductline = duct.Location.Curve

        lineDirection = ductline.Direction

        new_flange = placeFitting(duct, pointlist, familytype, lineDirection)

        # check if flange was created. If not, probably due to too small DN
        if new_flange == 0:
            continue
        debug2.append('            ')

        # check if flange need to be flipped
        try:
            flange_connectors = new_flange.MEPModel.ConnectorManager.Connectors
        except:
            try:
                flange_connectors = new_flange.ConnectorManager.Connectors
            except:
                flange_connectors = []
                cons = []
        # make subscriptable
        f_cons = []
        for a in flange_connectors:
            f_cons.append(a)
        # valve_cons.append([x for x in connectors])
        flange_a_con_position = f_cons[0].Origin
        flange_b_con_position = f_cons[1].Origin

        if valve_number_of_connectors[i] == 2:
            debug2.append('gammel variant, fungerer for ventiler med 2 connectors')
            # gammel variant, fungerte ikke for ventiler med bare en connector
            if flange_a_con_position.DistanceTo(opposite_valve_connector[i].Origin) < flange_b_con_position.DistanceTo(
                    opposite_valve_connector[i].Origin):
                need_to_flip = f_cons[0].GetMEPConnectorInfo().IsPrimary
            else:
                need_to_flip = f_cons[1].GetMEPConnectorInfo().IsPrimary
            debug2.append('need_to_flip: ' + str(need_to_flip))
        else:
            # ny variant, bruker location
            # boundign box to find center of location
            debug2.append('ny variant, bruker bounding box')
            try:
                bb = valve[i].get_BoundingBox(None)
                if not bb is None:
                    centre = bb.Min + (bb.Max - bb.Min) / 2
            except:
                debug2.append('feil med bounding box løsning')
            # if flange_a_con_position.DistanceTo(valve_location[i].Point) < flange_b_con_position.DistanceTo(valve_location[i].Point):

            if flange_a_con_position.DistanceTo(centre) < flange_b_con_position.DistanceTo(centre):
                # flange side a is facing the valve
                need_to_flip = f_cons[0].GetMEPConnectorInfo().IsPrimary
            else:
                need_to_flip = f_cons[1].GetMEPConnectorInfo().IsPrimary
            debug2.append('need_to_flip: ' + str(need_to_flip))

        #TransactionManager.Instance.EnsureInTransaction(doc)
        transaction = Transaction(doc)
        transaction.start('Flip and move flange')
        if need_to_flip:

            vector = valve_connector[i].CoordinateSystem.BasisY

            line = Autodesk.Revit.DB.Line.CreateBound(valve_connector[i].Origin, valve_connector[i].Origin + vector)
            line = UnwrapElement(line)

            flipped = new_flange.Location.Rotate(line, math.pi)

        ##Move to right place
        # find primary and secondardy connector

        doc.Regenerate()

        if f_cons[0].GetMEPConnectorInfo().IsPrimary:
            # debug2.append('primary')
            primary_con_id = 0
            secondary_con_id = 1
        else:
            primary_con_id = 1
            secondary_con_id = 0
        # debug2.append('secondary')

        doc.Regenerate()

        new_flange.Location.Move((valve_connector[i].Origin - f_cons[secondary_con_id].Origin))

        doc.Regenerate()

        # modify pipe endpoints
        try:
            if pipe_endpoint_id[i] == 0:
                new_pipeline = Autodesk.Revit.DB.Line.CreateBound(f_cons[primary_con_id].Origin,
                                                                  pipe[i].Location.Curve.GetEndPoint(1))
                pipe[i].Location.Curve = new_pipeline
                debug2.append('a modify pipe endpoints')
            else:
                new_pipeline = Autodesk.Revit.DB.Line.CreateBound(pipe[i].Location.Curve.GetEndPoint(0),
                                                                  f_cons[primary_con_id].Origin)
                pipe[i].Location.Curve = new_pipeline
                debug2.append('b modify pipe endpoints')
        except:
            debug2.append('fail')

        doc.Regenerate()

        # duct_piping_system_type = pipe_connector[i].MEPSystem
        # duct_piping_system_type = pipe[i].get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
        duct_piping_system_type = pipe[i].get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()

        # Gir: Rentvann UV
        # debug2.append(duct_piping_system_type)
        # Gir: Rentvann UV 4
        # debug2.append(valve_connector[i].MEPSystem.Name)
        # Gir:
        # debug2.append(valve_connector[i].PipeSystemType.ToString())

        for k, sys in enumerate(list_piping_system_id):
            if sys == duct_piping_system_type:
                rorsystem = list_piping_system[k]

        # try:
        if 1:
            # disconnect
            valve_connector[i].DisconnectFrom(pipe_connector[i])
            doc.Regenerate()

            # make new connections
            pipe_connector[i].ConnectTo(f_cons[primary_con_id])
            # doc.Regenerate()
            f_cons[secondary_con_id].ConnectTo(valve_connector[i])

            doc.Regenerate()
        # except:
        # debug2.append('fail 2')

        #TransactionManager.Instance.TransactionTaskDone()
        transaction.Commit()

        # add to output report
        # output_report.append(duct_piping_system_type)
        # output_report.append(list([str(int(pipe_connector[i].Radius*304.8*2)),duct_piping_system_type,1]))
        try:
            output_report.append(
                list([str(duct_piping_system_type), 'DN' + str(int(pipe_connector[i].Radius * 304.8 * 2))]))
        except:
            try:
                output_report.append(list(['udefinert system', 'DN' + str(int(pipe_connector[i].Radius * 304.8 * 2))]))
            except:
                try:
                    output_report.append(list([str(duct_piping_system_type), 'udefinert DN']))
                except:
                    output_report.append(list(['udefinert system', 'udefinert DN']))

        debug2.append(new_flange)
        debug2.append(valve_number_of_connectors[i])

    except:
        try:
            output_report_errors.append(
                list([str(duct_piping_system_type), 'DN' + str(int(pipe_connector[i].Radius * 304.8 * 2))]))
        except:
            try:
                output_report_errors.append(
                    list(['udefinert system', 'DN' + str(int(pipe_connector[i].Radius * 304.8 * 2))]))
            except:
                try:
                    output_report_errors.append(list([str(duct_piping_system_type), 'udefinert DN']))
                except:
                    output_report_errors.append(list(['udefinert system', 'udefinert DN']))

debug8 = valve_number_of_connectors

if not len(output_report):
    report_tekst = 'Ingen flenser ble lagt til. Det fantes ingen koblinger mellom rør og utstyr som mangler flens. \r\n'
else:
    report_compressed = []
    for i in output_report:
        funnet = 0
        for a in report_compressed:
            if i == a[0]:
                a[1] = a[1] + 1
                funnet = 1
        if funnet == 0:
            report_compressed.append(list([i, 1]))

    report_tekst = 'Følgende flenser ble lagt til: \r\n \r\n'
    for j in report_compressed:
        report_tekst = report_tekst + ' - ' + j[0][0] + ' ' + str(j[0][1]) + ': ' + str(j[1]) + ' stk \r\n'

if len(output_report_errors):
    report_errors_compressed = []
    for i in output_report_errors:
        funnet = 0
        for a in report_errors_compressed:
            if i == a[0]:
                a[1] = a[1] + 1
                funnet = 1
        if funnet == 0:
            report_errors_compressed.append(list([i, 1]))

    report_tekst = report_tekst + '\r\nFølgende flenser ble IKKE lagt til eller er feilplassert: \r\n \r\n'
    for j in report_errors_compressed:
        report_tekst = report_tekst + ' - ' + j[0][0] + ' ' + str(j[0][1]) + ': ' + str(j[1]) + ' stk \r\n'

button = TaskDialogCommonButtons.None
result = TaskDialogResult.Ok
TaskDialog.Show('Autoflens ferdig', report_tekst, button)

# OUT = debug2
# OUT = EQ_families
# OUT = debug8
# OUT = output_report
#OUT = output_report_errors
# OUT = 'done'
