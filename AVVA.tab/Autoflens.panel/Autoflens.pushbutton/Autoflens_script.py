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

# prøv å fjerne disse, og se om fungerer.
# clr.AddReference('ProtoGeometry')
# from Autodesk.DesignScript.Geometry import *

doc = HOST_APP.doc

clr.AddReference("RevitNodes")
# import Revit

# prøv å fjerne disse, og se om fungerer.
#clr.ImportExtensions(Revit.Elements)
#clr.ImportExtensions(Revit.GeometryConversion)

# clr.AddReference("RevitAPI")

# from Autodesk.Revit.DB import *

# clr.AddReference("RevitAPIUI")
from Autodesk.Revit import UI, DB

from System.Collections.Generic import List

# from Autodesk.Revit.DB.Plumbing import *
from Autodesk.Revit.DB import Plumbing

pipingSystem = DB.FilteredElementCollector(doc).OfClass(Plumbing.PipingSystemType).ToElements()

output_report = []
output_report_errors = []

list_piping_system = []
list_piping_system_id = []

for i in pipingSystem:
    list_piping_system.append(i)
    list_piping_system_id.append(i.Id)


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


debug = []
debug2 = []
debug4 = []
debug7 = []

# global variables for rotating new families
tempfamtype = None
xAxis = DB.XYZ(1, 0, 0)


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

    # point = XYZ(point.X,point.Y,point.Z-level.Elevation)
    point = DB.XYZ(point.X, point.Y, point.Z)

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

    result = {}
    connpoints = []
    # famconns = newfam.MEPModel.ConnectorManager.Connectors

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
#class FamOpt1(IFamilyLoadOptions):
class FamOpt1:
    def __init__(self): pass

    def OnFamilyFound(self, familyInUse, overwriteParameterValues): return True

    def OnSharedFamilyFound(self, familyInUse, source, overwriteParameterValues): return True


# function for å endre type connector
def changecontype(con):
    if (con.get_Parameter(DB.BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).Set(20)):
        return True
    else:
        return False


def CheckValveConnectors(valve_family):
    famdoc = doc.EditFamily(valve_family)
    fam_connections = DB.FilteredElementCollector(famdoc).WherePasses(
        con_filter).WhereElementIsNotElementType().ToElements()
    for a in fam_connections:
        try:
            if a.get_Parameter(
                    DB.BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM).AsValueString() == 'Global':
                debug4.append('treff på global')
                try:  # this might fail if the parameter exists or for some other reason
                    if (changecontype(a)):
                        pass
                    else:
                        debug4.append('feil ved forsøk på å endre con type')
                        # 20 is the integer-id for Domestic Cold Water
                    famdoc.LoadFamily(doc, FamOpt1())
                    debug4.append('jabadadoo')
                except:  # you might want to import traceback for a more detailed error report
                    debug4.append('neeeeei')
        except:
            debug4.append('mech equipment?')
    famdoc.Close(False)


def AddFlange(pipe, valve_connector, gasket):
    pointlist = valve_connector.Origin

    # Krage-løsflens
    if gasket:
        familytype = flange_family_type[0]
    else:
        familytype = flange_family_type[1]

    # Sveiseflens
    # if gasket[i]:
    #	familytype = flange_family_type[2]
    # else:
    #	familytype = flange_family_type[3]

    # create duct location line
    ductline = pipe.Location.Curve
    lineDirection = ductline.Direction
    new_flange = placeFitting(pipe, pointlist, familytype, lineDirection)

    return new_flange


###########################################################
## start algorithm for finding missing flanges
###########################################################

# COMPANY SPESIFIC CODE FOR FLANGE TYPES TO BE SELECTED
PA1 = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WhereElementIsElementType()

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


for typ in flange_family_type:
    if typ != 0:
        if typ.IsActive == False:
            typ.Activate()
            doc.Regenerate()

# prepare connection filter for later collector within family editor
con_cat_list = [DB.BuiltInCategory.OST_ConnectorElem]
con_typed_list = List[DB.BuiltInCategory](con_cat_list)
con_filter = DB.ElementMulticategoryFilter(con_typed_list)

# collect all mechanical equipment and pipe accessories in project


cat_list = [DB.BuiltInCategory.OST_PipeAccessory, DB.BuiltInCategory.OST_MechanicalEquipment]
typed_list = List[DB.BuiltInCategory](cat_list)
filter = DB.ElementMulticategoryFilter(typed_list)
EQ = DB.FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements()

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
    # Filter out flanges and other parts where type-name i "Standard"
    if i.Name != 'Standard':
        # Filter out equipment without connectors
        # Find connectors
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
        if len(cons) == 0:
            continue

        # Checking the connector-types of the family
        valve_type_id = i.GetTypeId()
        valve_element_type = doc.GetElement(valve_type_id)
        valve_family = valve_element_type.Family
        valve_family_name = valve_family.Name
        if valve_family.Name not in checked_valve_families:
            CheckValveConnectors(valve_family)
            checked_valve_families.append(valve_family_name)

transaction = DB.Transaction(doc)
transaction.Start("Autoflens")

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
                        x.Owner.Id != j.Owner.Id and x.ConnectorType != DB.ConnectorType.Logical]
            except:
                refs = None
            if type(refs) == list:
                if refs == []:
                    reft = None
                else:
                    refs = refs[0]

            for y in j.AllRefs:
                if y.Owner.Id != j.Owner.Id and y.ConnectorType != DB.ConnectorType.Logical:
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
                        ##pipe.append(refs)
                        pipe = refs
                        # pipe_endpoints.append([refs.Location.Curve.GetEndPoint(0), refs.Location.Curve.GetEndPoint(1)])
                        pipe_endpoints = [refs.Location.Curve.GetEndPoint(0), refs.Location.Curve.GetEndPoint(1)]
                        # pipe_connector.append(pipecon)
                        pipe_connector = pipecon
                        # valve_connector.append(j)
                        valve_connector = j
                        ##problem with this line for equipment with only 1, or 3 or more connections
                        if len(cons) == 2:
                            # opposite_valve_connector.append(cons[1 - n])
                            opposite_valve_connector = cons[1 - n]
                        else:
                            # opposite_valve_connector.append(-1)
                            opposite_valve_connector = -1
                        # valve_number_of_connectors.append(len(cons))
                        valve_number_of_connectors = len(cons)
                        # valve_location.append(i.Location)
                        valve_location = i.Location
                        valve = i
                        if pipe.Location.Curve.GetEndPoint(0).DistanceTo(j.Origin) < pipe.Location.Curve.GetEndPoint(
                                1).DistanceTo(j.Origin):
                            # pipe_endpoint_id.append(0)
                            pipe_endpoint_id = 0
                        else:
                            pipe_endpoint_id = 1
                        # pakning eller ikke
                        if 'innspent' in i.Symbol.FamilyName or 'kyvespjeldsventil' in i.Symbol.FamilyName:
                            gasket = False
                        else:
                            gasket = True

                        ''''# Checking the connector-types of the family
                        valve_type_id = i.GetTypeId()
                        valve_element_type = doc.GetElement(valve_type_id)
                        valve_family = valve_element_type.Family
                        valve_family_name = valve_family.Name
                        if valve_family.Name not in checked_valve_families:
                            transaction.Commit()
                            CheckValveConnectors(valve_family)
                            transaction = Transaction(doc)
                            transaction.Start("Continue main script")

                            checked_valve_families.append(valve_family_name)
                        '''

                        #####################
                        # add flange
                        ####################
                        new_flange = AddFlange(pipe, valve_connector, gasket)
                        # check if flange was created. If not, probably due to too small DN
                        if new_flange == 0:
                            continue
                        else:
                            status = "ok"

                        doc.Regenerate()

                        #########################
                        # check if flange need to be flipped
                        ########################
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

                        if valve_number_of_connectors == 2:
                            debug2.append('gammel variant, fungerer for ventiler med 2 connectors')
                            # gammel variant, fungerte ikke for ventiler med bare en connector
                            if flange_a_con_position.DistanceTo(
                                    opposite_valve_connector.Origin) < flange_b_con_position.DistanceTo(
                                opposite_valve_connector.Origin):
                                need_to_flip = f_cons[0].GetMEPConnectorInfo().IsPrimary
                            else:
                                need_to_flip = f_cons[1].GetMEPConnectorInfo().IsPrimary
                            debug2.append('need_to_flip: ' + str(need_to_flip))
                        else:
                            # ny variant, bruker bounding box location
                            debug2.append('ny variant, bruker bounding box')
                            try:
                                bb = valve.get_BoundingBox(None)
                                if not bb is None:
                                    centre = bb.Min + (bb.Max - bb.Min) / 2
                            except:
                                debug2.append('feil med bounding box løsning')
                            # if flange_a_con_position.DistanceTo(valve_location.Point) < flange_b_con_position.DistanceTo(valve_location.Point):

                            if flange_a_con_position.DistanceTo(centre) < flange_b_con_position.DistanceTo(centre):
                                # flange side a is facing the valve
                                need_to_flip = f_cons[0].GetMEPConnectorInfo().IsPrimary
                            else:
                                need_to_flip = f_cons[1].GetMEPConnectorInfo().IsPrimary
                            debug2.append('need_to_flip: ' + str(need_to_flip))

                        if need_to_flip:
                            try:
                                vector = valve_connector.CoordinateSystem.BasisY
                                line = DB.Line.CreateBound(valve_connector.Origin,
                                                           valve_connector.Origin + vector)
                                # line = UnwrapElement(line)
                                flipped = new_flange.Location.Rotate(line, math.pi)
                            except:
                                status = status + ' failed to flip'
                        doc.Regenerate()
                        ###################################
                        # Move flange
                        ###################################

                        try:
                            if f_cons[0].GetMEPConnectorInfo().IsPrimary:
                                # debug2.append('primary')
                                primary_con_id = 0
                                secondary_con_id = 1
                            else:
                                primary_con_id = 1
                                secondary_con_id = 0
                                # debug2.append('secondary')
                            new_flange.Location.Move((valve_connector.Origin - f_cons[secondary_con_id].Origin))
                        except:
                            status = status + ' failed to move'


                        ########################
                        # Modify pipe endpoints
                        ########################


                        doc.Regenerate()
                        #try:
                        if 1:
                            # modify pipe endpoints
                            if pipe_endpoint_id == 0:
                                new_pipeline = DB.Line.CreateBound(f_cons[primary_con_id].Origin,
                                                                   pipe.DB.Location.Curve.GetEndPoint(1))
                                pipe.DB.Location.Curve = new_pipeline
                                debug2.append('a modify pipe endpoints')
                            else:
                                new_pipeline = DB.Line.CreateBound(pipe.DB.Location.Curve.GetEndPoint(0),
                                                                   f_cons[primary_con_id].Origin)
                                pipe.DB.Location.Curve = new_pipeline
                                debug2.append('b modify pipe endpoints')

                        #except:
                            status = status + 'failed to modify pipe endpoints'

                        doc.Regenerate()
                        duct_piping_system_type = pipe.get_Parameter(
                            DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()

                        for k, sys in enumerate(list_piping_system_id):
                            if sys == duct_piping_system_type:
                                rorsystem = list_piping_system[k]

                        ##########################################
                        # Connect pipes to flange
                        ##########################################

                        try:

                            # disconnect
                            valve_connector.DisconnectFrom(pipe_connector)
                            doc.Regenerate()

                            # make new connections
                            pipe_connector.ConnectTo(f_cons[primary_con_id])

                            f_cons[secondary_con_id].ConnectTo(valve_connector)
                            doc.Regenerate()

                        except:
                            status = status + ' failed to connect pipes to flange'

                        # add to output report
                        # output_report.append(duct_piping_system_type)
                        if status == 'ok':
                            try:
                                output_report.append(
                                    list([str(duct_piping_system_type),
                                          'DN' + str(int(pipe_connector.Radius * 304.8 * 2))]))
                            except:
                                try:
                                    output_report.append(
                                        list(['udefinert system', 'DN' + str(int(pipe_connector.Radius * 304.8 * 2))]))
                                except:
                                    try:
                                        output_report.append(list([str(duct_piping_system_type), 'udefinert DN']))
                                    except:
                                        output_report.append(list(['udefinert system', 'udefinert DN']))
                        else:
                            try:
                                output_report_errors.append(
                                    list([str(duct_piping_system_type),
                                          'DN' + str(int(pipe_connector.Radius * 304.8 * 2)), status]))
                            except:
                                try:
                                    output_report_errors.append(
                                        list(['udefinert system', 'DN' + str(int(pipe_connector.Radius * 304.8 * 2)),
                                              status]))
                                except:
                                    try:
                                        output_report_errors.append(
                                            list([str(duct_piping_system_type), 'udefinert DN', status]))
                                    except:
                                        output_report_errors.append(list(['udefinert system', 'udefinert DN', status]))

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
        report_tekst = report_tekst + ' - ' + j[0][0] + ' ' + str(j[0][1]) + ': ' + str(j[1]) + ' stk ' + str(
            j[0][2]) + '\r\n'

transaction.Commit()

button = UI.TaskDialogCommonButtons.None
result = UI.TaskDialogResult.Ok
UI.TaskDialog.Show('Autoflens ferdig', report_tekst, button)
# TaskDialog.Show('Autoflens ferdig',  ' '.join([str(elem) for elem in debug4]), button)

# OUT = debug2
# OUT = EQ_families
# OUT = debug8
# OUT = output_report
# OUT = output_report_errors
# OUT = 'done'
