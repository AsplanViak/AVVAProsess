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
__title__ = 'Velg ventiler/utstyr'  # Denne overstyrer navnet på scriptfilen
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

modus = 'valgte'   #'alle' eller 'valgte'
flensetype = 'sveiseflens' #'kragelosflens eller 'sveiseflens'

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

    point = DB.XYZ(point.X,point.Y,point.Z-level.Elevation)
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
            #print('Feil med innlasting av family med endrede connectors')
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
pipingSystem = DB.FilteredElementCollector(doc).OfClass(Plumbing.PipingSystemType).ToElements()

list_piping_system = []
list_piping_system_id = []
for i in pipingSystem:
    list_piping_system.append(i)
    list_piping_system_id.append(i.Id)

###########################################################
## start algorithm for finding missing flanges
###########################################################


# prepare connection filter for later collector within family editor
con_cat_list = [DB.BuiltInCategory.OST_ConnectorElem]
con_typed_list = List[DB.BuiltInCategory](con_cat_list)
con_filter = DB.ElementMulticategoryFilter(con_typed_list)

# collect all mechanical equipment and pipe accessories in project
if modus == 'alle':
    cat_list = [DB.BuiltInCategory.OST_PipeAccessory, DB.BuiltInCategory.OST_MechanicalEquipment]
    typed_list = List[DB.BuiltInCategory](cat_list)
    filter = DB.ElementMulticategoryFilter(typed_list)
    EQ = DB.FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements()
else:
    # make selection in UI for selecting pipe accessories and mech eq ++
    cat_list = ['Pipe Accessories', 'Mechanical Equipment']
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
            CheckConnectors(valve_family, 20)
            checked_valve_families.append(valve_family_name)

#transaction = DB.Transaction(doc)
#transaction.Start("Autoflens")

# ACTIVATE FLANGE TYPES TO BE USED
PA1 = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_PipeAccessory).WhereElementIsElementType()

flange_family_type = [0, 0, 0, 0]
n = 0

for i in PA1:
    if flensetype == 'kragelosflens':

        if 'Krage-Løsflens_med pakning' in i.Family.Name:
            print('hit1')
            flange_family_type[0] = i
            #sjekk connector type flens her
            CheckConnectors(i.Family, 28)
            n = n + 1
            continue
        if 'Krage-Løsflens_uten pakning' in i.Family.Name:
            print('hit1')
            flange_family_type[1] = i
            #sjekk connector type flens her
            CheckConnectors(i.Family, 28)
            n = n + 1
            continue
    else:

        if 'Sveiseflens_med pakning' in i.Family.Name:
            print('hit2')
            flange_family_type[2] = i
            n = n + 1
            #sjekk connector type flens her
            CheckConnectors(i.Family, 28)
            continue
        if 'Sveiseflens_uten pakning' in i.Family.Name:
            print('hit3')
            flange_family_type[3] = i
            #sjekk connector type flens her
            CheckConnectors(i.Family, 28)
            n = n + 1
            continue

    if n == 2:
        break

transaction = DB.Transaction(doc)
transaction.Start("Autoflens")

for typ in flange_family_type:
    if typ != 0:
        if typ.IsActive == False:
            typ.Activate()
            doc.Regenerate()


for i in EQ:
    # Filter out flanges and other parts where type-name i "Standard"
    if i.Name != 'Standard':
        # find connectors to valve
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
                    pipe_connector = y

            # check if connected to straight pipe
            if refs is not None:
                cat_name = None
                try:
                    cat_name = refs.Category.Name
                except:
                    continue


                if cat_name == 'Pipe Fittings':

                    status = ' Årsak: Utstyr er koblet direkte mot rørdel "' + refs.Symbol.FamilyName + '"'
                    duct_piping_system_type = refs.get_Parameter(DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
                    output_report_errors.append(report(duct_piping_system_type, pipe_connector, status))
                    continue

                if cat_name == 'Pipes':
                    pipe = refs
                    duct_piping_system_type = pipe.get_Parameter(
                        DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()

                    # checking if pipe is longer than 20mm. don't want to add flanges on very short pipes which should not be there, stuck between valves.
                    # LINE BELOW IS ONLY FOR PROJECTS USING MM AS UNIT
                    if pipe.Location.Curve.GetEndPoint(0).DistanceTo(pipe.Location.Curve.GetEndPoint(1)) < 20 / 304.8:
                        status = ' Årsak: For kort rørstrekk til å få plass til flens.'
                        output_report_errors.append(report(duct_piping_system_type, pipe_connector, status))
                        continue
                    else:
                        # Preparing lists with corresponding indexes:
                        ##pipe.append(refs) ##moved
                        # pipe_endpoints.append([refs.Location.Curve.GetEndPoint(0), refs.Location.Curve.GetEndPoint(1)])
                        pipe_endpoints = [refs.Location.Curve.GetEndPoint(0), refs.Location.Curve.GetEndPoint(1)]
                        # pipe_connector.append(pipeconnector) ##OBS like navn.
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

                        #####################
                        # add flange
                        ####################
                        new_flange = AddFlange(pipe, valve_connector, gasket)
                        # check if flange was created. If not, probably due to too small DN
                        if new_flange == 0:
                            output_report.append(report(duct_piping_system_type, pipe_connector, 'Aktuell flens er ikke lastet inn i prosjektet'))
                            continue
                        elif new_flange == 1:
                            continue

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
                            # fungerer best for ventiler etc med 2 stk connectors
                            flange_a_dist = flange_a_con_position.DistanceTo(opposite_valve_connector.Origin)
                            flange_b_dist = flange_b_con_position.DistanceTo(opposite_valve_connector.Origin)
                            if flange_a_dist < flange_b_dist:
                                need_to_flip = f_cons[0].GetMEPConnectorInfo().IsPrimary
                            else:
                                need_to_flip = f_cons[1].GetMEPConnectorInfo().IsPrimary
                        else:
                            # fungerer for ventiler etc med 1 eller mer enn 2 connectors. Bruker bounding box location
                            try:
                                bb = valve.get_BoundingBox(None)
                                if not bb is None:
                                    centre = bb.Min + (bb.Max - bb.Min) / 2
                            except:
                                pass
                            if flange_a_con_position.DistanceTo(centre) < flange_b_con_position.DistanceTo(centre):
                                # flange side a is facing the valve
                                need_to_flip = f_cons[0].GetMEPConnectorInfo().IsPrimary
                            else:
                                need_to_flip = f_cons[1].GetMEPConnectorInfo().IsPrimary

                        #Flip
                        if need_to_flip:
                            try:
                                vector = valve_connector.CoordinateSystem.BasisY
                                line = DB.Line.CreateBound(valve_connector.Origin,
                                                           valve_connector.Origin + vector)
                                flipped = new_flange.Location.Rotate(line, math.pi)
                            except:
                                status = ' Årsak: Feil ved flipping av flens.'
                                doc.Delete(new_flange.Id)
                                output_report_errors.append(report(duct_piping_system_type, pipe_connector, status))
                                continue

                        doc.Regenerate()

                        ###################################
                        # Move flange
                        ###################################

                        try:
                            if f_cons[0].GetMEPConnectorInfo().IsPrimary:
                                #primary
                                primary_con_id = 0
                                secondary_con_id = 1
                            else:
                                #secondary
                                primary_con_id = 1
                                secondary_con_id = 0
                            new_flange.Location.Move((valve_connector.Origin - f_cons[secondary_con_id].Origin))

                        except:
                            status = ' Årsak: Feil ved flytting av flens.'
                            doc.Delete(new_flange.Id)
                            output_report_errors.append(report(duct_piping_system_type, pipe_connector, status))
                            continue

                        ########################
                        # Modify pipe endpoints and connect to flange
                        ########################

                        doc.Regenerate()
                        try:
                            # modify pipe endpoints
                            if pipe_endpoint_id == 0:
                                new_pipeline = DB.Line.CreateBound(f_cons[primary_con_id].Origin,
                                                                   pipe.Location.Curve.GetEndPoint(1))
                                pipe.Location.Curve = new_pipeline
                                #a modify pipe endpoints
                            else:
                                new_pipeline = DB.Line.CreateBound(pipe.Location.Curve.GetEndPoint(0),
                                                                   f_cons[primary_con_id].Origin)
                                pipe.Location.Curve = new_pipeline
                                #b modify pipe endpoints
                            doc.Regenerate()

                            # disconnect
                            valve_connector.DisconnectFrom(pipe_connector)
                            # doc.Regenerate()

                            # make new connections
                            pipe_connector.ConnectTo(f_cons[primary_con_id])
                            f_cons[secondary_con_id].ConnectTo(valve_connector)
                            doc.Regenerate()

                        except:
                            status = ' Feil ved justering/tilkobling av rør.'
                            doc.Delete(new_flange.Id)
                            output_report_errors.append(report(duct_piping_system_type, pipe_connector, status))
                            continue

                        # add to output report
                        output_report.append(report(duct_piping_system_type,pipe_connector, ''))


report_tekst = ''

if not len(output_report) and not len(output_report_errors):
    report_tekst = 'Ingen flenser ble lagt til. Det fantes ingen koblinger mellom rør og utstyr som mangler flens. \r\n'

if len(output_report):
    report_compressed = []
    for i in output_report:
        funnet = 0
        for a in report_compressed:
            if i == a[0]:
                a[1] = a[1] + 1
                funnet = 1
        if funnet == 0:
            report_compressed.append(list([i, 1]))

    report_tekst = report_tekst + 'Følgende flenser ble lagt til: \r\n \r\n'
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

    report_tekst = report_tekst + '\r\nFølgende flenser ble IKKE lagt til: \r\n \r\n'
    for j in report_errors_compressed:
        report_tekst = report_tekst + ' - ' + j[0][0] + ' ' + str(j[0][1]) + ': ' + str(j[1]) + ' stk ' + str(
            j[0][2]) + '\r\n'

transaction.Commit()

button = UI.TaskDialogCommonButtons.None
result = UI.TaskDialogResult.Ok
UI.TaskDialog.Show('Autoflens ferdig', report_tekst, button)