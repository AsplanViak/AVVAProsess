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
__title__ = 'Bytt System Type på rør GPT versjon'  # Denne overstyrer navnet på scriptfilen
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

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Form, Button, ComboBox, ComboBoxStyle, DialogResult

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

class SystemTypeForm(IExternalEventHandler):
    def __init__(self):
        self.system_types = None
        self.selected_type = None

    def Execute(self, uiapp):
        try:
            # Open the Revit document
            uidoc = uiapp.ActiveUIDocument
            doc = uidoc.Document

            # Get the available system types
            self.system_types = doc.GetElement(uidoc.Selection.GetElementIds().FirstElement()).MEPSystem.GetTypeId().GetAllPlumbingSystemTypes()

            # Create a form
            form = SystemTypeSelectionForm(self.system_types)
            form.ShowDialog()

            # Get the selected system type
            self.selected_type = form.selected_type

        except Exception as ex:
            TaskDialog.Show("Error", str(ex))

    def GetName(self):
        return "System Type Selection"

class SystemTypeSelectionForm(Form):
    def __init__(self, system_types):
        self.system_types = system_types
        self.selected_type = None

        self.Text = "Select System Type"
        self.Width = 300
        self.Height = 150

        self.combo_box = ComboBox()
        self.combo_box.Location = System.Drawing.Point(20, 20)
        self.combo_box.Width = 260
        self.combo_box.DropDownStyle = ComboBoxStyle.DropDownList

        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.DialogResult = DialogResult.OK
        self.ok_button.Location = System.Drawing.Point(100, 70)

        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.DialogResult = DialogResult.Cancel
        self.cancel_button.Location = System.Drawing.Point(180, 70)

        for system_type in self.system_types:
            self.combo_box.Items.Add(system_type.Name)

        self.Controls.Add(self.combo_box)
        self.Controls.Add(self.ok_button)
        self.Controls.Add(self.cancel_button)

    def Dispose(self):
        self.selected_type = self.system_types[self.combo_box.SelectedIndex]
        self.Close()

# Custom ExternalCommand class
class SystemTypeSelectionCommand(IExternalCommand):
    def Execute(self, commandData):
        form = SystemTypeForm()
        uiapp = commandData.Application
        uiapp.RegisterExternalEventHandler(form)
        form.Execute(uiapp)
        return Result.Succeeded

# Register the command with PyRevit
from pyrevit import revit, DB
from pyrevit import script

#def push_button():
#command = SystemTypeSelectionCommand()
#script.start_command(command)

SystemTypeSelectionCommand()

# Register the button
#button = revit.FamilyPushButton(push_button, __title__, __doc__)
#button.toolip = 'Select System Type'

