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
__title__ = 'Sjekk logging'  # Denne overstyrer navnet på scriptfilen
__author__ = 'Asplan Viak'  # Dette kommer opp som navnet på utvikler av knappen
__doc__ = "Alt-klikk denne knappen og utforsk kildekoden for å lage egne knapper."  # Dette blir hjelp teksten som kommer opp når man holder musepekeren over knappen.
# End MÅ ha

# Kan sløyfes
__cleanengine__ = False  # Dette forteller tolkeren at den skal sette opp en ny ironpython motor for denne knappen, slik at den ikke kommer i konflikt med andre funksjoner, settes nesten alltid til FALSE, TRUE når du jobber med knappen.
__fullframeengine__ = False  # Denne er nødvendig for å få tilgang til noen moduler, denne gjør knappen vesentrlig tregere i oppstart hvis den står som TRUE
__context__ = "zerodoc"  # Denne forteller tolkeren at knappen skal kunne brukes selv om et Revit dokument ikke er åpent.
__helpurl__ = "google.no"  # Hjelp URL når bruker trykker F1 over knapp.
__min_revit_ver__ = 2010  # knapp deaktivert hvis klient bruker lavere versjon
__max_revit_ver__ = 2032  # Skjønner?
__beta__ = False  # Knapp deaktivert hos brukere som ikke har spesifikt aktivert betaknapper


# Finn flere variabler her:
# https://pyrevit.readthedocs.io/en/latest/articles/scriptanatomy.html

from pyrevit import DB, UI  # Dette er alt du trenger for å få tilgang til nesten hele Revit sin API.
from pyrevit import script, forms  # Se eksempelbruk under

if __shiftclick__:  # variabel som er True hvis bruker shiftklikker
    klikkmetode = "Shiftklikk"
    print "Shiftklikk"
else:
    klikkmetode = "Normalklikk"
    print "Normalklikk"

melding = "Dette er en mal/demoknapp for utviklere,\n" \
          "alt-klikk på knapp for å lese kildekode med mer informasjon."
forms.alert(melding, title=klikkmetode, cancel=True, yes=True, no=True, retry=True, exit=True)
