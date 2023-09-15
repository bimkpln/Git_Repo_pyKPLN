# -*- coding: utf-8 -*-
"""


"""
__author__ = 'KPLN'
__title__ = ""
__doc__ = '' \

"""
Архитекурное бюро KPLN

"""
import math
import clr
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import TextInput, Alert
from rpw.ui.forms import CommandLink, TaskDialog
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows')
from Autodesk.Revit.DB import *


doors = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Doors).ToElements()
curtains = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_CurtainWallPanels).ToElements()
windows = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Windows).ToElements()

for d in doors:
	if type(d) == FamilyInstance:
		if d.Mirrored == True:
			if not 'олотно' in d.Symbol.FamilyName:
				print(d.Id, d.Name.encode(errors='xmlcharrefreplace'))

for w in windows:
	if type(w) == FamilyInstance:
		if w.Mirrored == True:

			print(w.Id, w.Name.encode(errors='xmlcharrefreplace'))

for c in curtains:
	if type(c) == FamilyInstance:
		name = c.Symbol.FamilyName
		if name.startswith('135_'):
			if c.Mirrored == True:
				print(c.Id, c.Name.encode(errors='xmlcharrefreplace'))
