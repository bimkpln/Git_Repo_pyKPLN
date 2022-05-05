# -*- coding: utf-8 -*-
"""
Задать пределы

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Задать пределы"
__doc__ = 'Определение параметров «offset_down» и «offset_up»' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import TextInput

collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()

with db.Transaction(name = "Write value"):
	try:
		for element in collector_elements:
			fam_name = element.Symbol.FamilyName
			if fam_name.startswith("199_Отверстие в стене"):
				try:
					offset_down =math.fabs(element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble())
					level = doc.GetElement(element.LevelId)
					level_elevation = level.Elevation
					offset_up = 3000 / 304.8 - element.LookupParameter("Высота").AsDouble() - offset_down - 250 / 304.8
					levels_all = []
					levels_height_all = []
					levels_height_all_sorted = []
					levels_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
					for lvl in levels_collector:
						levels_all.append(lvl)
						levels_height_all.append(lvl.Elevation)
						levels_height_all_sorted.append(lvl.Elevation)
					levels_height_all_sorted.sort()
					for i in range(0, len(levels_height_all_sorted)):
						if level_elevation == levels_height_all_sorted[i]:
							try:
								offset_up = math.fabs(levels_height_all_sorted[i+1] - level_elevation - element.LookupParameter("Высота").AsDouble() - offset_down - 250 / 304.8)
							except: pass
							element.LookupParameter("offset_down").Set(offset_down)
							element.LookupParameter("offset_up").Set(offset_up)
				except: pass
	except: pass