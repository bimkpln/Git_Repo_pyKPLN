# -*- coding: utf-8 -*-
"""
Вырезать отверстия для шахт

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Отверстия для шахт"
__doc__ = 'Создание отверстий для шахт в монолитных перекрытиях' \

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

from rpw.ui.forms import TextInput, Alert, select_folder
import datetime

def intersects(floor, element):
	e_solid = element.get_Geometry(DB.Options())
	element_bb = e_solid.GetBoundingBox()
	f_solid = floor.get_Geometry(DB.Options())
	floor_bb = f_solid.GetBoundingBox()
	element_max = element_bb.Max
	element_min = element_bb.Min
	floor_max = floor_bb.Max
	floor_min = floor_bb.Min
	if floor_max.Z < element_min.Z or floor_min.Z > element_max.Z:
		return False
	elif floor_max.Y < element_min.Y or floor_min.Y > element_max.Y:
		return False
	elif floor_max.X < element_min.X or floor_min.X > element_max.X:
		return False
	return True

def logger(errors):
	try:
		now = datetime.datetime.now()
		filename = "{}-{}-{}_{}-{}-{}_({})_ОТВЕРСТИЯ_ШАХТЫ.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), revit.username)
		file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
		text = "report\nfile:{}\nversion:{}\nuser:{}\nerrors:{};".format(doc.PathName, revit.version, revit.username, str(errors))
		file.write(text.encode('utf-8'))
		file.close()
	except: pass

collector_elements = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
collector_floors = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType()
collector_viewFamilyType = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType()
elements = []
floors = []
errors = 0

for i in collector_viewFamilyType:
	if i.ViewFamily == DB.ViewFamily.ThreeDimensional:
		viewFamilyType = i
		break

for element in collector_elements:
	fam_name = element.Symbol.FamilyName 
	if fam_name.lower().startswith("199_шахта") or "шахта" in fam_name.lower() or "шахты" in fam_name.lower():
		elements.append(element)
			
for element in collector_floors:
	floor_name = element.Name
	if floor_name.startswith("00_"):
		floors.append(element)

with db.Transaction(name='Обрезка'):
	View_3D = DB.View3D.CreateIsometric(doc, viewFamilyType.Id)
	for element in elements:
		for floor in floors:
			if intersects(floor, element):
				try:
					if DB.InstanceVoidCutUtils.CanBeCutWithVoid(floor):
						DB.InstanceVoidCutUtils.AddInstanceVoidCut(doc, floor, element)
					else:
						errors += 1
				except Exception as e:
					print(str(e))
					errors += 1
	try: doc.Delete(View_3D.Id)
	except: errors += 1
logger(errors)