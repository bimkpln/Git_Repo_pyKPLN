# -*- coding: utf-8 -*-
"""
Color_By_RoomId

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Окрасить по помещению"
__doc__ = 'Окраска элементов отделки по помещениям.\n\n' \
          'Необходимо запускать на 3D виде' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI, coreutils
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import Alert

elements = []
rooms = []
view = doc.ActiveView
if str(view.ViewType) == "ThreeD":
	dict = ["Solid fill", "Сплошная заливка", "сплошная заливка", "solid fill"]
	for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
		rooms.append(room)
	for element in DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements():
		elements.append(element)
	for element in DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements():
		elements.append(element)
	for element in DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements():
		elements.append(element)
	for name in dict:
		try:
			fillPatternElement = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Drafting, name)
			if fillPatternElement != None:
				break
		except: pass
	if fillPatternElement != None:
		error_settings = DB.OverrideGraphicSettings()
		white = DB.Color(255, 50, 10)
		error_settings.SetProjectionFillColor(white)
		error_settings.SetProjectionLineWeight(7)
		error_settings.SetHalftone(False)
		error_settings.SetProjectionFillPatternId(fillPatternElement.Id)
		error_settings.SetSurfaceTransparency(0)
		with db.Transaction(name='color'):
			if len(rooms) != 0:
				for room in rooms:
					settings = DB.OverrideGraphicSettings()
					red = coreutils.random_color()
					green = coreutils.random_color()
					blue = coreutils.random_color()
					color = DB.Color(red, green, blue)
					settings.SetProjectionFillColor(color)
					settings.SetProjectionLineWeight(1)
					settings.SetHalftone(False)
					settings.SetProjectionFillPatternId(fillPatternElement.Id)
					settings.SetSurfaceTransparency(70)
					room_id = str(room.Id)
					if len(elements) != 0:
						for element in elements:
							element_id = "_"
							try:
								element_id = str(element.LookupParameter("О_Id помещения").AsString())
							except: pass
							if element_id == room_id:
								view.SetElementOverrides(element.Id, settings)
							elif element_id == "" or element_id.lower() == "null" or element_id.lower() == "none":
								view.SetElementOverrides(element.Id, error_settings)
					else: Alert("Отсутствуют элементы на виде/в проекте", title="Отделка", header = "Ошибка")
			else: Alert("В проекте отсутствуют помещения!", title="Отделка", header = "Ошибка")
	else:
		Alert("В проекте отсутствует «Сплошная штриховка»!", title="Отделка", header = "Ошибка")
else: Alert("Необходимо запускать скрипт на 3D виде!", title="Отделка", header = "Ошибка")