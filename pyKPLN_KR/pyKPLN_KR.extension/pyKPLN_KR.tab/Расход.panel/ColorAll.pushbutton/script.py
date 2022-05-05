# -*- coding: utf-8 -*-
"""
arm_element

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Задать\nцвета"
__doc__ = '' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from rpw.ui import Pick
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import TextInput, Alert, select_folder
import datetime
import System
from System.Windows.Forms import *
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

def GetMax(view, value_max):
	v = value_max
	for element in DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType():
		try:
			col_value = element.LookupParameter("КР_К_Армирование").AsString()
			value = int(col_value[:(len(col_value)-8)])
			if v < value:
				v = value
		except: pass
	return v
def GetMin(view, value_min):
	v = value_min
	for element in DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType():
		try:
			col_value = element.LookupParameter("КР_К_Армирование").AsString()
			value = int(col_value[:(len(col_value)-8)])
			if v > value:
				v = value
		except: pass
	return v
def GetColorFrom(value, min=100, max=800, inverted=False):
	v = int(value * 255 / ((max-min)/4))
	if not inverted:
		if v > 255: return 255
		elif v < 0: return 0
		else: return v
	else:
		if v > 255: return 0
		elif v < 0: return 255
		else: return 255-v

def GetColor(view, value_max, value_min, element):
	try:
		col_value = element.LookupParameter("КР_К_Армирование").AsString()
		value = (value_max - value_min) - (int(col_value[:(len(col_value)-8)]) - value_min)
		if value <= (value_max - value_min) * 0.25:
			red = 255
			blue = GetColorFrom(value, value_min, value_max, True)
			green = GetColorFrom(value, value_min, value_max, False)
		elif value <= (value_max - value_min) * 0.5:
			red = GetColorFrom(value - (value_max - value_min) * 0.25	, value_min, value_max, True)
			blue = 0
			green = 255
		elif value <= (value_max - value_min) * 0.75:
			red = 0
			blue = GetColorFrom(value - (value_max - value_min) * 0.5, value_min, value_max, False)
			green = 255
		elif value <= (value_max - value_min):
			red = 0
			blue = 255
			green = GetColorFrom(value - (value_max - value_min) * 0.75, value_min, value_max, True)
		return DB.Color(red, green, blue)
	except:
		return DB.Color(150, 150, 150)

def run(target_view):
	elements = []
	for element in DB.FilteredElementCollector(doc, target_view.Id).WhereElementIsNotElementType():
		try:
			if element.Category.Id == DB.ElementId(-2001320) or element.Category.Id == DB.ElementId(-2000032) or element.Category.Id == DB.ElementId(-2000011) or element.Category.Id == DB.ElementId(-2001300) or element.Category.Id == DB.ElementId(-2001330) or element.Name.startswith("00") or element.Symbol.FamilyName.startswith("00"):
				elements.append(element)
		except:
			if (element.Category.Id == DB.ElementId(-2001320) or element.Category.Id == DB.ElementId(-2000032) or element.Category.Id == DB.ElementId(-2000011) or element.Category.Id == DB.ElementId(-2001300) or element.Category.Id == DB.ElementId(-2001330)):
				elements.append(element)
	commands = [CommandLink('Да', return_value=True), CommandLink('Нет', return_value=False)]
	dialog = TaskDialog('Использовать относительный цветовой диапазон вида?',
						title = "Армирование",
						title_prefix=False,
						commands=commands,
						show_close=False)
	result = dialog.show()
	with db.Transaction(name = "WriteValue"):
		if result:
			value_min = GetMin(target_view, 800)
			value_max = GetMax(target_view, 0)
		else:
			value_min = 100
			value_max = 800
		for element in elements:
			set_color(element, target_view, value_max, value_min)

def set_color(element, view, value_max, value_min):
	fillPatternElement = None
	dict = ["Solid fill", "Сплошная заливка", "сплошная заливка", "solid fill"]
	for name in dict:
		try:
			fillPatternElement = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Drafting, name)
			if fillPatternElement != None:
				break
		except: pass
	col_value = element.LookupParameter("КР_К_Армирование").AsString()
	error_settings = DB.OverrideGraphicSettings()
	settings = DB.OverrideGraphicSettings()	
	settings.SetProjectionLineWeight(1)
	if fillPatternElement != None:
		settings.SetProjectionFillPatternId(fillPatternElement.Id)
	try:
		color = GetColor(view, value_max, value_min, element)
		settings.SetSurfaceTransparency(0)
		settings.SetProjectionLineWeight(1)
		settings.SetHalftone(False)
		settings.SetProjectionFillColor(color)
		view.SetElementOverrides(element.Id, settings)
	except Exception as e:
		print(str(e))
		color = DB.Color(150, 150, 150)
		settings.SetSurfaceTransparency(10)
		settings.SetProjectionLineWeight(6)
		settings.SetHalftone(False)
		settings.SetProjectionFillColor(color)
		view.SetElementOverrides(element.Id, settings)

mv_name = "SYS_KR_ARM"
if uidoc.ActiveView.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() == mv_name:
	collector_viewFamily = DB.FilteredElementCollector(doc).OfClass(DB.View3D).WhereElementIsNotElementType()
	bool = False
	for i in collector_viewFamily:
		if i.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() == mv_name:
			run(i)
else:
	run(uidoc.ActiveView)