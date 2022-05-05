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

def GetColor(view, max, min, element):
	global value_max
	global value_min
	for element in DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType():
		try:
			col_value = element.LookupParameter("КР_К_Армирование").AsString()
			value = int(col_value[:(len(col_value)-8)])
			if value_max < value:
				value_max = value
			if value_min > value:
				value_min = value
		except: pass

def GetColorFrom(value, max=800, inverted=False):
	v = int(value * 255 / max)
	if not inverted:
		if v > 255: return 255
		elif v < 0: return 0
		else: return v
	else:
		if v > 255: return 0
		elif v < 0: return 255
		else: return 255-v

def GetColor(view, value_max, value_min, element):
	col_value = element.LookupParameter("КР_К_Армирование").AsString()
	value = value_max - int(col_value[:(len(col_value)-8)])
	if value <= value_max * 0.25:
		x = value/2
		red = 255
		blue = (100 - math.floor(math.fabs(x))) * 255 / 100
		green = math.floor(math.fabs(x)) * 255 / 100
	elif value <= value_max * 0.5:
		x = (value - 200)/2
		red = (100 - math.floor(math.fabs(x))) * 255 / 100
		blue = 0
		green = 255
	elif value <= value_max * 0.75:
		x = (value - 400)/2
		red = 0
		blue = math.floor(math.fabs(x)) * 255 / 100
		green = 255
	elif value <= value_max:
		x = (value - 600)/2
		red = 0
		blue = 255
		green = (100 - math.floor(math.fabs(x))) * 255 / 100
	return color = DB.Color(red, green, blue)

def run(target_view):
	elements = []
	for element in DB.FilteredElementCollector(doc, target_view.Id).WhereElementIsNotElementType():
		try:
			if element.Category.Id == DB.ElementId(-2001320) or element.Category.Id == DB.ElementId(-2000032) or element.Category.Id == DB.ElementId(-2000011) or element.Category.Id == DB.ElementId(-2001300) or element.Category.Id == DB.ElementId(-2001330) or element.Name.startswith("00") or element.Symbol.FamilyName.startswith("00"):
				elements.append(element)
		except:
			if (element.Category.Id == DB.ElementId(-2001320) or element.Category.Id == DB.ElementId(-2000032) or element.Category.Id == DB.ElementId(-2000011) or element.Category.Id == DB.ElementId(-2001300) or element.Category.Id == DB.ElementId(-2001330)):
				elements.append(element)
	with db.Transaction(name = "WriteValue"):
		for element in elements:
			set_color(element, target_view)

def set_color(element, view):
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
		###
		value_max = 800
		value_min = 0
		###
		col_value = 800 - int(col_value[:(len(col_value)-8)])
		if col_value <= 200:
			x = col_value/2
			red = 255
			blue = (100 - math.floor(math.fabs(x))) * 255 / 100
			green = math.floor(math.fabs(x)) * 255 / 100
		elif col_value <= 400:
			x = (col_value - 200)/2
			red = (100 - math.floor(math.fabs(x))) * 255 / 100
			blue = 0
			green = 255
		elif col_value <= 600:
			x = (col_value - 400)/2
			red = 0
			blue = math.floor(math.fabs(x)) * 255 / 100
			green = 255
		elif col_value <= 800:
			x = (col_value - 600)/2
			red = 0
			blue = 255
			green = (100 - math.floor(math.fabs(x))) * 255 / 100
		color = DB.Color(red, green, blue)
		settings.SetSurfaceTransparency(0)
		settings.SetProjectionLineWeight(1)
		settings.SetHalftone(False)
		settings.SetProjectionFillColor(color)
		view.SetElementOverrides(element.Id, settings)
	except:
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