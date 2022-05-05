# -*- coding: utf-8 -*-
"""
INGRD_EXPORT

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Выгрузка"
__doc__ = 'Заполнение системных параметров INGRD с разгруппировкой всех элементов проекта!' \
"""
"""
from pyrevit.framework import clr
import re
import os
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction
from System.Collections.Generic import *
from System.Windows.Forms import *
from System.Drawing import *
import System
from System import Guid
import datetime
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

errors = []
def PrintError(error_description, element, e):
	string = "Необходимо проверить параметр «{}» в категории «{}»\n\t\tОшибка: {};".format(error_description, str(element.Category.Name), str(e))
	if not string in errors:
		print(string)
		errors.append(string)

def SetLevel(element):
	try:
		level = doc.GetElement(element.LevelId)
		level_name = level.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString()
		if "этаж" in level_name.lower() and " " in level_name:
			value = level_name.split(" ")[0]
			element.LookupParameter("INGD_Номер этажа").Set(str(value))
		elif "техпространство" in level_name.lower() and " " in level_name: 
			value = level_name.split(" ")[0]
			value += ".1"
			element.LookupParameter("INGD_Номер этажа").Set(str(value))
		else:
			value = level_name.split(" ")[0]
			element.LookupParameter("INGD_Номер этажа").Set(str(value))

	except:
		try:
			level = element.Host
			level_name = level.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString()
			if "этаж" in level_name.lower() and " " in level_name:
				value = level_name.split(" ")[0]
				element.LookupParameter("INGD_Номер этажа").Set(str(value))
			elif "техпространство" in level_name.lower() and " " in level_name: 
				value = level_name.split(" ")[0]
				value += ".1"
				element.LookupParameter("INGD_Номер этажа").Set(str(value))
			else:
				value = level_name.split(" ")[0]
				element.LookupParameter("INGD_Номер этажа").Set(str(value))
		except:
			try:
				level = doc.GetElement(element.Host.LevelId)
				level_name = level.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString()
				if "этаж" in level_name.lower() and " " in level_name:
					value = level_name.split(" ")[0]
					element.LookupParameter("INGD_Номер этажа").Set(str(value))
				elif "техпространство" in level_name.lower() and " " in level_name: 
					value = level_name.split(" ")[0]
					value += ".1"
					element.LookupParameter("INGD_Номер этажа").Set(str(value))
				else:
					value = level_name.split(" ")[0]
					element.LookupParameter("INGD_Номер этажа").Set(str(value))
			except:
				pass

def GetValue(element):
	try:
		value = element.LookupParameter("КР_К_Армирование").AsString()
		num = int(value[:(len(value)-8)])
		return num
	except:
		return None

def SetARM(element):
	try:
		v = GetValue(element)
		if v != None:
			element.get_Parameter(Guid("e7694374-128a-4a72-9183-d4c163130fab")).Set(str(v))
	except: pass

def SetValue(height=None, width=None, length=None, element=None):
	if element != None:
		if height != None:
			try:
				element.get_Parameter(Guid("4cca470f-775b-4166-9161-5b13cd33dc41")).Set(height)
			except Exception as e:
				PrintError("INGD_Высота", element, e)
		if width != None:
			try:
				element.get_Parameter(Guid("9dec54c1-6a57-4f35-b2ca-64bfb877226b")).Set(width)
			except Exception as e: 
				PrintError("INGD_Ширина", element, e)
		if length != None:
			try:
				element.get_Parameter(Guid("4b318cfa-7db5-4870-8fbf-5799e3d542d6")).Set(length)
			except Exception as e: 
				PrintError("INGD_Длина", element, e)

with db.Transaction(name = "INGRD"):
	for group in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_IOSModelGroups).WhereElementIsNotElementType().ToElements():
		elements = group.UngroupMembers()
	doc.Regenerate()
	for element in DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
		SetLevel(element)
		SetARM(element)
		if type(element) == DB.Floor:
			SetValue(height=element.FloorType.GetCompoundStructure().GetWidth(), width=None, length=None, element=element)
		if type(element) == DB.Wall:
			SetValue(height=element.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble(), width=element.Width, length=element.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), element=element)
		if type(element) == DB.FamilyInstance:
			if type(element.GetAnalyticalModel()) == DB.Structure.AnalyticalModelStick:
				SetValue(height=None, width=None, length=element.get_Parameter(DB.BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH).AsDouble(), element=element)