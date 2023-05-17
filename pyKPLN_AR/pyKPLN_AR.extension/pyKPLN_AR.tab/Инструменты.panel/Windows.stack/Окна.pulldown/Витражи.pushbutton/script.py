# -*- coding: utf-8 -*-
"""
ОсноваВитраж

"""
__author__ = 'KPLN'
__title__ = "Витражи"
__doc__ = 'Запись имени типа витража\n«00_Основа_Элемента» - имя основы элемента' \

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
from rpw.ui.forms import TextInput, Alert
from rpw.ui.forms import CommandLink, TaskDialog
import datetime

# def get_description(length_feet):
# 	comma = "."
# 	if length_feet >= 0:
# 		sign = "+"
# 	else:
# 		sign = "-"
# 	length_feet_abs = math.fabs(length_feet)
# 	length_meters = int(round(length_feet_abs * 304.8 / 5, 0) * 5)
# 	length_string = str(length_meters)
# 	if len(length_string) == 7:
# 		value = length_string[:4] + comma + length_string[4:]
# 	elif len(length_string) == 6:
# 		value = length_string[:3] + comma + length_string[3:]
# 	elif len(length_string) == 5:
# 		value = length_string[:2] + comma + length_string[2:]
# 	elif len(length_string) == 4:
# 		value = length_string[:1] + comma + length_string[1:]
# 	elif len(length_string) == 3:
# 		value = "0{}".format(comma) + length_string
# 	elif len(length_string) == 2:
# 		value = "0{}0".format(comma) + length_string
# 	elif len(length_string) == 1:
# 		value = "0{}00".format(comma) + length_string
# 	value = sign + value
# 	return value

#ПРОВЕРКА НАЛИЧИЯ ПАРАМЕТРОВ
try:
	param_found = False
	app = doc.Application
	category_set_elements = app.Create.NewCategorySet()
	category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Windows))
	category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Doors))
	category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_MechanicalEquipment))
	category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_CurtainWallMullions))
	category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_CurtainWallPanels))
	originalFile = app.SharedParametersFilename
	app.SharedParametersFilename = "X:\\BIM\\4_ФОП\\00_Архив\\ФОП2019_КП_АР.txt"
	SharedParametersFile = app.OpenSharedParameterFile()
	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	while it.MoveNext():
		d_Definition = it.Key
		d_Name = it.Key.Name
		d_Binding = it.Current
		d_catSet = d_Binding.Categories	
		if d_Name == "00_Основа_Элемента":
			if d_Binding.GetType() == DB.InstanceBinding:
				if str(d_Definition.ParameterType) == "Text":
					if d_Definition.VariesAcrossGroups:
						if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_CurtainWallMullions)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_CurtainWallPanels)):
							param_found = True

	with db.Transaction(name = "AddSharedParameter"):
		for dg in SharedParametersFile.Groups:
			if dg.Name == "АРХИТЕКТУРА - Дополнительные":
				if not param_found:
					externalDefinition = dg.Definitions.get_Item("00_Основа_Элемента")
					newIB = app.Create.NewInstanceBinding(category_set_elements)
					doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
					doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)

	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	with db.Transaction(name = "SetAllowVaryBetweenGroups"):
		while it.MoveNext():
			d_Definition = it.Key
			d_Name = it.Key.Name
			d_Binding = it.Current
			if d_Name == "00_Основа_Элемента":
				d_Definition.SetAllowVaryBetweenGroups(doc, True)
except: pass

#ОСНОВА
categories = [DB.BuiltInCategory.OST_CurtainWallMullions, DB.BuiltInCategory.OST_CurtainWallPanels]
with db.Transaction(name='Запись параметров'):
	for category in categories:
		for element in DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements():
			try:
				host = element.Host.Name
				parameter = element.LookupParameter("00_Основа_Элемента").Set(host)
			except:
				pass