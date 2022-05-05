# -*- coding: utf-8 -*-
"""
Основа дверей

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Уровни и основа дверей"
__doc__ = 'Присвоение значения основы и наименования уровня элемента всем дверям (Наименованием семейства начинается с «100_»)' \

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

def in_list(element, list):
	for i in list:
		if element == i:
			return True
	return False

def logger(n):
	try:
		now = datetime.datetime.now()
		filename = "{}-{}-{}_{}-{}-{}_({})_TOOLS_DOORS_HOST_LEVEL.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), revit.username)
		file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
		text = "report\nfile:{}\nversion:{}\nuser:{}\ncounted:{};".format(doc.PathName, revit.version, revit.username, str(n))
		file.write(text.encode('utf-8'))
		file.close()
	except: pass

collector_elements = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_Doors)

with db.Transaction(name = "Parameters"):
	group = "АРХИТЕКТУРА - Обязательные"
	common_parameters_file = "Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt"
	app = doc.Application
	CategorySet = app.Create.NewCategorySet()
	CategorySet.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Doors))
	CategorySet.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Windows))
	originalFile = app.SharedParametersFilename
	app.SharedParametersFilename = common_parameters_file
	SharedParametersFile = app.OpenSharedParameterFile()
	parameters = ["Э_Основа", "Э_Уровень_Наименование"]
	params_found = []
	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	while it.MoveNext():
		d_Definition = it.Key
		d_Name = it.Key.Name
		d_Binding = it.Current
		d_catSet = d_Binding.Categories	
		for i in range(0, len(parameters)):
			if d_Name == parameters[i]:
				if d_Binding.GetType() == DB.InstanceBinding:
					if str(d_Definition.ParameterType) == "Text":
						if d_Definition.VariesAcrossGroups == True:
							if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Doors)):
								params_found.append(d_Name)
							else:
								pass
	for dg in SharedParametersFile.Groups:
		if dg.Name == group:
			for item in parameters:
				if not in_list(item, params_found):
					externalDefinition = dg.Definitions.get_Item(item)
					newIB = app.Create.NewInstanceBinding(CategorySet)
					doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_IDENTITY_DATA)
					doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_IDENTITY_DATA)
	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	while it.MoveNext():
		for i in range(0, len(parameters)):
			if not in_list(parameters[i], params_found):
				d_Definition = it.Key
				d_Name = it.Key.Name
				d_Binding = it.Current
				if d_Name == parameters[i]:
					d_Definition.SetAllowVaryBetweenGroups(doc, True)
n = 0
with db.Transaction(name = "Write value"):
	for element in collector_elements:
		try:
			name = element.Symbol.FamilyName
			if name.startswith("100"):
				host = element.Host.Name
				parameter = element.LookupParameter(parameters[0])
				parameter.Set(host)
				levelid = element.LevelId
				level = doc.GetElement(levelid)
				name = level.Name
				parameter = element.LookupParameter(parameters[1])
				parameter.Set(name)
			n += 1
		except: pass
		logger(n)