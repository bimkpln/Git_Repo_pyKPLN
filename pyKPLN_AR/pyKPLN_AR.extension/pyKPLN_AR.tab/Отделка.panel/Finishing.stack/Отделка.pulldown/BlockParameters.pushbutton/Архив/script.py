# -*- coding: utf-8 -*-
"""
ROOMS_SharedParameters

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Заблокировать площади стадии П"
__doc__ = ''
"""
Архитекурное бюро KPLN

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
import System
import datetime
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

def in_list(element, list):
	for i in list:
		if element == i:
			return True
	return False

if os.path.exists("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\lib\\ФОП_Scripts.txt"):
	commands = [CommandLink('Да', return_value=True), CommandLink('Отмена', return_value=False)]
	dialog = TaskDialog('Добавить параметры в проект?',
						title = "Загрузчик параметров",
						title_prefix  =False,
						content = "ФОП - Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\lib\\ФОП_Scripts.txt",
						commands = commands,
						footer = 'См. правила работы с отделкой',
						show_close = False)
	result = dialog.show()
	if result:
		try:
			group = "АРХИТЕКТУРА - Квартирография"
			params_elements = ["П_ПОМ_Площадь",
						"П_ПОМ_Площадь_К",
						"П_КВ_Площадь_Общая",
						"П_КВ_Площадь_Жилая",
						"П_КВ_Площадь_Летние",
						"П_КВ_Площадь_Нежилая",
						"П_КВ_Площадь_Общая_К",
						"П_КВ_Площадь_Отапливаемые"]
			params_elements_type = ["Area",
						"Area",
						"Area",
						"Area",
						"Area",
						"Area",
						"Area",
						"Area"]
			params_elements_SetAllowVaryBetweenGroups = [True,
						True,
						True,
						True,
						True,
						True,
						True,
						True]
			params_found = []
			common_parameters_file = "Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\lib\\ФОП_Scripts.txt"
			app = doc.Application
			#ROOMS CATEGORY SET
			category_set_rooms = app.Create.NewCategorySet()
			insert_cat_rooms = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Rooms)
			category_set_rooms.Insert(insert_cat_rooms)
			#SHARED PARAMETERS FILE
			app.SharedParametersFilename = common_parameters_file
			SharedParametersFile = app.OpenSharedParameterFile()
			#CHECK CURRENT PARAMETERS
			map = doc.ParameterBindings
			it = map.ForwardIterator()
			it.Reset()
			while it.MoveNext():
				d_Definition = it.Key
				d_Name = it.Key.Name
				d_Binding = it.Current
				d_catSet = d_Binding.Categories		
				for i in range(0, len(params_elements)):
					if d_Name == params_elements[i]:
						if d_Binding.GetType() == DB.InstanceBinding:
							if str(d_Definition.ParameterType) == params_elements_type[i]:
								if d_Definition.VariesAcrossGroups == params_elements_SetAllowVaryBetweenGroups[i]:
									if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Rooms)):
										params_found.append(d_Name)
			with db.Transaction(name = "AddSharedParameter"):
				for dg in SharedParametersFile.Groups:
					if dg.Name == group:
						for item in params_elements:
							if not in_list(item, params_found):
								externalDefinition = dg.Definitions.get_Item(item)
								newIB = app.Create.NewInstanceBinding(category_set_rooms)
								doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
								doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
			map = doc.ParameterBindings
			it = map.ForwardIterator()
			it.Reset()
			with db.Transaction(name = "SetAllowVaryBetweenGroups"):
				while it.MoveNext():
					for i in range(0, len(params_elements)):
						if not in_list(params_elements[i], params_found):
							d_Definition = it.Key
							d_Name = it.Key.Name
							d_Binding = it.Current
							if d_Name == params_elements[i]:
								d_Definition.SetAllowVaryBetweenGroups(doc, params_elements_SetAllowVaryBetweenGroups[i])
		except Exception as e:
			Alert("Ошибка при загрузке параметров:\n[{}]".format(str(e)), title="Загрузчик параметров", header = "Ошибка")
else:
	Alert("Файл общих параметров не найден:\nZ:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\lib\\ФОП_Scripts.txt", title="Загрузчик параметров", header = "Ошибка")
