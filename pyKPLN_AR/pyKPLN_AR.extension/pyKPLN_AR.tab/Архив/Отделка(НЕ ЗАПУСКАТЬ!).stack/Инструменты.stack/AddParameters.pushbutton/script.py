# -*- coding: utf-8 -*-
"""
FW_SharedParameters

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Параметры"
__doc__ = 'Загрузка всех необходимых параметров для работы с отделкой в проект.\n' \
          'ФОП: Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt' \

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

if os.path.exists("Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt"):
	commands = [CommandLink('Да', return_value=True), CommandLink('Отмена', return_value=False)]
	dialog = TaskDialog('Добавить параметры в проект?',
						title = "Загрузчик параметров",
						title_prefix=False,
						content="ФОП - Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt",
						commands=commands,
						footer='См. правила работы с отделкой',
						show_close=False)
	result = dialog.show()
	if result:
		try:
			openings_params = ["О_Проем_ПОМ_Внутри_Id помещения", "О_Проем_ПОМ_Снаружи_Id помещения"]
			openings_SetAllowVaryBetweenGroups = [True, True]
			group = "АРХИТЕКТУРА - Отделка"
			elements_params = ["О_Id помещения", "О_Имя помещения", "О_Назначение помещения", "О_Номер помещения", "О_Номер секции", "О_Тип", "О_Группа"]
			elements_params_found = []
			elements_SetAllowVaryBetweenGroups = [True, True, True, True, True, True, True, True, True]
			room_params = ["О_ПОМ_Площадь стен", 
				  "О_ПОМ_Описание стен", 
				  "О_ПОМ_Площадь полов", 
				  "О_ПОМ_Описание полов", 
				  "О_ПОМ_Площадь потолков", 
				  "О_ПОМ_Описание потолков", 
				  "О_ПОМ_Площадь стен_Текст", 
				  "О_ПОМ_Площадь полов_Текст", 
				  "О_ПОМ_Площадь потолков_Текст", 
				  "О_ПОМ_Описание плинтусов", 
				  "О_ПОМ_Длина плинтусов_Текст", 
				  "О_ПОМ_ГОСТ_Описание стен", 
				  "О_ПОМ_ГОСТ_Описание плинтусов",
				  "О_ПОМ_ГОСТ_Описание полов",
				  "О_ПОМ_ГОСТ_Описание потолков", 
				  "О_ПОМ_ГОСТ_Площадь стен_Текст",
				  "О_ПОМ_ГОСТ_Площадь потолков_Текст",
				  "О_ПОМ_ГОСТ_Площадь полов_Текст", 
				  "О_ПОМ_ГОСТ_Длина плинтусов_Текст",
				  "О_ПОМ_Группа"]
			room_params_type = ["Area", 
					   "MultilineText", 
					   "Area", 
					   "MultilineText", 
					   "Area", 
					   "MultilineText", 
					   "Text", 
					   "Text", 
					   "Text", 
					   "MultilineText", 
					   "Text", 
					   "MultilineText", 
					   "MultilineText",
					   "MultilineText",
					   "MultilineText", 
					   "MultilineText", 
					   "MultilineText", 
					   "MultilineText", 
					   "MultilineText", 
					   "Text"]
			params_found = []
			category_lost = []
			elements_params_type = ["О_Описание", "О_Плинтус", "О_Плинтус_Высота"]
			elements_params_type_type = ["MultilineText", "YesNo", "Length"]
			room_SetAllowVaryBetweenGroups = [True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True]
			common_parameters_file = "Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt"
			app = doc.Application
			category_set_openings = app.Create.NewCategorySet()
			insert_cat_openings = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Doors)
			category_set_openings.Insert(insert_cat_openings)
			insert_cat_openings = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Windows)
			category_set_openings.Insert(insert_cat_openings)
			insert_cat_openings = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_MechanicalEquipment)
			category_set_openings.Insert(insert_cat_openings)
			category_set_elements = app.Create.NewCategorySet()
			insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Walls)
			category_set_elements.Insert(insert_cat_elements)
			insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Floors)
			category_set_elements.Insert(insert_cat_elements)
			insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Ceilings)
			category_set_elements.Insert(insert_cat_elements)
			originalFile = app.SharedParametersFilename
			category_set_rooms = app.Create.NewCategorySet()
			insert_cat_rooms = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Rooms)
			category_set_rooms.Insert(insert_cat_rooms)
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
				for i in range(0, len(openings_params)):
					if d_Name == openings_params[i]:
						if d_Binding.GetType() == DB.InstanceBinding:
							if str(d_Definition.ParameterType) == "Text":
								if d_Definition.VariesAcrossGroups == elements_SetAllowVaryBetweenGroups[i]:
									if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Doors)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Windows)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_MechanicalEquipment)):
										params_found.append(d_Name)
				for i in range(0, len(elements_params)):
					if d_Name == elements_params[i]:
						if d_Binding.GetType() == DB.InstanceBinding:
							if str(d_Definition.ParameterType) == "Text":
								if d_Definition.VariesAcrossGroups == elements_SetAllowVaryBetweenGroups[i]:
									if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Walls)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Floors)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Ceilings)):
										params_found.append(d_Name)
				for i in range(0, len(room_params)):
					if d_Name == room_params[i]:
						if d_Binding.GetType() == DB.InstanceBinding:
							if str(d_Definition.ParameterType) == room_params_type[i]:
								if d_Definition.VariesAcrossGroups == room_SetAllowVaryBetweenGroups[i]:
									if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Rooms)):
										params_found.append(d_Name)	
				for i in range(0, len(elements_params_type)):
					if d_Name == elements_params_type[i]:
						if d_Binding.GetType() == DB.TypeBinding:
							if str(d_Definition.ParameterType) == elements_params_type_type[i]:
								if not d_Definition.VariesAcrossGroups:
									if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Walls)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Floors)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Ceilings)):
										params_found.append(d_Name)
			with db.Transaction(name = "AddSharedParameter"):
				for dg in SharedParametersFile.Groups:
					if dg.Name == group:
						for it in openings_params:
							if not in_list(it, params_found):
								externalDefinition = dg.Definitions.get_Item(it)
								newIB = app.Create.NewInstanceBinding(category_set_openings)
								doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
								doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
						for it in elements_params:
							if not in_list(it, params_found):
								externalDefinition = dg.Definitions.get_Item(it)
								newIB = app.Create.NewInstanceBinding(category_set_elements)
								doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
								doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
						for it in room_params:
							if not in_list(it, params_found):
								externalDefinition = dg.Definitions.get_Item(it)
								newIB = app.Create.NewInstanceBinding(category_set_rooms)
								doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
								doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
						for it in elements_params_type:
							if not in_list(it, params_found):
								externalDefinition = dg.Definitions.get_Item(it)
								newIB = app.Create.NewTypeBinding(category_set_elements)
								doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
								doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)

			map = doc.ParameterBindings
			it = map.ForwardIterator()
			it.Reset()
			with db.Transaction(name = "SetAllowVaryBetweenGroups"):
				while it.MoveNext():
					for i in range(0, len(openings_params)):
						if not in_list(openings_params[i], params_found):
							d_Definition = it.Key
							d_Name = it.Key.Name
							d_Binding = it.Current
							if d_Name == openings_params[i]:
								d_Definition.SetAllowVaryBetweenGroups(doc, openings_SetAllowVaryBetweenGroups[i])
					for i in range(0, len(elements_params)):
						if not in_list(elements_params[i], params_found):
							d_Definition = it.Key
							d_Name = it.Key.Name
							d_Binding = it.Current
							if d_Name == elements_params[i]:
								d_Definition.SetAllowVaryBetweenGroups(doc, elements_SetAllowVaryBetweenGroups[i])
					for i in range(0, len(room_params)):
						if not in_list(room_params[i], params_found):
							d_Definition = it.Key
							d_Name = it.Key.Name
							d_Binding = it.Current
							if d_Name == room_params[i]:
								d_Definition.SetAllowVaryBetweenGroups(doc, room_SetAllowVaryBetweenGroups[i])
		except:
			Alert("Ошибка в файле общих параметров (отсутствует один или несколько параметров)", title="Загрузчик параметров", header = "Ошибка")
else:
	Alert("Файл общих параметров не найден:\nZ:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt", title="Загрузчик параметров", header = "Ошибка")