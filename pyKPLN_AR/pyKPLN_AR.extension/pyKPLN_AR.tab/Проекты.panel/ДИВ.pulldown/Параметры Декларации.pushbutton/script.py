# -*- coding: utf-8 -*-
"""
SHEETS_SharedParameters

"""
__author__ = 'Anastasiya Stepanova'
__title__ = "Параметры для декларации"
__doc__ = 'Загрузка всех необходимых параметров для работы с декларацией в проект.\n' \
          'ФОП: Y:\Жилые здания\ДИВНОЕ СитиXXI\7. Стадия РД\Рабочие материалы\АР\Скрипты\[BETA]_Declaration\Source\DEC_Sheet' \
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

if os.path.exists("Y:\\Жилые здания\\ДИВНОЕ СитиXXI\\7. Стадия РД\\Рабочие материалы\\АР\\Скрипты\\[BETA]_Declaration\\Source\\DEC_Sheet.txt"):
	commands = [CommandLink('Да', return_value=True), CommandLink('Отмена', return_value=False)]
	dialog = TaskDialog('Добавить параметры в проект?',
						title = "Загрузчик параметров",
						title_prefix  =False,
						content = "Y:\\Жилые здания\\ДИВНОЕ СитиXXI\\7. Стадия РД\\Рабочие материалы\\АР\\Скрипты\\[BETA]_Declaration\\Source\\DEC_Sheet.txt",
						commands = commands,
						footer = ' ',
						show_close = False)
	result = dialog.show()
	if result:
		try:
			group = "Sheet_DEC"
			params_elements = ["Sheet_DEC_Row_01",
						"Sheet_DEC_Row_02",
						"Sheet_DEC_Row_03",
						"Sheet_DEC_Row_04",
						"Sheet_DEC_Row_05",
						"Sheet_DEC_Row_06",
						"Sheet_DEC_Row_07",
						"Sheet_DEC_Row_08",
						"Sheet_DEC_Row_09",
						"Sheet_DEC_Row_10",
						"Sheet_DEC_Row_Name_01",
						"Sheet_DEC_Row_Name_02",
						"Sheet_DEC_Row_Name_03",
						"Sheet_DEC_Row_Name_04",
						"Sheet_DEC_Row_Name_05",
                        "Sheet_DEC_Row_Name_06",
                        "Sheet_DEC_Row_Name_07",
                        "Sheet_DEC_Row_Name_08",
                        "Sheet_DEC_Row_Name_09",
                        "Sheet_DEC_Row_Name_10",
                        "Sheet_DEC_Row_S_01",
                        "Sheet_DEC_Row_S_02",
                        "Sheet_DEC_Row_S_03",
                        "Sheet_DEC_Row_S_04",
                        "Sheet_DEC_Row_S_05",
                        "Sheet_DEC_Row_S_06",
                        "Sheet_DEC_Row_S_07",
                        "Sheet_DEC_Row_S_08",
                        "Sheet_DEC_Row_S_09",
                        "Sheet_DEC_Row_S_10",
                        "Sheet_DEC_S",
                        "Sheet_DEC_Title"]
			params_elements_type = ["Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
						"Text",
                        "Text",
                        "Text",
                        "Text",
                        "Text",
                        "Text",
                        "Area",
                        "Area",
                        "Area",
                        "Area",
                        "Area",
                        "Area",
                        "Area",
                        "Area"
                        "Area",
                        "Area",
                        "Area",
                        "Text"]
			params_elements_SetAllowVaryBetweenGroups = [True,
						True,
						True,
						True,
						True,
						True,
						True,
						True,
						True,
						True,
						True,
						True,
						True,
						True,
						True,
                        True,
                        True,
                        True,
                        True,
                        True,
                        True,
						True,
						True,
						True,
						True,
						True,
                        True,
						True,
						True,
						True,
						True,
                        True,
						True]
			params_found = []
			common_parameters_file = "Y:\\Жилые здания\\ДИВНОЕ СитиXXI\\7. Стадия РД\\Рабочие материалы\\АР\\Скрипты\\[BETA]_Declaration\\Source\\DEC_Sheet.txt"
			app = doc.Application
			#ROOMS CATEGORY SET
			category_set_rooms = app.Create.NewCategorySet()
			insert_cat_rooms = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Sheets)
			category_set_rooms.Insert(insert_cat_rooms)
			#PROJECT INFORMATION CATEGORY SET
			category_set_project_info = app.Create.NewCategorySet()
			insert_cat_project_info = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_ProjectInformation)
			category_set_project_info.Insert(insert_cat_project_info)
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
									if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Sheets)):
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
	Alert("Файл общих параметров не найден:Y:\Жилые здания\ДИВНОЕ СитиXXI\7. Стадия РД\Рабочие материалы\АР\Скрипты\[BETA]_Declaration\Source\DEC_Sheet", title="Загрузчик параметров", header = "Ошибка")
