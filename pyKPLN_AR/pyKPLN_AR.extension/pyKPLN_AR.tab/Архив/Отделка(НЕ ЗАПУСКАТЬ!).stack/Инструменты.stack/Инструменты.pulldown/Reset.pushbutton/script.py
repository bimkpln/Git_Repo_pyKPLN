# -*- coding: utf-8 -*-
"""
FW_Reset

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Сбросить значения"
__doc__ = 'Обнуление параметров отделки в помещенияэ и элементах отделки' \

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
from rpw.ui.forms import CommandLink, TaskDialog, Alert

rooms = []
elements = []

params_et = ["О_Имя помещения", "О_Назначение помещения", "О_Id помещения", "О_Номер помещения"]
params_ra = ["О_ПОМ_Площадь стен", "О_ПОМ_Площадь полов", "О_ПОМ_Площадь потолков"]
params_rt = ["О_ПОМ_Описание стен","О_ПОМ_Описание полов","О_ПОМ_Описание потолков", "О_ПОМ_Площадь стен_Текст", "О_ПОМ_Площадь полов_Текст", "О_ПОМ_Площадь потолков_Текст"]
commands = [CommandLink('Да', return_value='Yes'), CommandLink('Отмена', return_value='Cancel')]
dialog = TaskDialog('Сбросить все значения параметров в стенах и помещениях?',
					title = "Сброс параметров",
					title_prefix=False,
					content="Примечание: все параметры отделки кроме «О_Тип»",
					commands=commands,
					footer='См. правила работы с отделкой',
					show_close=False)
result = dialog.show()

if result != 'Cancel':
	for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
		rooms.append(room)
	for element in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements():
		elements.append(element)
	for element in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements():
		elements.append(element)
	for element in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements():
		elements.append(element)
	with db.Transaction(name = "write_nulls"):
		for room in rooms:
			for par in params_ra:
				try:
					parameter = room.LookupParameter(par)
					parameter.Set(0)
				except: pass
			for par in params_rt:
				try:
					parameter = room.LookupParameter(par)
					parameter.Set("")
				except: pass
		for element in elements:
			for par in params_et:
				try:
					parameter = element.LookupParameter(par)
					parameter.Set("")
				except: pass