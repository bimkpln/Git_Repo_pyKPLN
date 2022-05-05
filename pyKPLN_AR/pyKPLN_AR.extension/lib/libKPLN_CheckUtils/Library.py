# -*- coding: utf-8 -*-
import math
import time
import datetime
import os
import re

from rpw import doc, uidoc, DB, UI, db, ui, revit as Revit
from pyrevit import script
from pyrevit import forms
from pyrevit import revit, DB, UI
from pyrevit.framework import clr
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *

import System
from System.Windows.Forms import *
from System.Drawing import *
from System import Guid
 

class NamingRule():
	def __init__(self, parts):
		self.NamingParts = parts
		self.Reports = []
		self.Examples = []
		for part in self.NamingParts:
			self.Examples.append(part.Example)
		self.Example = "_".join(self.Examples)

	@property
	def Reports(self):
		return self.Reports

	def Check(self, value_list, element):
		self.Reports = []
		if len(value_list) == len(self.NamingParts) or len(value_list) == len(self.NamingParts)-1:
			for i in range(0, len(value_list)):
				if self.NamingParts[i].IsError(value_list[i], element):
					self.Reports.append(NamingReport(self.NamingParts[i].Header, self.NamingParts[i].Description, value_list[i], self.NamingParts[i].Example, True))
				else:
					self.Reports.append(NamingReport("Без предупреждений", "При проверке проблем с записью не обнаружено", value_list[i], self.NamingParts[i].Example, False))
		else:
			self.Reports.append(NamingReport("Невозможно инициировать проверку", "Недопустимое количество данных", None, self.Example, True))
		return self.Reports

class NamingReport():
	def __init__(self, header, description, part, example, is_error):
		self.Header = header
		self.Description = description
		self.Part = part
		self.Example = example
		self.IsError = is_error

	@property
	def IsError(self):
		return self.IsError

	@property
	def Header(self):
		return self.Header

	@property
	def Description(self):
		return self.Description

	@property
	def Part(self):
		return self.Part

	@property
	def Example(self):
		return self.Example

class CheckValue():
	def __init__(self, value, category, rule, element, name):
		self.Element = element
		self.Name = name
		self.Category = category
		self.ValueString = value
		self.NamingRule = rule
		self.ValueList = self.GetParts(self.ValueString.split("_"))
		self.Errors = []

	@property
	def Name(self):
		return self.Name

	@property
	def Element(self):
		return self.Element

	@property
	def Errors(self):
		return self.Errors

	@property
	def ValueList(self):
		return self.ValueList

	@property
	def Category(self):
		return self.Category

	def GetResult(self, element):
		return self.NamingRule.Check(self.ValueList, element)

	def GetParts(self, value_list):
		parts = []
		opened = False
		opened_k = 0
		temp = ""
		for i in range(0, len(value_list)):
			if "(" in value_list[i] and ")" in value_list[i]:
				parts.append(value_list[i])
			elif "(" in value_list[i] and not ")" in value_list[i]:
				if opened:
					temp += "_"
				temp += value_list[i]
				opened_k += 1
				opened = True
			elif ")" in value_list[i]:
				opened_k -= 1
				temp += "_" + value_list[i]
				if opened_k == 0:
					parts.append(temp)
					temp = ""
					opened = False
			else:
				if opened:
					temp += "_" + value_list[i]
				else:
					parts.append(value_list[i])
			if i == len(value_list)-1:
				if temp != "":
					parts.append(temp)
					temp = ""
		perfect_parts = []
		for part in parts:
			if part != "":
				perfect_parts.append(part)
		return perfect_parts

class NamingPart():
	@property
	def Example(self):
		return self.Example

	@property
	def Description(self):
		return self.Description

	@property
	def Header(self):
		return self.Header

	def IsInteger(self, value):
		try:
			num = int(value)
			return True
		except:
			return False

	def IsFloat(self, value):
		try:
			if '.' in value:
				num = float(value)
				return True
			else:
				return False
		except:
			return False

	def OnlyNumbers(self, value):
		for i in value:
			if not i in "0123456789":
				return False
		return True

	def NoSpaces(self, value):
		cleared = ""
		for c in value:
			if not c in " \n\t":
				cleared += c
		return cleared

class NP_AnyText(NamingPart):
	def __init__(self):
		self.Header = "Содержит недопустимые символы"
		self.Example = "Необязательное примечание"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		self.Header = "Содержит недопустимые символы"
		if not self.IsInteger(value):
			if "(" in value or ")" in value or "_" in value or "." in value or ",":
				return False
			return True
		else:
			self.Header = "Не должно быть числом"
			return True

class NP_AnyTextDescription(NamingPart):
	def __init__(self):
		self.Header = "Содержит недопустимые символы"
		self.Example = "Обязательное описание"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		if "(" in value or ")" in value or "_" in value or "." in value or ",":
			return False
		return True

class NP_AnyNumber(NamingPart):
	def __init__(self):
		self.Header = "Не является целым числом"
		self.Example = "13"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		if self.IsInteger(value):
			return False
		return True

class NP_AnyNumberAdd(NamingPart):
	def __init__(self):
		self.Header = "Не является числом"
		self.Example = "-01"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		if value[0] == "-":
			if self.IsInteger(value[1:]):
				return False
			return True
		elif self.IsInteger(value):
			return False
		else:
			return True

class NP_AnyFloat(NamingPart):
	def __init__(self):
		self.Header = "Не является дробным числом"
		self.Example = "10.5"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		if self.IsFloat(value):
			return False
		return True

class NP_Code_Common(NamingPart):
	def __init__(self):
		self.Header = "Не соответствует формату ##, где # - целое число"
		self.Example = "03"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		if self.OnlyNumbers(value) and len(value) > 1 and len(value) < 4:
			return False
		return True

class NP_Code_Stained(NamingPart):
	def __init__(self):
		self.Header = "Не соответствует формату В-#, где # - маркировка типоразмера витража"
		self.Example = "В-3"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		self.Header = "Не соответствует формату В-#, где # - маркировка типоразмера витража"
		if value.lower().startswith("в-"):
			if self.OnlyNumbers(value[2:]) and len(value[2:]) >= 1 and len(value[2:]) < 4:
				if value[2:] != element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString():
					self.Header = "Обозначенная маркировка типоразмера не соответствует фактическому значению"
					return False
				else:
					return True
			else:
				return True
		return True

class NP_Level(NamingPart):
	def __init__(self):
		self.Header = "Не соответствует формату ЧП(#.###), где #.### - высота уровня"
		self.Example = "ЧП(+0.000)"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		return False
		#Добавить проверку наименования уровня

class NP_Code_Department(NamingPart):
	def __init__(self):
		self.Header = "Не соответствует формату ??.#, где ?? - код раздела, # - код вида"
		self.Example = "АР.В"
		self.Description = "Недопустимое значение"
		self.dep_list = ["АР", "КР", "ОВ", "ВК", "ЭОМ", "СС", "ПБ", "ТХ", "КЖ", "КМ", "ИС", "ГП", "ТС", "ПОС", "НВК", "КФ", "МПБ", "ООС", "ИОС", "ЭЭ", "ПЗУ", "ООС", "АК", "ОДИ", "SYS"]
		self.code_list = ["В", "О", "Э", "ЗВ", "ЗИ", "К", "И", "SYS"]

	def IsError(self, value, element=None):
		self.Header = "Не соответствует формату ??.#, где ?? - код раздела, # - код вида"
		if value.count(".") == 1:
			values = value.split(".")
			if (values[0].upper() in self.code_list and values[1].upper() in self.dep_list) or (values[1].upper() in self.code_list and values[0].upper() in self.dep_list):
				self.Header = "Неизвестные значения в полях «Код вида» (доступные значения: {}), «Код раздела» (доступные значения: {})".format(", ".join(self.dep_list), ", ".join(self.code_list))
				return False
			return True
		else:
			return True

class NP_Code_State(NamingPart):
	def __init__(self):
		self.value_list = ["ЭП", "ПД", "П", "РД", "Р", "Т", "КОН", "В", "К", "SYS"]
		self.Header = "Не соответствует ни одному из доступных значений: {}".format(", ".join(self.value_list))
		self.Example = "ПД"
		self.Description = "Недопустимое значение"


	def IsError(self, value, element=None):
		if value.upper() in self.value_list:
			return False
		else:
			return True

class NP_Width(NamingPart):
	def __init__(self):
		self.Header = "Не соответствует формату ##мм, где ## - ширина стены в миллиметрах, округленная до целого значания"
		self.Example = "100мм"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		self.Header = "Не соответствует формату ##мм, где ## - ширина стены в миллиметрах, округленная до целого значания"
		if len(self.NoSpaces(value)) > 3:
			if self.NoSpaces(value)[-2:] == "мм":
				if self.IsInteger(self.NoSpaces(value)[:-2]):
					try:
						if str(float(self.NoSpaces(value)[:-2]) / 304.8) != str(element.GetCompoundStructure().GetWidth()):
							self.Header = "Обозначенная толщина стены не равна фактическому значению"
							return True
						else:
							return False
					except: return False
				else:
					return True
			else:
				return True
		else:
			return True

class NP_LayersPlus(NamingPart):
	def __init__(self):
		self.Header = "Не соответствует формату ВН(Сс-##_Сс-##), где ВН - расположение стены (НА - наружные / ВН - внутренние слои), Сс - шифр слоя, ## - толщина слоя"
		self.Example = "ВН(Шт-10_Ут-150)"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		self.Header = "Не соответствует формату ВН(Сс-##_Сс-##), где ВН - расположение стены (НА - наружные / ВН - внутренние слои), Сс - шифр слоя, ## - толщина слоя"
		common_width = 0
		if self.NoSpaces(value)[:2].upper() == "НА" or self.NoSpaces(value)[:2].upper() == "ВН":
			if self.NoSpaces(value)[2] == "(" and self.NoSpaces(value)[-1] == ")" and self.NoSpaces(value).count("(") == 1 and self.NoSpaces(value).count(")") == 1:
				parts = self.NoSpaces(value)[3:-1].split("_")
				for part in parts:
					if "-" in part:
						code = part.split("-")[0]
						width = part.split("-")[1]
						if self.IsInteger(width):
							width = int(width)
							common_width += width
						else:
							return True
					else:
						return True
				try:
					if str(common_width / 304.8) != str(element.GetCompoundStructure().GetWidth()):
						self.Header = "Сумма толщин обозначенных слоев не равна итоговому значению толщины элемента"
						return True
					else:
						return False
				except: return False
			else:
				return True
		else:
			return True

class NP_Layers(NamingPart):
	def __init__(self):
		self.Header = "Не соответствует формату (Сс-##_Сс-##),  Сс - шифр слоя, ## - толщина слоя"
		self.Example = "(Шт-10_Ут-150)"
		self.Description = "Недопустимое значение"

	def IsError(self, value, element=None):
		self.Header = "Не соответствует формату (Сс-##_Сс-##), где  Сс - шифр слоя, ## - толщина слоя"
		common_width = 0
		if self.NoSpaces(value)[0] == "(" and self.NoSpaces(value)[-1] == ")" and self.NoSpaces(value).count("(") == 1 and self.NoSpaces(value).count(")") == 1:
			parts = self.NoSpaces(value)[0:-1].split("_")
			for part in parts:
				if "-" in part:
					code = part.split("-")[0]
					width = part.split("-")[1]
					if self.IsInteger(width):
						width = int(width)
						common_width += width
					else:
						return True
				else:
					return True
			try:
				if str(common_width / 304.8) != str(element.GetCompoundStructure().GetWidth()):
					self.Header = "Сумма толщин обозначенных слоев не равна итоговому значению толщины элемента"
					return True
				else:
					return False
			except: return False
		else:
			return True

