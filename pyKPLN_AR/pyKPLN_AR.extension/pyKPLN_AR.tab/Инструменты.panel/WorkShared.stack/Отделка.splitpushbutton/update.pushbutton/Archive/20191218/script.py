# -*- coding: utf-8 -*-
"""
FW_Update

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Обновить"
__doc__ = 'Обновляет значение площади отделки для стен, плинтусов, потолков и полов.\n\n' \
          '«О_ПОМ_Описание стен», «О_ПОМ_Описание полов», «О_ПОМ_Описание потолков», «О_ПОМ_Описание плинтусов» - Текстовые параметры для определения состава отделки.\nЕсли стоит значение «Без отделки», то при расчете без моделирования помещений площадь будет расчитана как «0.00»\n\n' \
          '«О_ПОМ_Площадь стен», «О_ПОМ_Площадь полов», «О_ПОМ_Площадь потолков» - Параметры для записи общей площади отделки\n\n' \
		  '«О_ПОМ_Площадь стен_Текст», «О_ПОМ_Площадь полов_Текст», «О_ПОМ_Площадь потолков_Текст», «О_ПОМ_Длина плинтусов_Текст» - Параметры для записи площади отделки (текст с разделением по типоразмерам и итоговой площадью)' \

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
import System
from System.Windows.Forms import *
from System.Drawing import *
import datetime
from rpw.ui.forms import CommandLink, TaskDialog, Alert

errors = ""

class SortParameter(Form):
	def __init__(self):
		global group_by
		self.Name = "KPLN_Finishing"
		self.Text = "KPLN Параметры группировки помещений"
		self.Size = Size(300, 200)
		self.MinimumSize = Size(300, 200)
		self.MaximumSize = Size(300, 200)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = False
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		self.TopMost = True
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")

		self.button_ok = Button()
		self.button_pass = Button()
		self.label_faq = Label()
		self.cb1 = ComboBox()
		self.cb2 = ComboBox()
		self.cb3 = ComboBox()
		self.cb1.Items.Add("<неактивен>")
		self.cb2.Items.Add("<неактивен>")
		self.cb3.Items.Add("<неактивен>")
		self.cb1.Text = "<неактивен>"
		self.cb2.Text = "<неактивен>"
		self.cb3.Text = "<неактивен>"

		self.cb1.Parent = self
		self.cb1.Location = Point(12, 54)
		self.cb1.Size = Size(260, 21)
		self.cb1.DropDownStyle = ComboBoxStyle.DropDownList
		self.cb1.SelectedIndexChanged += self.CbOnChanged

		self.cb2.Parent = self
		self.cb2.Location = Point(12, 79)
		self.cb2.Size = Size(260, 21)
		self.cb2.DropDownStyle = ComboBoxStyle.DropDownList
		self.cb2.SelectedIndexChanged += self.CbOnChanged
		self.cb2.Enabled = False

		self.cb3.Parent = self
		self.cb3.Location = Point(12, 104)
		self.cb3.Size = Size(260, 21)
		self.cb3.DropDownStyle = ComboBoxStyle.DropDownList
		self.cb3.SelectedIndexChanged += self.CbOnChanged
		self.cb3.Enabled = False

		self.button_ok.Parent = self
		self.button_ok.Location = Point(12, 131)
		self.button_ok.Size = Size(75, 23)
		self.button_ok.Text = "Применить"
		self.button_ok.Click += self.OnOk

		self.button_pass.Parent = self
		self.button_pass.Location = Point(93, 131)
		self.button_pass.Size = Size(75, 23)
		self.button_pass.Text = "Отмена"
		self.button_pass.Click += self.OnPass

		self.label_faq.Parent = self
		self.label_faq.Location = Point(12, 13)
		self.label_faq.Size = Size(260, 38)
		self.label_faq.Text = "Определите параметр помещений для группирования:"

		room = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement()
		parameters = []
		for j in room.Parameters:
			try:
				if str(j.StorageType) == "String":
					parameters.append(j.Definition.Name)
			except: pass
		parameters.sort()
		for p in parameters:
			self.cb1.Items.Add(p)
			self.cb2.Items.Add(p)
			self.cb3.Items.Add(p)
		self.CenterToScreen()

	def CbOnChanged(self, sender, event):
		if self.cb1.Text != "<неактивен>":
			self.cb2.Enabled = True
		else:
			self.cb2.Enabled = False
			self.cb2.Text = "<неактивен>"
		if self.cb2.Text != "<неактивен>":
			self.cb3.Enabled = True
		else:
			self.cb3.Enabled = False
			self.cb3.Text = "<неактивен>"

	def OnPass(self, sender, args):
		global group_by
		group_by = []
		self.Close()

	def OnOk(self, sender, args):
		global group_by
		if self.cb1.Text != "<неактивен>":
			group_by.append(self.cb1.Text)
		if self.cb2.Text != "<неактивен>":
			group_by.append(self.cb2.Text)
		if self.cb3.Text != "<неактивен>":
			group_by.append(self.cb3.Text)
		self.Close()

class PickParameter(Form):
	def __init__(self):
		global custom_parameter
		self.Name = "KPLN_Finishing"
		self.Text = "KPLN Нестандартный параметр"
		self.Size = Size(300, 150)
		self.MinimumSize = Size(300, 150)
		self.MaximumSize = Size(300, 150)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = False
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		self.TopMost = True
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")

		self.button_ok = Button()
		self.button_pass = Button()
		self.label_faq = Label()
		self.cb = ComboBox()

		self.cb.Parent = self
		self.cb.Location = Point(12, 54)
		self.cb.Size = Size(260, 21)
		self.cb.DropDownStyle = ComboBoxStyle.DropDownList
		self.cb.SelectedIndexChanged += self.CbOnChanged

		self.button_ok.Parent = self
		self.button_ok.Location = Point(12, 81)
		self.button_ok.Size = Size(75, 23)
		self.button_ok.Text = "Применить"
		self.button_ok.Click += self.OnOk
		self.button_ok.Enabled = False

		self.button_pass.Parent = self
		self.button_pass.Location = Point(93, 81)
		self.button_pass.Size = Size(179, 23)
		self.button_pass.Text = "Использовать стандартный №"
		self.button_pass.Click += self.OnPass

		self.label_faq.Parent = self
		self.label_faq.Location = Point(12, 13)
		self.label_faq.Size = Size(260, 38)
		self.label_faq.Text = "Определите параметр помещений для расчета:"

		room = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement()
		parameters = []
		for j in room.Parameters:
			try:
				if j.IsShared and str(j.StorageType) == "String":
					parameters.append(j.Definition.Name)
			except: pass
		parameters.sort()
		for p in parameters:
			self.cb.Items.Add(p)

		self.CenterToScreen()

	def CbOnChanged(self, sender, event):
		if self.cb.Text != "":
			self.button_ok.Enabled = True

	def OnPass(self, sender, args):
		global custom_parameter
		custom_parameter = "[USE DEFAULT]"
		self.Close()

	def OnOk(self, sender, args):
		global custom_parameter
		custom_parameter = self.cb.Text
		self.Close()

def AddReport(element, description):
	global errors
	global output
	if element != False:
		link_id = output.linkify(element.Id)
		test = str(element.Id.ToString()) + "_" + description
		if not(test in errors):
			errors += test
			output.print_html("{} : {}<br><br>".format(link_id, description))
	else:
		test = description
		if not(test in errors):
			errors += test
			output.print_html("{}<br><br>".format(description))

def check_def_params(element):
	for param in ["О_Имя помещения", "О_Id помещения", "О_Номер помещения"]:
		if str(element.LookupParameter(param).AsString()) == "" or str(element.LookupParameter(param).AsString()) == "Null" or str(element.LookupParameter(param).AsString()) == "None":
			return False
	return True

def in_list(element, list):
	for i in list:
		if element == i:
			return True
	return False

def get_index(element, list):
	if in_list(element, list):
		for i in range(0, len(list)):
			if element == list[i]:
				return i
	return 0

def normalized_string(string):
	if string.lower() == "none" or string == "":
		return "<НЕЗАПОЛНЕНО!>"
	text = string
	while text.startswith("\n") or text.startswith("\t"):
		text = text[2:]
	while text.endswith("\n") or text.endswith("\t"):
		text = text[:(len(text)-2)]
	return text

def append_dict(types, areas, room, type_group, sort_key):
	for type_group in type_groups:
		if type_group[0] == types and type_group[3] == sort_key:
			for z in range(0, len(areas)):
				type_group[1][z] += areas[z]
			type_group[2].append(room)
			return
	type_groups.append([types, areas, [room], sort_key])
	return

def sort_by_key(list, keys):
	if len(keys) > 1:
		sorted_list = []
		sorted_keys = []
		for i in keys:
			sorted_keys.append(i)
		sorted_keys.sort()
		for i in range(0, len(sorted_keys)):
			for z in range(0, len(sorted_keys)):
				if sorted_keys[i] == keys[z]:
					sorted_list.append(list[z])
		return sorted_list
	else:
		return list

def append_areas(types, type_id, areas, el_area):
	for i in range (0, len(types)):
		if types[i] == type_id:
			areas[i] += el_area

def getType(element):
	try:
		type = element.FloorType
		return type
	except:
		try:
			type = element.WallType
			return type
		except:
			try:
				id = element.GetTypeId()
				type = doc.GetElement(id)
				return type
			except:
				forms.alert("=(")
				return False

def get_value(room, name):
	try:
		for j in room.Parameters:
			if j.Definition.Name == name and str(j.StorageType) == "String":
				value = room.LookupParameter(name).AsString()
				return value
	except: return "<НЕЗАПОЛНЕНО!>"

def room_check_parameters(room):
	list = []
	good = []
	errs = []
	room_params_description = ["О_ПОМ_Описание стен", "О_ПОМ_Описание полов", "О_ПОМ_Описание потолков", "О_ПОМ_Площадь стен_Текст", "О_ПОМ_Площадь полов_Текст", "О_ПОМ_Площадь потолков_Текст", "О_ПОМ_ГОСТ_Описание стен", "О_ПОМ_ГОСТ_Описание потолков", "О_ПОМ_ГОСТ_Площадь стен_Текст", "О_ПОМ_ГОСТ_Площадь потолков_Текст", "О_ПОМ_Группа"]
	room_params_area = ["О_ПОМ_Площадь стен", "О_ПОМ_Площадь полов", "О_ПОМ_Площадь потолков"]
	for i in range(0, 3):
		group = [all_walls, all_floors, all_ceilings][i]
		cat_list = ["[Стены]", "[Полы]", "[Потолки]"]
		cat = cat_list[i]
		for element in group:
			for param in def_params_elements:
				try:
					if str(element.LookupParameter(param).AsString()) == "Null":
						AddReport(element.Id, "{} Отсутствует значение параметра экземпляра «{}»".format(cat, param))
				except:
					AddReport(element.Id, "{} Отсутствует значение параметра «{}»".format(cat, param))
	if result == 'Yes':
		for i in range(0, 3):
			group = [all_walls, all_floors, all_ceilings][i]
			cat_list = ["[Стены]", "[Полы]", "[Потолки]"]
			cat = cat_list[i]
			for element in group:
				try:
					element_type = getType(element)
					if str(element_type.LookupParameter(def_param_description).AsString()) == "Null":
						AddReport(element_type.Id, "{} Отсутствует значение параметра типоразмера «{}»".format(cat, param))
				except:
					AddReport(element.Id, "{} Отсутствует параметр «{}»".format(cat, param))
	for j in room.Parameters:
		for param in room_params_description:
			if j.Definition.Name == param:
				if str(j.StorageType) != "String":
					AddReport(element.Id, "[Помещения] Некорректные еденицы : параметр «{}» != «Текст»".format(param))
					list.append(j.Definition.Name)
				else:
					good.append(j.Definition.Name)
		for param in room_params_area:
			if j.Definition.Name == param:
				if str(j.StorageType) != "Double":
					AddReport(element.Id, "[Помещения] Некорректные еденицы : параметр «{}» != «Площадь»".format(param))
					list.append(j.Definition.Name)
				else:
					good.append(j.Definition.Name)
	for param in room_params_area:
		if not in_list (param, list) and not in_list (param, good):
			AddReport(False, "[Помещения] Отсутствует параметр помещений «{}»".format(param))
	for param in room_params_description:
		if not in_list (param, list) and not in_list (param, good):
			AddReport(False, "[Помещения] Отсутствует параметр помещений «{}»".format(param))
	if len(errs) != 0:
		return errs
	else:
		return True

output = script.get_output()
output.set_title("KPLN : Ошибки в расчете")

rooms = []
walls = []
floors = []
ceilings = []
all_walls = []
all_floors = []
all_ceilings = []

x = 1.0

check_params_rooms = ["О_ПОМ_Площадь стен", "О_ПОМ_Описание стен", "О_ПОМ_Площадь полов", "О_ПОМ_Описание полов", "О_ПОМ_Площадь потолков", "О_ПОМ_Описание потолков", "О_ПОМ_Площадь стен_Текст", "О_ПОМ_Площадь полов_Текст", "О_ПОМ_Площадь потолков_Текст"]

log = ""
def_params_elements_check = ["О_Имя помещения", "О_Назначение помещения", "О_Id помещения", "О_Номер помещения"]
def_params_elements = ["О_Имя помещения", "О_Назначение помещения", "О_Id помещения", "О_Номер помещения"]
def_param_description = "О_Описание"
def_params_rooms = [["О_ПОМ_Площадь стен", "О_ПОМ_Описание стен", "О_ПОМ_Площадь стен_Текст"], ["О_ПОМ_Площадь полов", "О_ПОМ_Описание полов", "О_ПОМ_Площадь полов_Текст"], ["О_ПОМ_Площадь потолков", "О_ПОМ_Описание потолков", "О_ПОМ_Площадь потолков_Текст"]]
def_param_area = DB.BuiltInParameter.HOST_AREA_COMPUTED
def_param_room_p = DB.BuiltInParameter.ROOM_PERIMETER
def_param_room_h = DB.BuiltInParameter.ROOM_HEIGHT
def_param_room_s = DB.BuiltInParameter.ROOM_AREA

rooms_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
walls_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
floors_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
ceilings_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()
log = ""
elements_all = []
collector_doors = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_Doors)
collector_windows = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_Windows)
collector_openings = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
commands = [CommandLink('Да', return_value='Yes'), CommandLink('Нет', return_value='No'), CommandLink('Отмена', return_value='Cancel')]
dialog = TaskDialog('Перезаписать описание слоев в помещениях на значения из элементов?',
					title = "Описание слоев",
					title_prefix=False,
					content="Примечание: параметры помещений «О_ПОМ_Описание стен», «О_ПОМ_Описание полов», «О_ПОМ_Описание потолков»",
					commands=commands,
					footer='См. правила работы с отделкой',
					show_close=False)
result = dialog.show()
commands_2 = [CommandLink('Отделка не смоделирована (брать значения из помещений)', return_value='1'), CommandLink('Смоделирована только отделка стен (остальное брать из помещений)', return_value='2'), CommandLink('Вся отделка смоделирована (полы, стены, потолки)', return_value='3'), CommandLink('Отмена', return_value='Cancel')]
dialog2 = TaskDialog('Выберите тип расчета',
					title = "Тип расчета",
					title_prefix=False,
					content="Если значения из помещений то:\n - S отделки стен = [P помещения] * [H помещения]\n - S отделки полов и потолков = [S помещения]",
					commands=commands_2,
					footer='См. правила работы с отделкой',
					show_close=False)

if result != 'Cancel':
	result_2 = dialog2.show()

commands_x = [CommandLink('0%', return_value=1.0), CommandLink('+15%', return_value=1.15), CommandLink('+25%', return_value=1.25)]
dialogx = TaskDialog('Коэффициент запаса',
					title = "Выберите коэффициент",
					title_prefix=False,
					content="",
					commands=commands_x,
					footer='См. правила работы с отделкой',
					show_close=False)
x = dialogx.show()

if result != 'Cancel' and result_2 != 'Cancel':
	for door in collector_doors:
		try:
			door_name = door.Symbol.FamilyName
			if door_name.startswith("100_"):
				type = doc.GetElement(door.GetTypeId())
				height = type.get_Parameter(DB.BuiltInParameter.DOOR_HEIGHT).AsDouble()
				if height == 0:
					height = type.get_Parameter(DB.BuiltInParameter.GENERIC_HEIGHT).AsDouble()
				width = type.get_Parameter(DB.BuiltInParameter.DOOR_WIDTH).AsDouble()
				if width == 0:
					width = type.get_Parameter(DB.BuiltInParameter.GENERIC_WIDTH).AsDouble()
				inner_id = door.LookupParameter("О_Проем_ПОМ_Внутри_Id помещения").AsString()
				outer_id = door.LookupParameter("О_Проем_ПОМ_Снаружи_Id помещения").AsString()
				if inner_id > "" or outer_id > "":
					if height > 0 and width > 0:
						elements_all.append([inner_id, outer_id, width, height])
					else:
						AddReport(type.Id, "[Двери] Невозможно расчитать параметры «Высота», «Ширина»")
		except: pass
	for window in collector_windows:
		try:
			window_name = window.Symbol.FamilyName
			if not window_name.startswith("120_"):
				type = doc.GetElement(window.GetTypeId())
				height = type.get_Parameter(DB.BuiltInParameter.WINDOW_HEIGHT).AsDouble()
				if height == 0:
					height = type.get_Parameter(DB.BuiltInParameter.GENERIC_HEIGHT).AsDouble()
				width = type.get_Parameter(DB.BuiltInParameter.WINDOW_WIDTH).AsDouble()
				if width == 0:
					width = type.get_Parameter(DB.BuiltInParameter.GENERIC_WIDTH).AsDouble()
				inner_id = window.LookupParameter("О_Проем_ПОМ_Внутри_Id помещения").AsString()
				outer_id = window.LookupParameter("О_Проем_ПОМ_Снаружи_Id помещения").AsString()
				if inner_id > "" or outer_id > "":
					if height > 0 and width > 0:
						elements_all.append([inner_id, outer_id, width, height])
					else:
						AddReport(type.Id, "[Окна] Невозможно расчитать параметры «Высота», «Ширина»")
		except: pass
	for opening in collector_openings:
		try:
			opening_name = opening.Symbol.FamilyName
			if opening_name.startswith("199_"):
				type = doc.GetElement(opening.GetTypeId())
				height = type.get_Parameter(DB.BuiltInParameter.GENERIC_HEIGHT).AsDouble()
				if height == 0:
					height = type.LookupParameter("Высота").AsDouble()
				width = type.get_Parameter(DB.BuiltInParameter.GENERIC_WIDTH).AsDouble()
				if width == 0:
					width = type.LookupParameter("Ширина").AsDouble()
				inner_id = opening.LookupParameter("О_Проем_ПОМ_Внутри_Id помещения").AsString()
				outer_id = opening.LookupParameter("О_Проем_ПОМ_Снаружи_Id помещения").AsString()
				if inner_id > "" or outer_id > "":
					if height > 0 and width > 0:
						elements_all.append([inner_id, outer_id, width, height])
					else:
						AddReport(type.Id, "[Оборудование] Невозможно расчитать параметры «Высота», «Ширина»")
		except: pass

	with db.Transaction(name = "write_nulls"):
		for room in rooms_collector:
			if room.Area != 0.00:
				rooms.append(room)
				for parameter in ["О_ПОМ_Площадь стен", "О_ПОМ_Площадь полов", "О_ПОМ_Площадь потолков"]:
					par = room.LookupParameter(parameter)
					par.Set(0.00)
				for parameter in ["О_ПОМ_ГОСТ_Описание стен", "О_ПОМ_ГОСТ_Описание потолков", "О_ПОМ_ГОСТ_Площадь стен_Текст", "О_ПОМ_ГОСТ_Площадь потолков_Текст", "О_ПОМ_Группа"]:
					par = room.LookupParameter(parameter)
					par.Set("")
	commands_5 = [CommandLink('Да', return_value=True), CommandLink('Нет', return_value=False)]
	dialog5 = TaskDialog('Группировать помещения по параметрам?',
			title = "Группировка",
			title_prefix=False,
			content="",
			commands=commands_5,
			footer='Только для ведомости по ГОСТ',
			show_close=False)
	group_by = []
	result_5 = dialog5.show()
	if result_5:
		form = SortParameter()
		Application.Run(form)
	commands_4 = [CommandLink('Да', return_value=True), CommandLink('Нет', return_value=False)]
	dialog4 = TaskDialog('Расчитывать итоговую площадь?',
			title = "Площадь помещений",
			title_prefix=False,
			content="",
			commands=commands_4,
			footer='',
			show_close=False)
	result_4 = dialog4.show()
	if len(rooms) != 0:
		for wall in walls_collector:
			all_walls.append(wall)
			if check_def_params(wall):
				walls.append(wall)
		for floor in floors_collector:
			all_floors.append(floor)
			if check_def_params(floor):
				floors.append(floor)
		for ceiling in ceilings_collector:
			all_ceilings.append(ceiling)
			if check_def_params(ceiling):
				ceilings.append(ceiling)
		v = room_check_parameters(rooms[0])
		type_groups = []
		if v == True:
			groups = [walls, floors, ceilings]
			with db.Transaction(name = "KPLN: Обновить отделку"):
				for room in rooms:
					sort_by = "key_"
					if len(group_by) != 0:
						for par in group_by:
							sort_by += str(get_value(room, par))
					types = []
					areas = []
					for g in range(0, len(groups)):
						cat_list = ["[Стены]", "[Полы]", "[Потолки]"]
						cat = cat_list[g]
						a_max = 0.00
						parameters = def_params_rooms[g]
						if len(groups[g]) != 0 and ((result_2 != "1" and g == 0) or (g != 0 and result_2 == "3")):
							par = room.LookupParameter(parameters[0])
							par.Set(0)
							description_walls = []
							room_description_uniq = []
							room_description_area = []
							room_description_type = []
							plinth_room_description = []
							plinth_room_description_uniq = []
							plinth_room_description_length = []
							plinth_room_description_type = []
							room_area = 0.00
							if len(groups[g]) != 0:
								for element in groups[g]:
									if element.LookupParameter(def_params_elements[2]).AsString() == str(room.Id):
										type = getType(element)
										el_area = 0.00
										el_length = 0.00
										el_height = 0.00
										if g == 0:
											if type.LookupParameter("О_Плинтус").AsInteger() != 1:
												el_description = normalized_string(str(type.LookupParameter(def_param_description).AsString()))
												if el_description == "<НЕЗАПОЛНЕНО!>":
													AddReport(type, "{} Отсутствует значение параметра «{}»".format(cat, def_param_description))
												el_area = round(element.get_Parameter(def_param_area).AsDouble() * 0.09290304, 2)
												try:
													el_type = str(type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString())
													if el_type.lower() != "" and el_type.lower() != "none":
														el_type = "Тип: " + el_type
														if not in_list(str(type.Id), types):
															types.append(str(type.Id))
															areas.append(el_area)
														else:
															append_areas(types, str(type.Id), areas, el_area)
													else:
														AddReport(type, "{} Отсутствует значение параметра «Маркировка типоразмера»".format(cat))
														el_type = ""
												except:
													AddReport(type, "{} Отсутствует параметр «Маркировка типоразмера»".format(cat))
													el_type = ""
												room_area += el_area
												a_max += el_area
												if not in_list("{} : {}".format(el_type, el_description), room_description_uniq):
													room_description_uniq.append("{} : {}".format(el_type, el_description))
													description_walls.append(el_description)
													room_description_area.append(el_area)
													room_description_type.append(el_type)
												else:
													index = get_index(el_description, description_walls)
													room_description_area[index] += el_area
											else:
												el_description = normalized_string(str(type.LookupParameter(def_param_description).AsString()))
												if el_description == "<НЕЗАПОЛНЕНО!>":
													AddReport(type, "{} Отсутствует значение параметра «{}»".format(cat, def_param_description))
												el_area = element.get_Parameter(def_param_area).AsDouble() * 0.09290304
												try:
													el_height = type.LookupParameter("О_Плинтус_Высота").AsDouble() * 0.3048
													el_length = round(el_area / el_height, 2)
													try:
														el_type = str(type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString())
														if el_type.lower() != "" and el_type.lower() != "none":
															el_type = "Тип: " + el_type
														else:
															AddReport(type, "{} Отсутствует значение параметра «Маркировка типоразмера»".format(cat))
															el_type = ""
													except:
														AddReport(type, "{} Отсутствует параметр «Маркировка типоразмера»".format(cat))
														el_type = ""
													if not in_list("{} : {}".format(el_type, el_description), plinth_room_description_uniq):
														plinth_room_description_uniq.append("{} : {}".format(el_type, el_description))
														plinth_room_description.append(el_description)
														plinth_room_description_length.append(el_length)
														plinth_room_description_type.append(el_type)
													else:
														index = get_index(el_description, plinth_room_description)
														plinth_room_description_length[index] += el_length
												except:
													AddReport(type, "{} Отсутствует параметр «О_Плинтус_Высота» либо значение не установлено!".format(cat))

										else:
											el_description = normalized_string(str(type.LookupParameter(def_param_description).AsString()))
											if el_description == "<НЕЗАПОЛНЕНО!>":
												AddReport(type, "{} Отсутствует значение параметра «{}»".format(cat, def_param_description))
											el_area = round(element.get_Parameter(def_param_area).AsDouble() * 0.09290304, 2)
											try:
												el_type = str(type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString())
												if el_type.lower() != "" and el_type.lower() != "none":
													el_type = "Тип: " + el_type
													if g == 2:
														if not in_list(str(type.Id), types):
															types.append(str(type.Id))
															areas.append(el_area)
														else:
															append_areas(types, str(type.Id), areas, el_area)
												else:
													AddReport(type, "{} Отсутствует значение параметра «Маркировка типоразмера»".format(cat))
													el_type = ""
											except:
												AddReport(type, "{} Отсутствует параметр «Маркировка типоразмера»".format(cat))
												el_type = ""
											room_area += el_area
											a_max += el_area
											if not in_list("{} : {}".format(el_type, el_description), room_description_uniq):
												room_description_uniq.append("{} : {}".format(el_type, el_description))
												description_walls.append(el_description)
												room_description_area.append(el_area)
												room_description_type.append(el_type)
											else:
												index = get_index(el_description, description_walls)
												room_description_area[index] += el_area
							
								if result == 'Yes':
									#OTHER ELEMENTS
									text = ""
									for i in range(0, len(description_walls)):
										text += "\n\n‎     {}. {}\n{}".format(str(i+1), room_description_type[i], description_walls[i])#СОДЕРЖИТ UNICODE
									par = room.LookupParameter(parameters[1])
									if text != "":
										par.Set(text)
									else:
										par.Set("‎     1.\nБез отделки")#СОДЕРЖИТ UNICODE
										
									text = ""
									n = ""
									for i in range(0, len(description_walls)):
										n = ""
										ncount = description_walls[i].count("\n")
										for z in range(0, ncount+2):
											n += "\n"
										if i == (len(description_walls)-1):
											n = "\n"
										text += "{}.\n{} м²{}".format(str(i+1), str(round(room_description_area[i] * x, 2)), n)
									if result_4:
										if len(description_walls) > 1:
											text += "\nИтого: {} м²".format(str(round(a_max * x, 2)))
									par = room.LookupParameter(parameters[2])
									if text != "":
										par.Set(text)
									else:
										par.Set("")
									#PLINTHS
									if g == 0:
										text = ""
										for i in range(0, len(plinth_room_description)):
											text += "\n\n‎     {}. {}\n{}".format(str(i+1), plinth_room_description_type[i], plinth_room_description[i])#СОДЕРЖИТ UNICODE
										par = room.LookupParameter("О_ПОМ_Описание плинтусов")
										if text != "":
											par.Set(text)
										else:
											par.Set("‎     1.\nБез отделки")#СОДЕРЖИТ UNICODE
										text = ""
										n = ""
										for i in range(0, len(plinth_room_description)):
											n = ""
											ncount = plinth_room_description[i].count("\n")
											for z in range(0, ncount+2):
												n += "\n"
											if i == (len(plinth_room_description)-1):
												n = "\n"
											text += "{}.\n{} пог.м.{}".format(str(i+1), str(round(plinth_room_description_length[i] * x, 2)), n)
										if result_4:
											if len(plinth_room_description) > 1:
												text += "\nИтого: {} пог.м.".format(str(round(a_max * x, 2)))
										par = room.LookupParameter("О_ПОМ_Длина плинтусов_Текст")
										if text != "":
											par.Set(text)
										else:
											par.Set("")
								par = room.LookupParameter(parameters[0])
								par.Set(room_area)
						else:		
							par = room.LookupParameter(parameters[0])
							description_w = room.LookupParameter(parameters[1]).AsString()
							if not description_w:
								description_w = ""
							if description_w.lower() != "без отделки" and description_w.lower() != "":
								if g == 0:
									area = room.get_Parameter(def_param_room_p).AsDouble() * room.get_Parameter(def_param_room_h).AsDouble()
									for e in elements_all:
										if e[0] or e[1] == str(room.Id):
											area -= e[2] * e[3]
									par.Set(area)
								else:
									area = room.get_Parameter(def_param_room_s).AsDouble()
									par.Set(area)
							else:
								par.Set(0)
					types_sorted = []
					if len(types) != 0:
						for i in types:
							types_sorted.append(i)
						types_sorted.sort()
						areas_sorted = sort_by_key(areas, types)
						append_dict(types_sorted, areas_sorted, room, type_groups, sort_by)
			commands_3 = [CommandLink('Имена', return_value='1'), CommandLink('Имена + (стандартные номера помещений)', return_value='2'), CommandLink('Имена + (задать параметр)', return_value='3')]
			dialog3 = TaskDialog('Выберите тип отображения помещений в ведомости по ГОСТ',
					title = "Имена помещений",
					title_prefix=False,
					content="",
					commands=commands_3,
					footer='Значения записываются в параметр помещений - «О_ПОМ_Группа»',
					show_close=False)
			result_3 = dialog3.show()
			custom_parameter = ""
			if result_3 == "3":
				form = PickParameter()
				Application.Run(form)
			if custom_parameter == "[USE DEFAULT]":
				result_3 = "2"
			with db.Transaction(name = "KPLN: Обновить отделку ГОСТ"):
				for dict in type_groups:
					a_max_w = 0
					a_max_c = 0
					names = ""
					numbers = ""
					group = ""
					custom_values = ""
					for room in dict[2]:
						name = str(room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString())
						if not(name in names):
							if names == "":
								names += "{}".format(name)
							else:
								names += ", {}".format(name)
						number = str(room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
						if not(number in numbers):
							if numbers == "":
								numbers += "{}".format(number)
							else:
								numbers += ", {}".format(number)
						if result_3 == "3":
							try:
								value = str(room.LookupParameter(custom_parameter).AsString())
							except:
								value = "<НЕЗАПОЛНЕНО!>"
								AddReport(room, "[Помещения] Отсутствует значение параметра «{}»".format(custom_parameter))
							if not(value in custom_values):
								if custom_values == "":
									custom_values += "{}".format(value)
								else:
									custom_values += ", {}".format(value)
					description_w = ""
					description_c = ""
					area_w = ""
					area_c = ""
					description_walls = []
					description_ceilings = []
					area_walls = []
					area_ceilings = []
					num_w = 1
					num_c = 1
					for str_id, a in zip(dict[0], dict[1]):
						id = DB.ElementId(int(str_id))
						type = doc.GetElement(id)
						if int(type.Category.Id.ToString()) == -2000011:
							des = str(type.LookupParameter("О_Описание").AsString())
							if des == "None":
								AddReport(type.Id, "Ошибка значения")
							description_w += "\n\n‎    {}. Тип: {}\n{}".format(str(num_w), type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString(), des)#СОДЕРЖИТ UNICODE
							description_walls.append(str(type.LookupParameter("О_Описание").AsString()))
							if type.LookupParameter("О_Описание").AsString() == None:
								AddReport(type, "[Стены] Отсутствует значение параметра «О_Описание»")
							area_walls.append(a)
							a_max_w += a
							num_w += 1
						elif int(type.Category.Id.ToString()) == -2000038:
							des = str(type.LookupParameter("О_Описание").AsString())
							if des == "None":
								AddReport(type.Id, "Ошибка значения")
							description_c += "\n\n‎    {}. Тип: {}\n{}".format(str(num_c), type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString(), des)#СОДЕРЖИТ UNICODE
							description_ceilings.append(str(type.LookupParameter("О_Описание").AsString()))
							if type.LookupParameter("О_Описание").AsString() == None:
								AddReport(type, "[Потолки] Отсутствует значение параметра «О_Описание»")
							area_ceilings.append(a)
							a_max_c += a
							num_c += 1
					if description_w == "":
						description_w = "‎     1.\nБез отделки"#СОДЕРЖИТ UNICODE
					if description_c == "":
						description_c = "‎     1.\nБез отделки"#СОДЕРЖИТ UNICODE
					num = 1
					for i in range(0, len(area_walls)):
						a = area_walls[i]
						des = description_walls[i]
						n = ""
						ncount = des.count("\n")
						for z in range(0, ncount + 2):
							n += "\n"
						if i == len(description_walls)-1:
							n = "\n"
						area_w += "{}.\n{} м²{}".format(str(num), str(round(a * x, 2)), n)
						num += 1
					num = 1
					for i in range(0, len(area_ceilings)):
						a = area_ceilings[i]
						des = description_ceilings[i]
						n = ""
						ncount = des.count("\n")
						for z in range(0, ncount + 2):
							n += "\n"
						if i == len(description_ceilings)-1:
							n = "\n"
						area_c += "{}.\n{} м²{}".format(str(num), str(round(a * x, 2)), n)
						num += 1
					if result_4:
						if len(description_walls) > 1:
							area_w += "\nИтого: {} м²".format(str(round(a_max_w * x, 2)))
						if len(description_ceilings) > 1:
							area_c += "\nИтого: {} м²".format(str(round(a_max_c * x, 2)))
					if result_3 == "1":
						group = "{}".format(names)
					elif result_3 == "2":
						group = "{}\n({})".format(names, numbers)
					elif result_3 == "3":
						group = "{}\n({})".format(names, custom_values)
					for room in dict[2]:
						room.LookupParameter("О_ПОМ_Группа").Set(group)
						room.LookupParameter("О_ПОМ_ГОСТ_Описание стен").Set(description_w)
						room.LookupParameter("О_ПОМ_ГОСТ_Площадь стен_Текст").Set(area_w)
						room.LookupParameter("О_ПОМ_ГОСТ_Описание потолков").Set(description_c)
						room.LookupParameter("О_ПОМ_ГОСТ_Площадь потолков_Текст").Set(area_c)