# -*- coding: utf-8 -*-
"""
ROOMS_SharedParameters

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Параметры ADSK"
__doc__ = 'Заполнение в помещениях параметров «ADSK_Тип квартиры», «ADSK_Тип помещения» «ADSK_Номер квартиры», «ADSK_Номер помещения квартиры»' \
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
from System.Windows.Forms import *
from System.Drawing import *
import System
import datetime
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

class Commerce_Set(Form):
	def __init__(self):
		global department_dict
		#INIT
		self.Name = "DIV02"
		self.Text = "Коммерческие помещения"
		self.Size = Size(370, 152)
		self.MinimumSize = Size(370, 152)
		self.MaximumSize = Size(370, 152)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = True
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		self.TopMost = True
		self.AutoScroll = False

		#VARIABLES
		self.gb = []
		self.cb = []

		self.gb = GroupBox()
		self.gb.Parent = self
		self.gb.Text = "Определите тип для коммерческих помещений"
		self.gb.Size = Size(270, 70)
		self.gb.Location = Point(12, 12)

		self.cb = ComboBox()
		self.cb.Parent = self.gb
		self.cb.Location = Point(18, 26)
		self.cb.Size = Size(230, 21)
		self.cb.DropDownStyle = ComboBoxStyle.DropDownList
		for item in department_dict:
			self.cb.Items.Add(item)
		self.cb.Text = department_dict[len(department_dict)-1]

		self.btn1 = Button(Text="Применить")
		self.btn1.Parent = self
		self.btn1.Location = Point(12, 12 + 76)
		self.btn1.Click += self.Apply

		self.btn2 = Button(Text="Пропустить")
		self.btn2.Parent = self
		self.btn2.Location = Point(100, 12 + 76)
		self.btn2.Click += self.Cancel

	def Cancel(self, sender, args):
		self.Close()

	def Apply(self, sender, args):
		rooms = []
		rooms_numbers = []
		with db.Transaction(name = "SetValue"):
			for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
				for i in range(0, len(self.Dict)):
					if room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString() == self.cb.Text:
						number = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()
						rooms.append(room)
						if not self.in_list(number, rooms_numbers):
							rooms_numbers.append(number)
			rooms_numbers.sort()
			for room in rooms:
				for i in range(0, len(rooms_numbers)-1):
					if room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString() == rooms_numbers[i]:
						room.LookupParameter("ADSK_Номер помещения квартиры").Set(str(i+1))
		self.Close()

	def in_list(self, element, list):
		for i in list:
			if element == i:
				return True
		return False

class Manual_Set(Form):
	def __init__(self):
		global errors_number
		#INIT
		self.Name = "DIV01"
		self.Text = "Неопределенные элементы"
		self.Size = Size(370, 200)
		self.MinimumSize = Size(370, 200)
		self.MaximumSize = Size(370, 200)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = True
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		self.TopMost = True
		self.Dict = []
		self.SetDict = ["Жилая комната",
			"Спальня",
			"Лоджия",
			"Балкон",
			"Терраса",
			"Веранда",
			"Кухня",
			"Гостиная",
			"Столовая",
			"Кухня-гостиная",
			"Кладовая в квартире",
			"Гардеробная",
			"Постирочная",
			"Коридор",
			"Холл",
			"Прихожая",
			"С/У",
			"С/У гостевой",
			"С/У дл МГН",
			"С/У для посетителей",
			"Туалет",
			"Ванная",
			"Кабинет",
			"<Пропустить>"]

		self.AutoScroll = False

		#VARIABLES
		self.gb = []
		self.cb = []

		for room in errors_number:
			self.Text = "Неопределенные элементы : {}".format(room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString())
			name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
			if not self.in_list(name, self.Dict):
				self.Dict.append(name)
		if len(self.Dict) > 7:
			self.AutoScroll = True
			self.MaximumSize = Size(370, 76*7 + 76)
			self.MinimumSize = Size(370, 76*7 + 76)
			self.Size = Size(370, 76*7 + 76)
		else:
			self.MaximumSize = Size(370, 76*len(self.Dict) + 76)
			self.MinimumSize = Size(370, 200)
			self.Size = Size(370, 76*len(self.Dict) + 76)

		for i in range(0, len(self.Dict)):
			self.gb.append(GroupBox())
			self.gb[i].Parent = self
			self.gb[i].Text = self.Dict[i]
			self.gb[i].Size = Size(270, 70)
			self.gb[i].Location = Point(12, 12 + 76*i)

			self.cb.append(ComboBox())
			self.cb[i].Parent = self.gb[i]
			self.cb[i].Location = Point(18, 26)
			self.cb[i].Size = Size(230, 21)
			self.cb[i].DropDownStyle = ComboBoxStyle.DropDownList
			for item in self.SetDict:
				self.cb[i].Items.Add(item)
			self.cb[i].Text = self.SetDict[len(self.SetDict)-1]

			if i == len(self.Dict) - 1:
				self.btn = Button(Text="Применить")
				self.btn.Parent = self
				self.btn.Location = Point(12, 12 + 76*i + 76)
				self.btn.Click += self.Apply

	def Apply(self, sender, args):
		with db.Transaction(name = "SetValue"):
			for room in errors_number:
				for i in range(0, len(self.Dict)):
					if room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString() == self.gb[i].Text:
						if self.cb[i].Text != "<Пропустить>":
							room.LookupParameter("ADSK_Номер помещения квартиры").Set(str(i+1))
						else:
							room.LookupParameter("ADSK_Номер помещения квартиры").Set("")
		self.Close()

	def in_list(self, element, list):
		for i in list:
			if element == i:
				return True
		return False

def in_list(element, list):
	for i in list:
		if element == i:
			return True
	return False

def set_number(dict, name):
	for i in range(0, len(dict)):
		if dict[i].lower() == name.lower():
			room.LookupParameter("ADSK_Номер помещения квартиры").Set(str(i+1))
			return True
	return False

group = "02 Обязательные АРХИТЕКТУРА"
params_elements = ["ADSK_Тип квартиры",
			"ADSK_Тип помещения",
			"ADSK_Номер квартиры",
			"ADSK_Номер помещения квартиры"]
params_elements_type = ["Text",
			"Integer",
			"Text",
			"Text"]
params_elements_vary = [True,
			False,
			True,
			True]

try:
	params_found = []
	common_parameters_file = "Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\ДИВ.pulldown\\Параметры ADSK.pushbutton\\ADSK_ФОП2019.txt"
	app = doc.Application
	category_set_rooms = app.Create.NewCategorySet()
	category_set_rooms.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Rooms))
	app.SharedParametersFilename = common_parameters_file
	SharedParametersFile = app.OpenSharedParameterFile()
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
						if d_Definition.VariesAcrossGroups == params_elements_vary[i]:
							if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Rooms)):
								params_found.append(d_Name)
	with db.Transaction(name = "AddSharedParameter"):
		for dg in SharedParametersFile.Groups:
			if dg.Name == group:
				for item in params_elements:
					if not in_list(item, params_found):
						externalDefinition = dg.Definitions.get_Item(item)
						newIB = app.Create.NewInstanceBinding(category_set_rooms)
						doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_IDENTITY_DATA)
						doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_IDENTITY_DATA)
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
						d_Definition.SetAllowVaryBetweenGroups(doc, params_elements_vary[i])
except: forms.alert("Не удалось подгрузить все требуемые параметры в проект!")
rooms_appartments = []
rooms = []
errors_number = []
department_dict = []
with db.Transaction(name = "SetValues"):
	for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
		room.LookupParameter("ADSK_Тип квартиры").Set("")
		room.LookupParameter("ADSK_Номер квартиры").Set("")
		room.LookupParameter("ADSK_Номер помещения квартиры").Set("")
		try:
			b = room.LookupParameter("КВ_Код").AsString()
			room.LookupParameter("ADSK_Тип квартиры").Set(b)
		except: pass
		name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
		department = room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString()
		if not in_list(department, department_dict):
			department_dict.append(department)
		try:
			v = room.LookupParameter("КВ_Номер").AsString()
			if not v:
				v = ""
		except: v = ""
		room.LookupParameter("ADSK_Номер квартиры").Set(v)
		dict = ["Жилая комната",
			"Спальня",
			"Лоджия",
			"Балкон",
			"Терраса",
			"Веранда",
			"Кухня",
			"Гостиная",
			"Столовая",
			"Кухня-гостиная",
			"Кладовая в квартире",
			"Гардеробная",
			"Постирочная",
			"Коридор",
			"Холл",
			"Прихожая",
			"С/У",
			"С/У гостевой",
			"С/У дл МГН",
			"С/У для посетителей",
			"Туалет",
			"Ванная",
			"Кабинет"]
		if department == "Квартира" or department == "Кв":
			if not set_number(dict, name):
				errors_number.append(room)

if os.path.exists("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\ДИВ.pulldown\\Параметры ADSK.pushbutton\\ADSK_ФОП2019.txt"):
	commands = [CommandLink('Да', return_value=True), CommandLink('Отмена', return_value=False)]
	dialog = TaskDialog('Добавить ADSK параметры в проект и назназначить им значения?',
						title = "Дивное : Инструменты",
						title_prefix=False,
						content="ФОП - Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ADSK_ФОП2019.txt",
						commands=commands,
						footer='Предназначено только для проекта «Дивное»',
						show_close=False)
	result = dialog.show()
	if result:
		if len(errors_number) != 0:
			form_1 = Manual_Set()
			Application.Run(form_1)

		form_2 = Commerce_Set()
		Application.Run(form_2)
else:
	forms.alert("Не удалось найти файл общих параметров!\nКаталог: «Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\ДИВ.pulldown\\Параметры ADSK.pushbutton\\ADSK_ФОП2019.txt»")