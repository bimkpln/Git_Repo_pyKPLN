# -*- coding: utf-8 -*-
"""
FW_Update

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Обновить (не группировать по типам)"
__doc__ = 'Запись расчетных значений отделки в помещения без разделения помещений (но с группировкой по имени) по типоразмерам отделки\n\n' \
          ' - Для формирования ведомостей по ГОСТ использовать параметры имеющие индекс «_ГОСТ_»\n' \
          ' - Для остальных спецификаций - все параметры начиная с «О_Отделка_»' \

"""
Архитекурное бюро KPLN

"""
import math
import itertools
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


class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value


def create_check_boxes_by_name(elements):
    # Create check boxes for elements if they have Name property
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name) if not e.GetLinkDocument() is None]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='Для подключения элементов из связи - выбери нужную:',
                                                    width=600,
                                                    button_name='Запуск!/Связь не нужна!')
    return elements_checkboxes


class PickParameterWindow(Form):
	def __init__(self):
		global setting_custom_parameter
		self.Name = "KPLN_Finishing"
		self.Text = "KPLN Отделка : Номера помещений"
		self.Size = Size(300, 150)
		self.MinimumSize = Size(300, 150)
		self.MaximumSize = Size(300, 150)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = False
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		self.TopMost = True
		#INIT
		self.button_apply = Button()
		self.button_pass = Button()
		self.label_title = Label()
		self.combobox_param = ComboBox()
		#
		self.combobox_param.Parent = self
		self.combobox_param.Location = Point(12, 54)
		self.combobox_param.Size = Size(260, 21)
		self.combobox_param.DropDownStyle = ComboBoxStyle.DropDownList
		self.combobox_param.SelectedIndexChanged += self.CbOnChanged
		#
		self.button_apply.Parent = self
		self.button_apply.Location = Point(12, 81)
		self.button_apply.Size = Size(75, 23)
		self.button_apply.Text = "Применить"
		self.button_apply.Click += self.OnOk
		self.button_apply.Enabled = False
		#
		self.button_pass.Parent = self
		self.button_pass.Location = Point(93, 81)
		self.button_pass.Size = Size(179, 23)
		self.button_pass.Text = "Использовать стандартный №"
		self.button_pass.Click += self.OnPass
		#
		self.label_title.Parent = self
		self.label_title.Location = Point(12, 13)
		self.label_title.Size = Size(260, 38)
		self.label_title.Text = "Определите параметр помещений для расчета:"
		#
		self.parameters = []
		for j in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement().Parameters:
			try:
				if j.IsShared and str(j.StorageType) == "String":
					self.parameters.append(j.Definition.Name)
			except: pass
		self.parameters.sort()
		for parameter in self.parameters:
			self.combobox_param.Items.Add(parameter)
		self.CenterToScreen()

	def CbOnChanged(self, sender, event):
		if self.combobox_param.Text != "":
			self.button_apply.Enabled = True

	def OnPass(self, sender, args):
		global setting_custom_parameter
		setting_custom_parameter = ""
		self.Close()

	def OnOk(self, sender, args):
		global setting_custom_parameter
		setting_custom_parameter = self.combobox_param.Text
		self.Close()

class PreferencesWindow(Form):
	def __init__(self):
		global setting_group_by
		global setting_calculate_result
		global setting_k
		global setting_break
		global setting_add_numbers
		self.Name = "KPLN_Finishing"
		self.Text = "KPLN Отделка : Настройки"
		self.Size = Size(300, 450)
		self.MinimumSize = Size(300, 450)
		self.MaximumSize = Size(300, 450)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = False
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		self.TopMost = True
		#INIT
		self.button_apply = Button()
		self.button_pass = Button()
		self.label_title = Label()
		self.label_preferences = Label()
		self.label_preferences_types = Label()
		self.label_k = Label()
		self.combobox_param01 = ComboBox()
		self.combobox_param02 = ComboBox()
		self.combobox_param03 = ComboBox()
		self.checkbox_calculate_results = CheckBox()
		self.checkbox_add_numbers = CheckBox()
		self.checkbox_w = CheckBox()
		self.checkbox_f = CheckBox()
		self.checkbox_c = CheckBox()
		self.checkbox_p = CheckBox()
		self.combobox_k = ComboBox()
		self.combobox_param01.Items.Add("<неактивен>")
		self.combobox_param02.Items.Add("<неактивен>")
		self.combobox_param03.Items.Add("<неактивен>")
		self.combobox_param01.Text = "<неактивен>"
		self.combobox_param02.Text = "<неактивен>"
		self.combobox_param03.Text = "<неактивен>"
		#
		self.combobox_param01.Parent = self
		self.combobox_param01.Location = Point(12, 54)
		self.combobox_param01.Size = Size(260, 21)
		self.combobox_param01.DropDownStyle = ComboBoxStyle.DropDownList
		self.combobox_param01.SelectedIndexChanged += self.CheckBoxSelectedIndexChanged
		#
		self.combobox_param02.Parent = self
		self.combobox_param02.Location = Point(12, 79)
		self.combobox_param02.Size = Size(260, 21)
		self.combobox_param02.DropDownStyle = ComboBoxStyle.DropDownList
		self.combobox_param02.SelectedIndexChanged += self.CheckBoxSelectedIndexChanged
		self.combobox_param02.Enabled = False
		#
		self.combobox_param03.Parent = self
		self.combobox_param03.Location = Point(12, 104)
		self.combobox_param03.Size = Size(260, 21)
		self.combobox_param03.DropDownStyle = ComboBoxStyle.DropDownList
		self.combobox_param03.SelectedIndexChanged += self.CheckBoxSelectedIndexChanged
		self.combobox_param03.Enabled = False
		#
		self.button_apply.Parent = self
		self.button_apply.Location = Point(12, 382)
		self.button_apply.Size = Size(75, 23)
		self.button_apply.Text = "Применить"
		self.button_apply.Click += self.ButtonApply
		#
		self.button_pass.Parent = self
		self.button_pass.Location = Point(93, 382)
		self.button_pass.Size = Size(75, 23)
		self.button_pass.Text = "Отмена"
		self.button_pass.Click += self.ButtonCancel
		#
		self.label_title.Parent = self
		self.label_title.Location = Point(12, 13)
		self.label_title.Size = Size(260, 38)
		self.label_title.Text = "Группировка отделки помещений (только для ведомостей по ГОСТ):"
		#
		self.label_preferences.Parent = self
		self.label_preferences.Location = Point(12, 235)
		self.label_preferences.Size = Size(260, 18)
		self.label_preferences.Text = "Настройки:"
		#
		self.label_preferences_types.Parent = self
		self.label_preferences_types.Location = Point(12, 135)
		self.label_preferences_types.Size = Size(260, 18)
		self.label_preferences_types.Text = "Учитываемые при группировки типы отделки:"
		#
		self.checkbox_calculate_results.Parent = self
		self.checkbox_calculate_results.Location = Point(12, 255)
		self.checkbox_calculate_results.Text = "Рассчитывать итоговую площадь"
		self.checkbox_calculate_results.AutoSize = True
		self.checkbox_calculate_results.Checked = setting_calculate_result
		self.checkbox_calculate_results.CheckedChanged += self.CalculateResultChanged
		#
		self.checkbox_add_numbers.Parent = self
		self.checkbox_add_numbers.Location = Point(12, 275)
		self.checkbox_add_numbers.Text = "Имена с номерами помещений"
		self.checkbox_add_numbers.AutoSize = True
		self.checkbox_add_numbers.Checked = setting_add_numbers
		self.checkbox_add_numbers.CheckedChanged += self.AddNumbersChanged
		#
		self.checkbox_w.Parent = self
		self.checkbox_w.Location = Point(12, 160)
		self.checkbox_w.Text = "Стены"
		self.checkbox_w.AutoSize = True
		self.checkbox_w.Enabled = False
		self.checkbox_w.Checked = True
		self.checkbox_w.CheckedChanged += self.AddNumbersChanged
		#
		self.checkbox_f.Parent = self
		self.checkbox_f.Location = Point(12, 180)
		self.checkbox_f.Text = "Полы"
		self.checkbox_f.AutoSize = True
		self.checkbox_f.Checked = True
		self.checkbox_f.CheckedChanged += self.AddNumbersChanged
		#
		self.checkbox_c.Parent = self
		self.checkbox_c.Location = Point(132, 160)
		self.checkbox_c.Text = "Потолки"
		self.checkbox_c.AutoSize = True
		self.checkbox_c.Checked = True
		self.checkbox_c.CheckedChanged += self.AddNumbersChanged
		#
		self.checkbox_p.Parent = self
		self.checkbox_p.Location = Point(132, 180)
		self.checkbox_p.Text = "Плинтусы"
		self.checkbox_p.AutoSize = True
		self.checkbox_p.Checked = True
		self.checkbox_p.CheckedChanged += self.AddNumbersChanged
		#
		self.label_k.Parent = self
		self.label_k.Location = Point(12, 320)
		self.label_k.Size = Size(260, 18)
		self.label_k.Text = "Коэффициент расчета площади (длины):"
		#
		self.combobox_k.Parent = self
		self.combobox_k.Location = Point(12, 344)
		self.combobox_k.Size = Size(260, 21)
		self.combobox_k.DropDownStyle = ComboBoxStyle.DropDownList
		self.combobox_k.SelectedIndexChanged += self.KChanged
		self.combobox_k.Enabled = True
		self.combobox_k.Items.Add("1.0")
		self.combobox_k.Items.Add("1.05")
		self.combobox_k.Items.Add("1.1")
		self.combobox_k.Items.Add("1.15")
		self.combobox_k.Items.Add("1.25")
		self.combobox_k.Items.Add("1.5")
		self.combobox_k.Items.Add("2.0")
		self.combobox_k.Items.Add("0.9")
		self.combobox_k.Items.Add("0.85")
		self.combobox_k.Text = "1.0"
		#
		self.parameters = []
		for j in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement().Parameters:
			try:
				if str(j.StorageType) == "String":
					self.parameters.append(j.Definition.Name)
			except: pass
		self.parameters.sort()
		for parameter in self.parameters:
			self.combobox_param01.Items.Add(parameter)
			self.combobox_param02.Items.Add(parameter)
			self.combobox_param03.Items.Add(parameter)
		self.CenterToScreen()

	def CalculateResultChanged(self, sender, args):
		global setting_calculate_result
		setting_calculate_result = sender.Checked

	def KChanged(self, sender, event):
		global setting_k
		setting_k = float(sender.Text)

	def AddNumbersChanged(self, sender, args):
		global setting_add_numbers
		setting_add_numbers = sender.Checked

	def CheckBoxSelectedIndexChanged(self, sender, event):
		if self.combobox_param01.Text != "<неактивен>":
			self.combobox_param02.Enabled = True
		else:
			self.combobox_param02.Enabled = False
			self.combobox_param02.Text = "<неактивен>"
		if self.combobox_param02.Text != "<неактивен>":
			self.combobox_param03.Enabled = True
		else:
			self.combobox_param03.Enabled = False
			self.combobox_param03.Text = "<неактивен>"

	def ButtonCancel(self, sender, args):
		global setting_break
		setting_break = True
		self.Close()

	def ButtonApply(self, sender, args):
		global setting_types_f
		global setting_types_c
		global setting_types_p
		global setting_group_by
		if self.combobox_param01.Text != "<неактивен>":
			setting_group_by.append(self.combobox_param01.Text)
		if self.combobox_param02.Text != "<неактивен>":
			setting_group_by.append(self.combobox_param02.Text)
		if self.combobox_param03.Text != "<неактивен>":
			setting_group_by.append(self.combobox_param03.Text)
		setting_types_f = self.checkbox_f.Checked
		setting_types_c = self.checkbox_c.Checked
		setting_types_p = self.checkbox_p.Checked
		self.Close()

class ElementsStack():
	def __init__(self, plinth=False):
		self.IsPlinth = plinth
		self.CommonArea = 0.0
		self.InfoValue = []
		self.InfoDescription = []
		self.Area = []
		self.Types_ids = []
		self.Elements = []
		self.Count = 0

	def GetTypesIds(self):
		types_ids_sorted = []
		for int_id in self.Types_ids:
			types_ids_sorted.append(str(int_id))
		types_ids_sorted.sort()
		return "".join(types_ids_sorted)

	def GetElements(self):
		elements = []
		for set in self.Elements:
			for element in set:
				elements.append(element)
		return elements

	def CalculateArea(self):
		global GetType
		global setting_k
		global setting_calculate_result
		self.InfoValue = []
		self.CommonArea = 0.0
		for set in self.Elements:
			area = 0.0
			n = 0
			for element in set:
				if not self.IsPlinth:
					area += round(element.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble() * 0.09290304 * setting_k, 2)
				else:
					area += round((element.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble() * 0.09290304) / (GetType(element).LookupParameter("О_Плинтус_Высота").AsDouble() * 304.8 / 1000) * setting_k, 2)
				try:
					n = GetType(element).LookupParameter("О_Описание").AsString().count("\n")
				except:
					pass
			space = ""
			for i in range(0, n):
				space += "\n"
			#
			prefix = ""
			mark = GetType(element).get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString()
			if mark == None:
				mark = '?'
				prefix = 'ОШИБКА!\n    - МАРКА\n\n\n\n'.format(prefix)
			#
			if not self.IsPlinth:
				self.InfoValue.append("{}Тип: «{}»\n{}м²{}".format(prefix, mark, str(round(area, 2)), space))
			else:
				self.InfoValue.append("{}Тип: «{}»\n{} пог.м{}".format(prefix, mark, str(round(area, 2)), space))
			self.CommonArea += area
		self.InfoValue.sort()
		self.InfoDescription.sort()
		if setting_calculate_result and len(self.Types_ids) > 1:
			if not self.IsPlinth:
				self.InfoValue.append("Итого: {}м²".format(str(round(self.CommonArea, 2))))
			else:
				self.InfoValue.append("Итого: {} пог.м".format(str(round(self.CommonArea, 2))))

	def GetValueDescription(self):
		value = "\n\n".join(self.InfoValue)
		return value

	def GetDescription(self):
		value = "\n\n".join(self.InfoDescription)
		return value

	def GetCommonArea(self):
		return self.CommonArea

	def GetCommonLength(self):
		return self.CommonArea

	def Append(self, element):
		global GetType
		if not self.InList(GetType(element).Id.IntegerValue):
			self.Types_ids.append(GetType(element).Id.IntegerValue)
			self.Elements.append([element])
			prefix = ""
			error = False
			des = GetType(element).LookupParameter("О_Описание").AsString()
			if des == None:
				prefix += "    - ОТСУТСТВУЕТ ОПИСАНИЕ\n"
				des = "<отсутствует описание>"
				error = True
			mark = GetType(element).get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString()
			if mark == None:
				prefix += "    - ОТСУТСТВУЕТ МАРКИРОВКА ТИПОРАЗМЕРА\n"
				mark = '?'
				error = True
			if error:
				prefix = 'ОШИБКА ЗАПОЛНЕНИЯ!\n{}\n\n\n'.format(prefix)
			self.InfoDescription.append("{}‎Тип: «{}»\n{}".format(prefix, mark, des))
			self.Area.append(element.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble())
			self.CommonArea += element.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble()
			self.Count += 1
		else:
			for i in range(0, len(self.Types_ids)):
				if self.Types_ids[i] == GetType(element).Id.IntegerValue:
					self.Elements[i].append(element)
					self.Area[i] += element.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble()
					self.CommonArea += element.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble()
					self.Count += 1
		self.CalculateArea()

	def InList(self, id):
		for i in self.Types_ids:
			if i == id:
				return True
		return False

class RoomFinish():
	def __init__(self, room=DB.Element, walls=[], floors=[], ceilings=[], plinths=[]):
		self.Id = room.Id.IntegerValue
		self.Room = room
		self.Rooms = []
		self.Rooms.append(room)
		self.Walls = ElementsStack()
		self.Floors = ElementsStack()
		self.Ceilings = ElementsStack()
		self.Plinths = ElementsStack(True)
		self.Key = ""
		self.Count = 0

	@property
	def Id(self):
		return self.Id

	@property
	def Room(self):
		return self.Room

	def AppendRoom(self, room):
		self.Rooms.append(room)

	def Unique(self, list): 
		unique_list = [] 
		for x in list: 
			if x not in unique_list: 
				unique_list.append(x) 
		return unique_list

	def IsEmpty(self):
		if self.Count==0:
			return True
		else:
			return False

	def GetKey(self, input=[]):
		global setting_types_f
		global setting_types_c
		global setting_types_p
		if len(input) == 0:
			#key_floors = ""
			#key_ceilings = ""
			#key_plinths = ""
			#Activate to enable uniq types
			#if setting_types_f: key_floors = self.Floors.GetTypesIds()
			#if setting_types_c: key_ceilings = self.Ceilings.GetTypesIds()
			#if setting_types_p: key_plinths = self.Plinths.GetTypesIds()
			#self.Key = "{}{}{}{}".format(self.Walls.GetTypesIds(), key_floors, key_ceilings, key_plinths)
			self.Key = self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
		else:
			add_key = ""
			key_floors = ""
			key_ceilings = ""
			key_plinths = ""
			if setting_types_f: key_floors = self.Floors.GetTypesIds()
			if setting_types_c: key_ceilings = self.Ceilings.GetTypesIds()
			if setting_types_p: key_plinths = self.Plinths.GetTypesIds()
			for param in input:
				try:
					add_key += str(self.Room.LookupParameter(param).AsString())
				except Exception as e: print(str(e))
			#self.Key = add_key + "{}{}{}{}".format(self.Walls.GetTypesIds(), key_floors, key_ceilings, key_plinths)
			self.Key = add_key + self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
		return self.Key

	def GetElements(self):
		list_of_elements = self.Walls.GetElements() + self.Floors.GetElements() + self.Ceilings.GetElements() + self.Plinths.GetElements()
		return list_of_elements

	def ApplyNull(self):
		global SetLookupParameter
		try:
			SetLookupParameter(self.Room, "О_ПОМ_Описание стен", "")
			SetLookupParameter(self.Room, "О_ПОМ_Описание полов", "")
			SetLookupParameter(self.Room, "О_ПОМ_Описание потолков", "")
			SetLookupParameter(self.Room, "О_ПОМ_Описание плинтусов", "")
			SetLookupParameter(self.Room, "О_ПОМ_Площадь стен_Текст", "")
			SetLookupParameter(self.Room, "О_ПОМ_Площадь полов_Текст", "")
			SetLookupParameter(self.Room, "О_ПОМ_Площадь потолков_Текст", "")
			SetLookupParameter(self.Room, "О_ПОМ_Длина плинтусов_Текст", "")
			SetLookupParameter(self.Room, "О_ПОМ_Площадь стен", 0.0)
			SetLookupParameter(self.Room, "О_ПОМ_Площадь полов", 0.0)
			SetLookupParameter(self.Room, "О_ПОМ_Площадь потолков", 0.0)
		except Exception as e: print((str(e)))

	def ApplyNullGrouped(self):
		global SetLookupParameter
		for room in self.Rooms:
			try:
				SetLookupParameter(room, "О_ПОМ_Группа", "")
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Описание стен", "")
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Описание полов", "")
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Описание потолков", "")
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Описание плинтусов", "")
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Площадь стен_Текст", "")
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Площадь полов_Текст", "")
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Площадь потолков_Текст", "")
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Длина плинтусов_Текст", "")
			except Exception as e: print((str(e)))

	def ApplySingle(self):
		global SetLookupParameter
		try:
			SetLookupParameter(self.Room, "О_ПОМ_Описание стен", self.Walls.GetDescription())
			SetLookupParameter(self.Room, "О_ПОМ_Описание полов", self.Floors.GetDescription())
			SetLookupParameter(self.Room, "О_ПОМ_Описание потолков",self.Ceilings.GetDescription())
			SetLookupParameter(self.Room, "О_ПОМ_Описание плинтусов", self.Plinths.GetDescription())
			SetLookupParameter(self.Room, "О_ПОМ_Площадь стен_Текст", self.Walls.GetValueDescription())
			SetLookupParameter(self.Room, "О_ПОМ_Площадь полов_Текст", self.Floors.GetValueDescription())
			SetLookupParameter(self.Room, "О_ПОМ_Площадь потолков_Текст", self.Ceilings.GetValueDescription())
			SetLookupParameter(self.Room, "О_ПОМ_Длина плинтусов_Текст", self.Plinths.GetValueDescription())
			SetLookupParameter(self.Room, "О_ПОМ_Площадь стен", self.Walls.GetCommonArea())
			SetLookupParameter(self.Room, "О_ПОМ_Площадь полов", self.Floors.GetCommonArea())
			SetLookupParameter(self.Room, "О_ПОМ_Площадь потолков", self.Ceilings.GetCommonArea())
		except Exception as e: print((str(e)))

	def ApplyGrouped(self):
		global setting_break_second
		global setting_add_numbers
		global SetLookupParameter
		global setting_group_by
		try:
			g_names = []
			g_numbers = []
			for room in self.Rooms:
				g_names.append(room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString())
				if setting_custom_parameter == "":
					g_numbers.append(room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
				else:
					g_numbers.append(room.LookupParameter(setting_custom_parameter).AsString())
			g_names_sorted = self.Unique(g_names)
			g_numbers_sorted = self.Unique(g_numbers)
			if setting_add_numbers:
				group = "{}\n\n({})".format(", ".join(g_numbers_sorted), ", ".join(g_names_sorted))
			else:
				group = "{}".format(", ".join(g_names_sorted))
			for room in self.Rooms:
				SetLookupParameter(room, "О_ПОМ_Группа", group)
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Описание стен", self.Walls.GetDescription())
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Описание полов", self.Floors.GetDescription())
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Описание потолков", self.Ceilings.GetDescription())
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Описание плинтусов", self.Plinths.GetDescription())
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Площадь стен_Текст", self.Walls.GetValueDescription())
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Площадь полов_Текст", self.Floors.GetValueDescription())
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Площадь потолков_Текст", self.Ceilings.GetValueDescription())
				SetLookupParameter(room, "О_ПОМ_ГОСТ_Длина плинтусов_Текст", self.Plinths.GetValueDescription())
		except Exception as e: print(str(e))

	def AppendElement(self, element):
		global GetType
		if element.Category.Id.IntegerValue == -2000011:
			if GetType(element).LookupParameter("О_Плинтус").AsInteger() == 1:
				try:
					if GetType(element).LookupParameter("О_Плинтус_Высота").AsDouble() > 0.00:
						self.Plinths.Append(element)
						self.Count += 1
				except Exception as e: print(str(e))
			else:
				self.Walls.Append(element)
				self.Count += 1
		elif element.Category.Id.IntegerValue == -2000032:
			self.Floors.Append(element)
			self.Count += 1
		elif element.Category.Id.IntegerValue == -2000038:
			self.Ceilings.Append(element)
			self.Count += 1

def SetLookupParameter(element, parameter, value):
	try:
		element.LookupParameter(parameter).Set(value)
	except:
		if not '«{}» - Отсутствует параметр «{}». Решение: Воспользуйтесь загрузчиком параметров;'.format(element.Category.Name, parameter) in errors:
			print('«{}» - Отсутствует параметр «{}». Решение: Воспользуйтесь загрузчиком параметров;'.format(element.Category.Name, parameter))
			errors.append('«{}» - Отсутствует параметр «{}». Решение: Воспользуйтесь загрузчиком параметров;'.format(element.Category.Name, parameter))

def GetStack(id):
	for room_stack in rooms_all:
		if room_stack.Id == id:
			return room_stack

def GetStackGrouped(key):
	for stack_grouped in rooms_grouped:
		if key == stack_grouped.GetKey(setting_group_by):
			return stack_grouped
	return False

def GetType(element):
	try:
		id = element.GetTypeId()
		type = doc.GetElement(id)
		return type
	except:
		try:
			type = element.WallType
			return type
		except:
			try:
				type = element.FloorType
				return type
			except:
				return False

setting_group_by = []
setting_calculate_result = True
setting_add_numbers = False
setting_k = 1.0
setting_break = False
setting_custom_parameter = ""
setting_types_f = False
setting_types_c = False
setting_types_p = False
collector_rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
collectro_walls = list()
collector_floors = list()
collector_ceilings = list()
collectro_walls.extend(DB.FilteredElementCollector(doc).
                       OfCategory(DB.BuiltInCategory.OST_Walls).
                       WhereElementIsNotElementType().ToElements())
collector_floors.extend(DB.FilteredElementCollector(doc).
                        OfCategory(DB.BuiltInCategory.OST_Floors).
                        WhereElementIsNotElementType().ToElements())
collector_ceilings.extend(DB.FilteredElementCollector(doc).
                          OfCategory(DB.BuiltInCategory.OST_Ceilings).
                          WhereElementIsNotElementType().ToElements())
errors = []

linkModelInstances = DB.FilteredElementCollector(doc).\
                     OfClass(DB.RevitLinkInstance)
linkModels_checkboxes = create_check_boxes_by_name(linkModelInstances)
linkModels = [c.value for c in linkModels_checkboxes if c.state is True]
for link in linkModels:
    linkDoc = link.GetLinkDocument()
    collectro_walls.extend(DB.FilteredElementCollector(linkDoc).
                           OfCategory(DB.BuiltInCategory.OST_Walls).
                           WhereElementIsNotElementType().
                           ToElements())
    collector_floors.extend(DB.FilteredElementCollector(linkDoc).
                        OfCategory(DB.BuiltInCategory.OST_Floors).
                        WhereElementIsNotElementType().ToElements())
    collector_ceilings.extend(DB.FilteredElementCollector(linkDoc).
                            OfCategory(DB.BuiltInCategory.OST_Ceilings).
                            WhereElementIsNotElementType().ToElements())
errors = []

form = PreferencesWindow()
Application.Run(form)
ticker_current = 0
ticker_current_percent = 0
ticker_max = collector_rooms.Count * 2 + collectro_walls.Count + collector_floors.Count + collector_ceilings.Count
if not setting_break:
	with forms.ProgressBar(title='Создание коллекций : {value}%',indeterminate=False ,cancellable=False, step=1) as pb:
		if setting_add_numbers:
			form_pr = PickParameterWindow()
			Application.Run(form_pr)
		rooms_all = []
		for room in collector_rooms:
			ticker_current += 1
			ticker_current_percent = int(ticker_current * 100 / ticker_max)
			pb.update_progress(ticker_current_percent, max_value = 100)
			rooms_all.append(RoomFinish(room))
		for collector in [collectro_walls, collector_floors, collector_ceilings]:
			pb.title = 'Расчет значений : {value}%'
			for element in collector:
				ticker_current += 1
				ticker_current_percent = int(ticker_current * 100 / ticker_max)
				pb.update_progress(ticker_current_percent, max_value = 100)
				try:
					if element.LookupParameter("О_Id помещения").AsString() != "" and element.LookupParameter("О_Id помещения").AsString() != None:
						room_stack = GetStack(int(element.LookupParameter("О_Id помещения").AsString()))
						room_stack.AppendElement(element)
				except Exception as e: print(str(e))
		rooms_grouped = []
		with db.Transaction(name = "KPLN: Обновить отделку"):
			pb.title = 'Группировка элементов : {value}%'
			for room_finish in rooms_all:
				ticker_current += 1
				ticker_current_percent = int(ticker_current * 100 / ticker_max)
				pb.update_progress(ticker_current_percent, max_value = 100)
				if not room_finish.IsEmpty():
					room_finish.ApplySingle()
					stack_exist = GetStackGrouped(room_finish.GetKey(setting_group_by))
					if stack_exist == False:
						rooms_grouped.append(room_finish)
						ticker_max += 1
					else:
						stack_exist.AppendRoom(room_finish.Room)
						for element in room_finish.GetElements():
							stack_exist.AppendElement(element)
				else:
					room_finish.ApplyNull()
					room_finish.ApplyNullGrouped()
			for stack in rooms_grouped:
				pb.title = 'Запись значений (группа) : {value}%'
				ticker_current += 1
				ticker_current_percent = int(ticker_current * 100 / ticker_max)
				pb.update_progress(ticker_current_percent, max_value = 100)
				stack.ApplyGrouped()