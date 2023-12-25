# -*- coding: utf-8 -*-
"""
Roomdata

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Квартирография [Архив]"
__doc__ = 'UNIS: Только для ГЛУ К1'

"""
KPLN

"""

import math
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
import re
from itertools import chain
import datetime
from rpw.ui.forms import TextInput, Alert, select_folder

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "Квартирография"
		self.Text = "[UNI'S]   Квартирография"
		self.rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
		self.Size = Size(800, 445)
		self.temp_number = 40
		self.lb = []
		self.ControlBox = False
		self.values_1 = []
		self.values_2 = []
		self.values_3 = []
		self.TopMost = True
		self.values_4 = []
		self.check = []
		self.check_values = ""
		self.check_result = True
		self.lbl_names = ["Корпус", "Секция/Часть", "Этаж", "Номер квартиры (отн.)", "Внутреквартирный № пом."]
		self.ttip = ["Параметр тип: Текст", "Параметр тип: Текст", "Параметр тип: Текст", "Параметр тип: Текст", "Параметр тип: Текст"]
		self.def_values = ["ПОМ_Корпус", "ПОМ_Секция", "пом_этаж_номер", "ПОМ_Номер Квартиры", "ПОМ_Номер Помещения"]
		self.lbl2_names = ["S кв. [фактическая]", "S кв. [с коэффициентом]", "S кв. [жилая]", "S пом. [фактическая]", "S пом. [с коэффициентом]"]
		self.ttip2 = ["Параметр тип: Площадь", "Параметр тип: Площадь", "Параметр тип: Площадь", "Параметр тип: Площадь", "Параметр тип: Площадь"]
		self.def_values2 = ["ПОМ_Квартира.площадь.общая", "ПОМ_Квартира.площадь.общая_к", "ПОМ_Квартира.площадь.жилая", "КП_П_Площадь", "ПОМ_Квартира.комната.площадь_к"]
		self.lbl3_names = ["Номер квартиры (абс.)", "Описание квартиры" , "Код квартиры"]
		self.ttip3 = ["Параметр тип: Текст", "Параметр тип: Текст" , "Параметр тип: Текст"]
		self.def_values3 = ["квартира_номер_абсолютный", "ПОМ_Кваритра.наименование" , "ПОМ_Квартира.код"]
		self.cbx = []
		self.cbx2 = []
		self.cbx3 = []
		self.num = 0
		self.parameters_area = []
		self.parameters_text = []
		self.parameters_all = []
		self.f_out = []
		self.temp_constructor = "inactive"
		self.MinimumSize = Size(800, 445)
		self.MaximumSize = Size(800, 445)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.AllowDrop = True
		self.f_rooms_params = []
		self.CenterToScreen()
		self.MouseClick += self.OnMouseClick

		#TAKE ROOM PARAMETERS TO UI
		for b in self.rooms:
			for j in b.Parameters:
				if j.IsShared and j.UserModifiable and not j.IsReadOnly:
					if str(j.StorageType) == "Double":
						if str(j.DisplayUnitType) == "DUT_SQUARE_METERS":
							if str(j.Definition.Name) not in self.parameters_area: self.parameters_area.append(str(j.Definition.Name))
						if str(j.Definition.Name) not in self.parameters_all:
							self.parameters_all.append(str(j.Definition.Name))
					elif str(j.StorageType) == "String":
						if str(j.Definition.Name) not in self.parameters_text:
							self.parameters_text.append(str(j.Definition.Name))
						if str(j.Definition.Name) not in self.parameters_all:
							self.parameters_all.append(str(j.Definition.Name))
					elif str(j.Definition.Name) not in self.parameters_all:
						self.parameters_all.append(str(j.Definition.Name))
				elif str(j.StorageType) == "String" and j.HasValue and j.UserModifiable and not j.IsReadOnly:
					if str(j.Definition.Name) not in self.parameters_all:
						self.parameters_all.append(str(j.Definition.Name))
		self.parameters_text.append("Назначение")
		self.parameters_text.append("Комментарии")
		self.parameters_area.sort()
		self.parameters_text.sort()
		self.parameters_all.sort()

		tooltip = ToolTip()

		gb = GroupBox()
		gb.Text = "Формула номера помещения:"
		gb.Size = Size(765, 70)
		gb.Location = Point(10, 10)
		gb.Parent = self
		gb.ForeColor = Color.FromArgb(0,0,0,255)

		self.dd = Label(Text = "[Drag and Drop]")
		self.dd.Parent = gb
		self.dd.Location = Point(10, 23)
		self.dd.Size = Size(730,30)
		self.dd.DragDrop += self.OnDragDrop
		self.dd.DragEnter += self.OnDragEnter
		self.dd.MouseClick += self.OnMouseClick2
		self.dd.AllowDrop = True
		self.dd.TextAlign = ContentAlignment.MiddleCenter
		self.dd.ForeColor = Color.FromArgb(0,0,0,255)

		self.lb.append(ListBox())
		self.lb[0].Parent = self
		self.lb[0].AllowDrop = True
		self.lb[0].Size = Size(765, 100)
		self.lb[0].Location = Point(10,90)
		for z in range(len(self.parameters_all)):
			self.lb[0].Items.Add(self.parameters_all[z])
		self.lb[0].MouseDown += self.OnMousDown

		gb_rooms = GroupBox()
		gb_rooms.Text = "Квартирография:"
		gb_rooms.Size = Size(765, 170)
		gb_rooms.Location = Point(10, 195)
		gb_rooms.Parent = self

		self.d_wdth = 280

		self.CenterToScreen()

		for i in range(len(self.lbl_names)):
			self.num = i*5 + (700/(len(self.lbl_names)-1) - 5 * (len(self.lbl_names) + 1)) * i
			lblgb1 = Label(Text = self.lbl_names[i])
			lblgb1.Parent = gb_rooms
			lblgb1.Location = Point(15 + self.num, 15)
			lblgb1.AutoSize = True
			tooltip.SetToolTip(lblgb1, self.ttip[i])
			self.values_2.append("")
			self.cbx.append(ComboBox())
			self.cbx[i].Parent = gb_rooms
			self.cbx[i].Location = Point(15 + self.num, 35)
			self.num = 765/len(self.lbl_names) - 5 * (len(self.lbl_names) + 1)
			self.cbx[i].Size = Size(15 + self.num,15 + self.num)
			self.cbx[i].DropDownWidth = self.d_wdth
			self.cbx[i].Items.Add("")
			self.cbx[i].Items.Add("........................................................................................")
			for z in self.parameters_text:
				self.cbx[i].Items.Add(z)
			if self.def_values[i] in self.parameters_text:
				self.cbx[i].Text = self.def_values[i]
				#self.cbx[i].Enabled = False
				self.values_2[i] = self.def_values[i]
			else: self.cbx[i].Text = ""
			self.cbx[i].DropDownStyle = ComboBoxStyle.DropDownList
			self.cbx[i].SelectionChangeCommitted += self.CblOnChanged
			self.cbx[i].SelectedIndexChanged += self.CblOnChanged

			if i == len(self.lbl_names)-1:
				self.cbx[i].FlatStyle = FlatStyle.Flat

		for i in range(len(self.lbl2_names)):
			self.num = i*5 + (700/(len(self.lbl2_names)-1) - 5 * (len(self.lbl2_names)+1)) * i
			lblgb1 = Label(Text = self.lbl2_names[i])
			lblgb1.Parent = gb_rooms
			lblgb1.Location = Point(15 + self.num, 65)
			lblgb1.AutoSize = True
			tooltip.SetToolTip(lblgb1, self.ttip2[i])
			self.values_3.append("")
			self.cbx2.append(ComboBox())
			self.cbx2[i].Parent = gb_rooms
			self.cbx2[i].FlatStyle = FlatStyle.Flat
			self.cbx2[i].Location = Point(15 + self.num, 85)
			self.num = 765/len(self.lbl2_names) - 5 * (len(self.lbl2_names)+1)
			self.cbx2[i].Size = Size(15 + self.num,15 + self.num)
			self.cbx2[i].DropDownWidth = self.d_wdth
			self.cbx2[i].Items.Add("")
			self.cbx2[i].Items.Add("........................................................................................")
			for z in self.parameters_area:
				self.cbx2[i].Items.Add(z)		
			if self.def_values2[i] in self.parameters_area:
				self.cbx2[i].Text = self.def_values2[i]
				#self.cbx2[i].Enabled = False
				self.values_3[i] = self.def_values2[i]
			else: self.cbx2[i].Text = ""
			self.cbx2[i].DropDownStyle = ComboBoxStyle.DropDownList
			self.cbx2[i].SelectionChangeCommitted += self.CblOnChanged
			self.cbx2[i].SelectedIndexChanged += self.CblOnChanged

		for i in range(len(self.lbl3_names)):
			self.num = i*5 + (535/(len(self.lbl3_names)-1) - 5 * (len(self.lbl3_names)+1)) * i
			lblgb1 = Label(Text = self.lbl3_names[i])
			lblgb1.Parent = gb_rooms
			lblgb1.Location = Point(15 + self.num, 115)
			lblgb1.AutoSize = True
			tooltip.SetToolTip(lblgb1, self.ttip3[i])
			self.values_4.append("")
			self.cbx3.append(ComboBox())
			self.cbx3[i].Parent = gb_rooms
			self.cbx3[i].FlatStyle = FlatStyle.Flat
			self.cbx3[i].Location = Point(15 + self.num, 135)
			self.num = 765/len(self.lbl3_names) - 5 * (len(self.lbl3_names)+1)
			self.cbx3[i].Size = Size(self.num, self.num)
			self.cbx3[i].DropDownWidth = self.d_wdth
			self.cbx3[i].Items.Add("")
			self.cbx3[i].Items.Add("........................................................................................")
			for z in self.parameters_text:
				self.cbx3[i].Items.Add(z)
			if self.def_values3[i] in self.parameters_text:
				self.cbx3[i].Text = self.def_values3[i]
				#self.cbx3[i].Enabled = False
				self.values_4[i] = self.def_values3[i]
			else: self.cbx3[i].Text = ""
			self.cbx3[i].DropDownStyle = ComboBoxStyle.DropDownList
			self.cbx3[i].SelectionChangeCommitted += self.CblOnChanged
			self.cbx3[i].SelectedIndexChanged += self.CblOnChanged

		button1 = Button()
		button1.Parent = self
		button1.Text = "Запуск"
		button1.Location = Point(10, 375)
		button1.Click += self.OnClick_1

		tooltip.SetToolTip(button1, "Запустить скрипт")  

		button2 = Button()
		button2.Parent = self
		button2.Text = "Закрыть"
		button2.Location = Point(100, 375)
		button2.Click += self.OnClick_2

		tooltip.SetToolTip(button2, "Отменить и закрыть форму")

		self.CenterToScreen()

	def logger(self, result):
		try:
			pass
		except: pass

	def OnMousDown(self, sender, event):
		sender.DoDragDrop(sender.Text, DragDropEffects.Copy)

	def OnMouseClick2(self, sender, args):
		try:
			if sender == self.dd:
				self.temp_constructor = "inactive"
				sender.Text = "[Drag and Drop]"
				self.values_1 = []
		except: pass

	def OnMouseClick(self, args):
		try: pass
		except: pass

	def OnDragEnter(self, sender, event):
		event.Effect = DragDropEffects.Copy

	def CblOnChanged(self, sender, event):
		for i in range (len(self.cbx)):
			if self.cbx[i] == sender:
				if sender.Text != "........................................................................................":
					self.values_2[i] = sender.Text
				else:
					sender.Text = ""
					self.values_2[i] = ""
		for i in range (len(self.cbx2)):
			if self.cbx2[i] == sender:
				if sender.Text != "........................................................................................":
					self.values_3[i] = sender.Text
				else:
					sender.Text = ""
					self.values_3[i] = ""
		for i in range (len(self.cbx3)):
			if self.cbx3[i] == sender:
				if sender.Text != "........................................................................................":
					self.values_4[i] = sender.Text
				else:
					sender.Text = ""
					self.values_4[i] = ""

	def OnDragDrop(self, sender, event):
		sender.ForeColor = Color.FromArgb(0,0,0,255)
		if self.temp_constructor == "inactive":
			sender.Text = event.Data.GetData(DataFormats.Text)
			self.values_1.append(event.Data.GetData(DataFormats.Text))
			self.temp_constructor = "one"
		elif self.temp_constructor == "one":
			sender.Text = sender.Text + "  .  " + event.Data.GetData(DataFormats.Text)
			self.values_1.append(event.Data.GetData(DataFormats.Text))
			self.temp_constructor = "two"
		elif self.temp_constructor == "two":
			sender.Text = sender.Text + "  .  " + event.Data.GetData(DataFormats.Text)
			self.values_1.append(event.Data.GetData(DataFormats.Text))
			self.temp_constructor = "three"
		elif self.temp_constructor == "three":
			sender.Text = sender.Text + "  .  " + event.Data.GetData(DataFormats.Text)
			self.values_1.append(event.Data.GetData(DataFormats.Text))
			self.temp_constructor = "four"
		elif self.temp_constructor == "four":
			sender.Text = sender.Text + "  .  " + event.Data.GetData(DataFormats.Text)
			self.values_1.append(event.Data.GetData(DataFormats.Text))
			self.temp_constructor = "five"
		else:
			sender.Text = event.Data.GetData(DataFormats.Text)
			self.temp_constructor = "one"
			self.values_1 = []
			self.values_1.append(event.Data.GetData(DataFormats.Text))

	def OnClick_1(self, sender, args):
		if self.dd.Text == "[Drag and Drop]":
			self.dd.ForeColor = Color.FromArgb(0,255,0,0)
			pass
		else:
			self.dd.ForeColor = Color.FromArgb(0,0,0,255)
			self.check = []
			self.check_result = True
			for i in self.values_2:
				self.check.append(i)
			for i in self.values_3:
				self.check.append(i)
			for i in self.values_4:
				self.check.append(i)
			for i in self.check:
				if i == "": self.check_result = False
			if self.check_result:
				self.f_out.append(self.values_1)
				self.f_out.append(self.values_2)
				self.f_out.append(self.values_3)
				self.f_out.append(self.values_4)
				self.Enabled = False
				try:
					with db.Transaction(name = "Calculate room data"):
						self.runscript()
				except: 
					self.Show()
					self.Enabled = True
					self.Text = "[UNI'S]   Квартирография"
					self.CenterToScreen()
					self.Close()
				self.Show()
				self.Enabled = True
				self.Text = "[UNI'S]   Квартирография"
				self.CenterToScreen()
				self.Close()
			else:
				forms.alert("Ошибка:\nЗаполните все параметры!")
				self.logger("Ошибка:\nЗаполните все параметры!")
				self.f_out.append("close")
				self.Close()

	def OnClick_2(self, sender, args):
		self.f_out.append("close")
		self.Close()

	def runscript(self): 
		self.Text = "[UNI'S]   initialize"
		self.IN = self.f_out
		self.department_par = DB.BuiltInParameter.ROOM_DEPARTMENT
		self.name_par = DB.BuiltInParameter.ROOM_NAME
		self.collector_rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
		self.collector_levels = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
		self.levels = []
		self.allrooms = []
		self.rooms = []
		self.rooms_sorted = []
		self.rooms_keys = []
		self.rooms_keys_sorted = []
		self.rooms_other_keys = []
		self.rooms_other_keys_sorted = []
		self.rooms_other = []
		self.rooms_other_sorted = []

		for level in self.collector_levels:
			self.levels.append(level)
		for room in self.collector_rooms:
			try:
				if room.get_Parameter(self.name_par).AsString() != "":
					if room.Area > 0:
						if room.get_Parameter(self.department_par).AsString() == "Кв":
							self.rooms.append(room)
						else: self.rooms_other.append(room)			
						self.allrooms.append(room)
			except: pass

		self.values = self.IN[0]
		self.rooms_korpus = self.IN[1][0]
		self.rooms_section = self.IN[1][1]
		self.rooms_elevate = self.IN[1][2]
		self.rooms_flatnum = self.IN[1][3]
		self.rooms_roomnum = self.IN[1][4]
		self.flat_num_abs = self.IN[3][0]
		self.flat_name = self.IN[3][1]
		self.flat_description = self.IN[3][2]
		self.area_flat_fact = self.IN[2][0]
		self.area_flat_k = self.IN[2][1]
		self.area_flat_living = self.IN[2][2]
		self.area_room_fact = self.IN[2][3]
		self.area_room_k = self.IN[2][4]

		self.rooms_dict = []
		self.rooms_dict_r = []
		self.rooms_other_dict = []
		self.rooms_other_dict_r = []
		self.rooms_string1 = ""
		self.rooms_amount = 0
		self.rooms_string2 = ""	
		if not self.p_check():
			self.Hide()
			forms.alert("Ошибка!\nОстались незаполненными значения параметров:\n" + self.check_values)
			self.logger("Ошибка!\nОстались незаполненными значения параметров:\n" + self.check_values)
			self.Close()
			pass
		for room1 in self.rooms:			
			self.rooms_sorted.append(self.rooms[0])
			if self.rooms_korpus != "":
				self.rooms_keys.append(room1.LookupParameter(self.rooms_korpus).AsString() + "." + room1.LookupParameter(self.rooms_section).AsString() + "." + room1.LookupParameter(self.rooms_elevate).AsString() + "." + room1.LookupParameter(self.rooms_flatnum).AsString() + "." + str(room1.Id))
				self.Text = "[UNI'S]   append keys 1/2:" + room1.LookupParameter(self.rooms_korpus).AsString() + "." + room1.LookupParameter(self.rooms_section).AsString() + "." + room1.LookupParameter(self.rooms_elevate).AsString() + "." + room1.LookupParameter(self.rooms_flatnum).AsString() + "." + str(room1.Id)
			else:
				self.rooms_keys.append(room1.LookupParameter(self.rooms_section).AsString() + "." + room1.LookupParameter(self.rooms_elevate).AsString() + "." + room1.LookupParameter(self.rooms_flatnum).AsString() + "." + str(room1.Id))
				self.Text = "[UNI'S]   append keys 1/2:" + room1.LookupParameter(self.rooms_section).AsString() + "." + room1.LookupParameter(self.rooms_elevate).AsString() + "." + room1.LookupParameter(self.rooms_flatnum).AsString() + "." + str(room1.Id)
		for i in self.rooms_keys:
			self.rooms_keys_sorted.append(i)
			self.Text = "[UNI'S]   appended:" + str(i)
		for room2 in self.rooms_other:
			self.rooms_other_sorted.append(self.rooms_other[0])
			if self.rooms_korpus != "":
				self.rooms_other_keys.append(room2.LookupParameter(self.rooms_korpus).AsString() + "." + room2.LookupParameter(self.rooms_section).AsString() + "." + room2.LookupParameter(self.rooms_elevate).AsString() + ".0." + str(room2.Id))
				self.Text = "[UNI'S]   append keys 2/2:" + room2.LookupParameter(self.rooms_korpus).AsString() + "." + room2.LookupParameter(self.rooms_section).AsString() + "." + room2.LookupParameter(self.rooms_elevate).AsString() + ".0." + str(room2.Id)
			else:
				self.rooms_other_keys.append(room2.LookupParameter(self.rooms_section).AsString() + "." + room2.LookupParameter(self.rooms_elevate).AsString() + ".0." + str(room2.Id))
				self.Text = "[UNI'S]   append keys 2/2:" + room2.LookupParameter(self.rooms_section).AsString() + "." + room2.LookupParameter(self.rooms_elevate).AsString() + ".0." + str(room2.Id)
		for i in self.rooms_other_keys:
			self.rooms_other_keys_sorted.append(i)
		self.rooms_keys_sorted.sort()
		self.rooms_other_keys_sorted.sort()
		for i in range(len(self.rooms)):
			self.rooms_sorted[self.rooms_keys_sorted.index(self.rooms_keys[i])] = self.rooms[i]
			self.Text = "[UNI'S]   sorting 1/4:#" + str(i)
		for i in range(len(self.rooms_other)):
			self.rooms_other_sorted[self.rooms_other_keys_sorted.index(self.rooms_other_keys[i])] = self.rooms_other[i]
			self.Text = "[UNI'S]   sorting 2/4:#" + str(i)
		for r in range(len(self.rooms_sorted)):
			self.Text = "[UNI'S]   sorting 3/4:#" + str(r)
			if self.rooms_korpus != "":
				self.key = self.rooms_sorted[r].LookupParameter(self.rooms_korpus).AsString() + "." + self.rooms_sorted[r].LookupParameter(self.rooms_section).AsString() + "." + self.rooms_sorted[r].LookupParameter(self.rooms_elevate).AsString() + "." + self.rooms_sorted[r].LookupParameter(self.rooms_flatnum).AsString()
			else:
				self.key = self.rooms_sorted[r].LookupParameter(self.rooms_section).AsString() + "." + self.rooms_sorted[r].LookupParameter(self.rooms_elevate).AsString() + "." + self.rooms_sorted[r].LookupParameter(self.rooms_flatnum).AsString()
			if self.key not in self.rooms_dict:
				self.rooms_dict.append(self.key)
				self.rooms_dict_r.append([])
			self.rooms_dict_r[self.rooms_dict.index(self.key)].append(self.rooms_sorted[r])
		for r in range(len(self.rooms_other_sorted)):
			self.Text = "[UNI'S]   sorting 4/4:#" + str(r)
			if self.rooms_korpus != "":
				self.key = self.rooms_other_sorted[r].LookupParameter(self.rooms_korpus).AsString() + "." + self.rooms_other_sorted[r].LookupParameter(self.rooms_section).AsString() + "." + self.rooms_other_sorted[r].LookupParameter(self.rooms_elevate).AsString()
			else:
				self.key = self.rooms_other_sorted[r].LookupParameter(self.rooms_section).AsString() + "." + self.rooms_other_sorted[r].LookupParameter(self.rooms_elevate).AsString()
			if self.key not in self.rooms_other_dict:
				self.rooms_other_dict.append(self.key)
				self.rooms_other_dict_r.append([])
			self.rooms_other_dict_r[self.rooms_other_dict.index(self.key)].append(self.rooms_other_sorted[r])
		#ROOMNUMBER:
		for i in range(len(self.rooms_dict)):
			self.num_room = 0
			for z in self.rooms_dict_r[i]:
				self.par = z.LookupParameter(self.rooms_roomnum)
				self.par.Set(str(self.num_room))
				self.Text = "[UNI'S]   write num (==flats):" + str(i) + "/" + str(z.GetTypeId()) + "_" + str(self.num_room)
				self.num_room += 1
		for i in range(len(self.rooms_other_dict)):
			self.num_room = 0
			for z in self.rooms_other_dict_r[i]:
				self.par2 = z.LookupParameter(self.rooms_roomnum)
				self.par2.Set(str(self.num_room))
				self.Text = "[UNI'S]   write num (!=flats):" + str(i) + "/" + str(z.GetTypeId()) + "_null"
				self.num_room += 1
		#FLAT NUMBER:
		self.num_flat_abs = 1	
		for i in range(len(self.rooms_dict)):
			for z in self.rooms_dict_r[i]:
				self.par3 = z.LookupParameter(self.flat_num_abs)
				self.par3.Set(str(self.num_flat_abs))
				self.Text = "[UNI'S]   write  flat numabs (==flats):" + str(i) + "/" + str(z.GetTypeId()) + "_" + str(self.flat_num_abs)
			self.num_flat_abs += 1
		for i in range(len(self.rooms_other_dict)):
			for z in self.rooms_other_dict_r[i]:
				self.par4 = z.LookupParameter(self.flat_num_abs)
				self.par4.Set(str(0))
				self.Text = "[UNI'S]   write flat numabs (!=flats):" + str(i) + "/" + str(z.GetTypeId()) + "_null"
		#AREAS [NOT FLATS]:
		for room in self.rooms_other_sorted:
			self.room_a = round(room.Area * 0.09290304, 2)
			self.Text = "[UNI'S]   area write null (!=flats):" + str(room.GetTypeId())
			if self.area_flat_fact != "":
				self.par = room.LookupParameter(self.area_flat_fact)
				self.par.Set(0)
			if self.area_flat_k != "":
				self.par = room.LookupParameter(self.area_flat_k)
				self.par.Set(0)
			if self.area_flat_living != "":
				self.par = room.LookupParameter(self.area_flat_living)
				self.par.Set(0)
			if self.area_room_fact != "":
				self.par = room.LookupParameter(self.area_room_fact)
				self.par.Set(self.room_a / 0.09290304)
			if self.area_room_k != "":
				self.par = room.LookupParameter(self.area_room_k)
				self.par.Set(self.room_a / 0.09290304)
		#[FLAT]:
		for i in range(len(self.rooms_dict)):
			self.temp_area_flat_k = 0
			self.temp_area_flat_f = 0
			self.temp_area_flat_living = 0
			self.rooms_string1 = ""
			self.rooms_amount = 0
			self.rooms_string2 = ""
			for z in self.rooms_dict_r[i]:
				self.room_a = float()
				self.room_a = round(z.Area * 0.09290304 / 10, 2)
				self.room_half = round(z.Area * 0.09290304 / 2, 2)
				self.room_full = round(z.Area * 0.09290304, 2)
				self.par = z.LookupParameter(self.area_room_fact)
				if z.get_Parameter(self.name_par).AsString() == "Лоджия":
					self.area = math.floor(self.room_half * 100)/100
					self.temp_area_flat_k += self.area
					self.par.Set(self.area * 2 / 0.09290304)
				elif z.get_Parameter(self.name_par).AsString() == "Терраса" or z.get_Parameter(self.name_par).AsString() == "Балкон":
					self.area = math.floor(self.room_a * 3 * 100)/100
					self.temp_area_flat_k += self.area
					self.par.Set(self.room_a * 10 / 0.09290304)
				else:
					self.area = self.room_full
					self.temp_area_flat_k += self.area
					self.temp_area_flat_f += self.area
					self.par.Set(self.area / 0.09290304)
				self.par2 = z.LookupParameter(self.area_room_k)
				self.par2.Set(self.area / 0.09290304)
				if z.get_Parameter(self.name_par).AsString() == "Жилая комната" or z.get_Parameter(self.name_par).AsString() == "Гостиная" or z.get_Parameter(self.name_par).AsString() == "Кухня-гостиная":
					self.rooms_amount += 1
					self.temp_area_flat_living += self.room_full
				if z.get_Parameter(self.name_par).AsString() == "Кухня-гостиная":
					self.rooms_string1 = "Студия"
				self.Text = "[UNI'S]   area write " + str(self.area) + "/" + str(room.GetTypeId())
	
				if self.rooms_amount == 0:
					self.rooms_string1 = "Нет жилых помещений"
					self.rooms_string2 = "ER#MIN"
				elif self.rooms_amount == 1 and self.rooms_string1 == "Студия":
					self.rooms_string2 = "C"
				elif self.rooms_amount == 1:
					self.rooms_string1 = "Однокомнатная квартира"
					self.rooms_string2 = "1"
				elif self.rooms_amount == 2:
					self.rooms_string1 = "Двухкомнатная квартира"
					self.rooms_string2 = "2"
				elif self.rooms_amount == 3:
					self.rooms_string1 = "Трехкомнатная квартира"
					self.rooms_string2 = "3"
				elif self.rooms_amount == 4:
					self.rooms_string1 = "Четырехкомнатная квартира"
					self.rooms_string2 = "4"
				elif self.rooms_amount == 5:
					self.rooms_string1 = "Пятикомнатная квартира"
					self.rooms_string2 = "5"
				elif self.rooms_amount == 6:
					self.rooms_string1 = "Шестикомнатная квартира"
					self.rooms_string2 = "6"
				elif self.rooms_amount == 7:
					self.rooms_string1 = "Семикомнатная квартира"
					self.rooms_string2 = "7"
				elif self.rooms_amount == 8:
					self.rooms_string1 = "Восьмикомнатная квартира"
					self.rooms_string2 = "8"
				elif self.rooms_amount == 9:
					self.rooms_string1 = "Девятикомнатная квартира"
					self.rooms_string2 = "9"
				else:
					self.rooms_string1 = "living_rooms_#" + str(self.rooms_amount)
					self.rooms_string2 = "ER#MAX"
				self.Text = "[UNI'S]   flat: «" + str(self.rooms_string1) + "» detected"

			for x in self.rooms_dict_r[i]:
				if self.area_flat_fact != "":
					self.par = x.LookupParameter(self.area_flat_fact)
					self.par.Set(self.temp_area_flat_f / 0.09290304)
				if self.area_flat_k != "":
					self.par = x.LookupParameter(self.area_flat_k)
					self.par.Set(self.temp_area_flat_k / 0.09290304)
				if self.area_flat_living != "":
					self.par = x.LookupParameter(self.area_flat_living)
					self.par.Set(self.temp_area_flat_living / 0.09290304)
				if self.flat_name != "":
					self.par2 = x.LookupParameter(self.flat_name)
					self.par2.Set(self.rooms_string1)
				if self.flat_description != "":
					self.par2 = x.LookupParameter(self.flat_description)
					self.par2.Set(self.rooms_string2)
		#CONSTRUCTOR
		for room in self.allrooms:
			self.string = ""
			for value in self.values:
				if self.string != "": self.string += "." + room.LookupParameter(value).AsString()
				else: self.string += room.LookupParameter(value).AsString()
				self.par = room.LookupParameter("Номер")
			self.par.Set(self.string)
			self.Text = "[UNI'S]   number constructor: «" + self.string + "» appended!"
		self.Text = "[UNI'S]   Done!"
		#self.Close()
		self.Hide()
		forms.alert("Обработано " + str(len(self.allrooms)) + " помещений!")
		self.logger("Обработано " + str(len(self.allrooms)) + " помещений!")

	def p_check(self):
		self.check_values = ""
		for room in self.allrooms:	
			for j in room.Parameters:
				if j.Definition.Name == self.rooms_korpus:
					if not j.HasValue:
						if j.Definition.Name not in self.check_values:
							if self.check_values != "": self.check_values += " "
							self.check_values += "«" + j.Definition.Name + "»"
				elif j.Definition.Name == self.rooms_section:
					if not j.HasValue:
						if j.Definition.Name not in self.check_values:
							if self.check_values != "": self.check_values += " "
							self.check_values += "«" + j.Definition.Name + "»"				
				elif j.Definition.Name == self.rooms_elevate:
					if not j.HasValue:
						if j.Definition.Name not in self.check_values:
							if self.check_values != "": self.check_values += " "
							self.check_values += "«" + j.Definition.Name + "»"				
				elif j.Definition.Name == self.rooms_flatnum:
					if not j.HasValue:
						if j.Definition.Name not in self.check_values:
							if self.check_values != "": self.check_values += " "
							self.check_values += "«" + j.Definition.Name + "»"				
		if self.check_values != "": return False
		else: return True

		#self.Close()

form = CreateWindow()

Application.Run(form)