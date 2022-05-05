# -*- coding: utf-8 -*-
"""
CreateGroups

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Сгруппировать элементы по помещениям"
__doc__ = 'Запись списка помещений, которым принадлежат элементы.' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import revit, DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
import System
from System.Windows.Forms import *
from System.Drawing import *

class PickAddParameterWindow(Form):
	def __init__(self):
		global setting_custom_number
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
		global setting_custom_number
		setting_custom_number = ""
		self.Close()

	def OnOk(self, sender, args):
		global setting_custom_number
		setting_custom_number = self.combobox_param.Text
		self.Close()

class PickParameterWindow(Form):
	def __init__(self):
		global setting_custom_parameter
		self.Name = "KPLN_Finishing"
		self.Text = "KPLN Параметры группировки элементов"
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
		self.button_pass.Text = "Пропустить"
		self.button_pass.Click += self.OnPass

		self.label_faq.Parent = self
		self.label_faq.Location = Point(12, 13)
		self.label_faq.Size = Size(260, 38)
		self.label_faq.Text = "Определите параметр элементов для группировки:"

		parameters = []
		wparameters = []
		fparameters = []
		cparameters = []
		for j in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().FirstElement().Parameters:
			try:
				if str(j.StorageType) == "String":
					wparameters.append(j.Definition.Name)
			except: pass
		for j in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().FirstElement().Parameters:
			try:
				if str(j.StorageType) == "String":
					fparameters.append(j.Definition.Name)
			except: pass
		for j in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().FirstElement().Parameters:
			try:
				if str(j.StorageType) == "String":
					cparameters.append(j.Definition.Name)
			except: pass
			for i in wparameters:
				if i in fparameters and i in cparameters and not i in parameters:
					parameters.append(i)
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
		global setting_custom_parameter
		setting_custom_parameter = []
		self.Close()

	def OnOk(self, sender, args):
		global setting_custom_parameter
		if self.cb1.Text != "<неактивен>":
			setting_custom_parameter.append(self.cb1.Text)
		if self.cb2.Text != "<неактивен>":
			setting_custom_parameter.append(self.cb2.Text)
		if self.cb3.Text != "<неактивен>":
			setting_custom_parameter.append(self.cb3.Text)
		self.Close()

class Stack():
	def __init__(self, element, room):
		global GetType
		self.Elements = []
		self.Elements.append(element)
		self.TypeId = GetType(element).Id.IntegerValue
		self.Rooms = []
		self.Keys = []
		self.Keys_numbers = []
		self.Rooms.append(room)
		self.Keys.append(room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString())
		self.Keys_numbers.append(room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())

	def GetKey(self):
		global setting_custom_parameter
		dict = []
		if len(setting_custom_parameter) != 0:
			for i in setting_custom_parameter:
				dict.append(str(self.Elements[0].LookupParameter(i).AsString()))
			return "{}{}".format("-".join(dict), str(GetType(self.Elements[0]).Id.IntegerValue))
		else:
			return str(GetType(element).Id.IntegerValue)

	def AddElement(self, element):
		self.Elements.append(element)

	def AddRoom(self, room):
		global setting_custom_number
		try:
			self.Rooms.append(room)
			r_key = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
			if setting_custom_number == "":
				rnum_key = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()
			else:
				rnum_key = room.LookupParameter(setting_custom_number).AsString()
			if not r_key in self.Keys:
				self.Keys.append(r_key)
			if not rnum_key in self.Keys_numbers:
				self.Keys_numbers.append(rnum_key)
		except Exception as e:
			print(str(e))


	def SetKey(self):
		name = ", ".join(self.Keys)
		numbers = ", ".join(self.Keys_numbers)
		for element in self.Elements:
			element.LookupParameter("О_Группа").Set("{}\n\n{}".format(name, numbers))

def GetType(element):
	try:
		type = element.WallType
		if type == None: raise Exception
		return type
	except:
		try:
			type = element.FloorType
			if type == None: raise Exception
			return type
		except:
			try:
				type = element.GetTypeId()
				if type == None: 
					raise Exception
				etype = doc.GetElement(type)
				if etype == None: 
					raise Exception
				return etype
			except:
				return False

def GetStack(element):
	dict = []
	try:
		if len(setting_custom_parameter) != 0:
			for i in setting_custom_parameter:
				dict.append(str(element.LookupParameter(i).AsString()))
			for stack in stacks:
				if stack.GetKey() == "{}{}".format("-".join(dict), str(GetType(element).Id.IntegerValue)):
					return stack
		else:
			for stack in stacks:
				if stack.GetKey() == str(GetType(element).Id.IntegerValue):
					return stack
		return False
	except Exception as e:
		print(str(e))
		return False



stacks = []
setting_custom_parameter = []
setting_custom_number = ""

form_pr = PickParameterWindow()
Application.Run(form_pr)

form_pr2 = PickAddParameterWindow()
Application.Run(form_pr2)

for collector in [DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements(), 
				  DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements(), 
				  DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()]:
	for element in collector:
		try:
			if element.LookupParameter("О_Id помещения").AsString() != "":
				value = element.LookupParameter("О_Id помещения").AsString()
				if value == None: continue
				elementId = DB.ElementId(int(value))
				room = doc.GetElement(elementId)
				stack = GetStack(element)
				if stack != False:
					stack.AddElement(element)
					stack.AddRoom(room)
				else:
					new_stack = Stack(element, room)
					stacks.append(new_stack)
		except Exception as e:
			print(str(e))

if len(stacks) != 0:
	for stack in stacks:
		with db.Transaction(name = "KPLN: Определить группу"):
			try:
				stack.SetKey()
			except Exception as e:
				print(str(e))
				break