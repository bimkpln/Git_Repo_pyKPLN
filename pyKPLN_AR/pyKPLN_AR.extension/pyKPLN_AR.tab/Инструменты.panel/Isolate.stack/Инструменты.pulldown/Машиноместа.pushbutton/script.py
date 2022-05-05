# -*- coding: utf-8 -*-
"""
MarkAuto

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Маркировать машиноместа"
__doc__ = 'Маркировка машиномест по сплайну с указанием префикса, постфикса и начального номера' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from rpw.ui import Pick
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import TextInput, Alert, select_folder
import datetime
import System
from System.Windows.Forms import *
from System.Drawing import *

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "KPLN Машиноместа"
		self.Text = "KPLN Машиноместа"
		self.Size = Size(418, 608)
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
		self.button_create = Button(Text = "Ок")
		self.ControlBox = True
		self.TopMost = True
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.MinimumSize = Size(418, 240)
		self.MaximumSize = Size(418, 240)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.CenterToScreen()
		self.BackColor = Color.FromArgb(255, 255, 255)
		self.value = "Выберите параметр"
		self.el = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Parking).WhereElementIsNotElementType().FirstElement()
		self.params = []
		for j in self.el.Parameters:
			if str(j.Definition.ParameterType) == "Text":
				self.params.append(j.Definition.Name)
		self.params.sort()
		self.lb1 = Label(Text = "Выберите параметр для записи марки")
		self.lb1.Parent = self
		self.lb1.Size = Size(343, 15)
		self.lb1.Location = Point(30, 30)

		self.lb2 = Label(Text = "Необходимо выбрать текстовый параметр.\nРекомендуется использовать системный параметр «Марка».")
		self.lb2.Parent = self
		self.lb2.Size = Size(343, 40)
		self.lb2.Location = Point(30, 75)

		self.btn_confirm = Button(Text = "Далее")
		self.btn_confirm.Parent = self
		self.btn_confirm.Location = Point(30, 150)
		self.btn_confirm.Enabled = False
		self.btn_confirm.Click += self.OnOk

		self.cb = ComboBox()
		self.cb.Parent = self
		self.cb.Size = Size(343, 100)
		self.cb.Location = Point(30, 50)
		self.cb.DropDownStyle = ComboBoxStyle.DropDownList
		self.cb.SelectedIndexChanged += self.allow_ok
		self.cb.Items.Add("Выберите параметр")
		self.cb.Text = "Выберите параметр"
		for par in self.params:
			self.cb.Items.Add(par)

	def OnOk(self, sender, args):
		self.Close()

	def allow_ok(self, sender, event):
		if self.cb.Text != "Выберите параметр":
			self.btn_confirm.Enabled = True
		else:
			self.btn_confirm.Enabled = False
		self.value = self.cb.Text

class AutoMobileSelectionFilter (UI.Selection.ISelectionFilter):
	def AllowElement(self, element = DB.Element):
		if element.Category.Id == DB.ElementId(-2001180) and ("машиномест" in str(element.Symbol.FamilyName).lower() or "машиномест" in str(element.Name).lower()):
			return True
		return False
	def AllowReference(self, refer, point):
		return False

class SplineSelectionFilter (UI.Selection.ISelectionFilter):
	def AllowElement(self, element = DB.Element):
		if element.Category.Id == DB.ElementId(-2000051):
			return True
		return False
	def AllowReference(self, refer, point):
		return False
parameter_name = "Марка"
try:
	list_elements = []
	list_elements_sorted = []
	list_parameters = []
	list_parameters_sorted = []
	Alert("1. Выбрать линию\n\n2. Выбрать машиноместа\n\n3. Указать остальные необходимые параметры:\n   - префикс, постфикс, начало нумерации", title="KPLN Машиноместа", header = "Инструкция:")
	otype = UI.Selection.ObjectType.Element
	filter_spline = SplineSelectionFilter()
	filter_auto = AutoMobileSelectionFilter()
	spline = uidoc.Selection.PickObject(otype, filter_spline, "KPLN : Веберите сплайн")
	Alert("и нажмите кнопку «Готово»", title="KPLN Машиноместа", header = "Теперь выберите все машиноместа")
	elements = uidoc.Selection.PickObjects(otype, filter_auto, "KPLN : Веберите машиноместа")
	if spline and len(elements) != 0:
		line = doc.GetElement(spline)
		nurbs = line.Location.Curve
		with db.Transaction(name = "mark"):
			for e in elements:
				element = doc.GetElement(e)
				location = element.Location.Point
				pp = nurbs.Project(location).Parameter
				list_elements.append(element)
				list_parameters.append(pp)
				list_parameters_sorted.append(pp)
		list_parameters_sorted.sort()
		for i in list_parameters_sorted:
			for z in range(0, len(list_parameters)):
				if i == list_parameters[z]:
					list_elements_sorted.append(list_elements[z])
		num = 1
		prefix = TextInput('Префикс для маркировки', default = "", description = "Укажите префикс: <префикс><марка>", exit_on_close = False)
		postfix = TextInput('Постфикс для маркировки', default = "", description = "Укажите постфикс: <марка><постфикс>", exit_on_close = False)
		startnumber = TextInput('Начать нумерацию со значения', default = "0", description = "Примечание: только числовое значение!", exit_on_close = False)
		try:
			num = int(startnumber)
		except:
			Alert("Произошла ошибка при считывании номера, будет использовано значение по умолчанию «1»", title="KPLN Машиноместа", header = "Ошибка")
			num = 1
		form = CreateWindow()
		Application.Run(form)
		parameter_name = form.value
		with db.Transaction(name = "mark"):
			for element in list_elements_sorted:
				try:
					element.LookupParameter(parameter_name).Set(prefix + str(num) + postfix)
					num += 1
				except: pass
except: 
	pass