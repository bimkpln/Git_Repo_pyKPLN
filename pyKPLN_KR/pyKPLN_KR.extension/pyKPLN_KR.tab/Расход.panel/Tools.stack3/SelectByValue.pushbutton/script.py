# -*- coding: utf-8 -*-
"""
SelectByValue

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Выбрать по значению"
__doc__ = 'Выбирает все элементы с выбранными значениями' \

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
from System.Collections.Generic import *
import System
from System.Windows.Forms import *
from System.Drawing import *
from itertools import chain
import datetime
from rpw.ui.forms import TextInput
import datetime

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "Редактор наборов"
		self.Text = "Выберите значения расхода стали"
		self.Size = Size(418, 608)
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_KR (alpha)\\pyKPLN_KR.extension\\pyKPLN_KR.tab\\icon.ico")
		self.button_create = Button(Text = "Ок")
		self.ControlBox = True
		self.TopMost = True
		self.MinimumSize = Size(418, 480)
		self.MaximumSize = Size(418, 480)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.CenterToScreen()

		self.listbox = ListView()
		self.c_cb = ColumnHeader()
		self.c_cb.Text = ""
		self.c_cb.Width = -2
		self.c_cb.TextAlign = HorizontalAlignment.Center
		self.c_name = ColumnHeader()
		self.c_name.Text = "Расход"
		self.c_name.Width = -2
		self.c_name.TextAlign = HorizontalAlignment.Left

		self.SuspendLayout()
		self.listbox.Dock = DockStyle.Fill
		self.listbox.View = View.Details

		self.listbox.Parent = self
		self.listbox.Size = Size(400, 400)
		self.listbox.Location = Point(1, 1)
		self.listbox.FullRowSelect = True
		self.listbox.GridLines = True
		self.listbox.AllowColumnReorder = True
		self.listbox.Sorting = SortOrder.Ascending
		self.listbox.Columns.Add(self.c_cb)
		self.listbox.Columns.Add(self.c_name)
		self.listbox.LabelEdit = True
		self.listbox.CheckBoxes = True
		self.listbox.MultiSelect = True

		self.button_ok = Button(Text = "Ok")
		self.button_ok.Parent = self
		self.button_ok.Location = Point(10, 410)
		self.button_ok.Click += self.OnOk

		self.button_ok = Button(Text = "Закрыть")
		self.button_ok.Parent = self
		self.button_ok.Location = Point(100, 410)
		self.button_ok.Click += self.OnCancel

		self.Text = "Выберите значения для выбора"
		self.collector_elements = DB.FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
		self.item = []
		self.values = []

		for element in self.collector_elements:
			try:
				self.value = element.LookupParameter("КР_К_Армирование").AsString()
				if not self.in_list(self.values, self.value):
					self.values.append(self.value)
			except: pass
		self.values.sort()

		for v in self.values:
			self.item.append(ListViewItem())
			self.item[len(self.item)-1].Text = ""
			self.item[len(self.item)-1].Checked = False
			self.item[len(self.item)-1].SubItems.Add(str(v))
			self.listbox.Items.Add(self.item[len(self.item)-1])

	def in_list(self, list, item):
		for i in list:
			if i == item:
				return True
		return False

	def OnOk(self, sender, args):
		self.select_elements(self.define_values())
		self.Close()

	def define_values(self):
		list_of_values = []
		for i in self.item:
			if i.Checked:
				si = i.SubItems
				list_of_values.append(si[1].Text)
		return list_of_values

	def OnCancel(self, sender, args):
		self.Close()

	def select_elements(self, values):
		if len(values) != 0:
			elements = []
			active_view = doc.ActiveView.Id
			collector = DB.FilteredElementCollector(doc, active_view).WhereElementIsNotElementType().ToElements()
			for element in collector:
				try:
					if self.in_list(values, element.LookupParameter("КР_К_Армирование").AsString()):
						elements.append(element)
				except: pass
			if len(elements) != 0:
				selection = uidoc.Selection
				collection = List[DB.ElementId]([element.Id for element in elements])
				selection.SetElementIds(collection)


form = CreateWindow()
Application.Run(form)