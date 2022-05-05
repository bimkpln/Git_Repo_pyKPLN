# -*- coding: utf-8 -*-
"""
SelectByMaterial

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Выбрать по материалу"
__doc__ = 'Выбирает все элементы с выбранными материалами' \

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
		self.Name = "Выбор по материалу"
		self.Text = "Выбор по материалу"
		self.Size = Size(418, 608)
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
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
		self.c_name.Text = "Имя"
		self.c_name.Text = "Имя"
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

		self.Text = "Выберите материалы для выбора"
		self.materials = []
		self.materials_names = []
		self.materials_sort = []
		self.materials_sorted = []
		self.materials_sort_name = []
		self.item = []		
		self.collector_all_elements = DB.FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
		for element in self.collector_all_elements:
			for material_id in element.GetMaterialIds(False):
				if not self.MtlInList(self.materials_sort, doc.GetElement(material_id)):
					self.materials_sort.append(doc.GetElement(material_id))
					self.materials_sort_name.append(doc.GetElement(material_id).Name)
		self.materials_sort_name.sort()
		self.materials_sort_name.reverse()
		for i in self.materials_sort_name:
			for material in self.materials_sort:
				if i == material.Name:
					self.materials_sorted.append(material)
					break
		for material in self.materials_sorted:
			try:
				self.materials.append(material)
				name = material.Name
				self.materials_names.append(name)
				self.item.append(ListViewItem())
				self.item[len(self.item)-1].Text = ""
				self.item[len(self.item)-1].Checked = False
				self.item[len(self.item)-1].SubItems.Add(name)
				self.listbox.Items.Add(self.item[len(self.item)-1])
			except: pass

	def MtlInList(self, list, material):
		for i in list:
			if str(i.Id) == str(material.Id):
				return True
		return False

	def OnOk(self, sender, args):
		self.select_mat_elements(self.define_materials())
		self.Close()

	def Has_Material(self, element, material):
		try:
			materials = element.GetMaterialIds(False)
			for i in materials:
				for z in material:
					if str(z.Id) == str(i):
						return True
			return False
		except: return False

	def define_materials(self):
		list_of_materials = []
		for i in self.item:
			if i.Checked:
				si = i.SubItems
				materialname = si[1].Text
				for material in self.materials:
					if materialname == material.Name:
						list_of_materials.append(material)
		return list_of_materials

	def OnCancel(self, sender, args):
		self.Close()

	def select_mat_elements(self, materials):
		if len(materials) != 0:
			elements = []
			active_view = doc.ActiveView.Id
			collector = DB.FilteredElementCollector(doc, active_view).WhereElementIsNotElementType().ToElements()
			for element in collector:
				if self.Has_Material(element, materials):
					elements.append(element)
			if len(elements) != 0:
				selection = uidoc.Selection
				collection = List[DB.ElementId]([element.Id for element in elements])
				selection.SetElementIds(collection)


form = CreateWindow()
Application.Run(form)