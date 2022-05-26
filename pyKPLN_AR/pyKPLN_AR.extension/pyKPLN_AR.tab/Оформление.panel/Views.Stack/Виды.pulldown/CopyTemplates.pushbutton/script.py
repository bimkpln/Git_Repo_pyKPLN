# -*- coding: utf-8 -*-
"""
Copy_Template

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Копировать\nшаблоны"
__doc__ = 'Копирование шаблонов видов' \


"""
Архитекурное бюро KPLN

"""
import webbrowser
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
from rpw.ui.forms import Alert, CommandLink, TaskDialog

class CopyHandler(DB.IDuplicateTypeNamesHandler):
	def OnDuplicateTypeNamesFound(self, args):
		return DB.DuplicateTypeAction.UseDestinationTypes
class CreateWindow(Form):
	def __init__(self): 
		self.Name = "Копировать шаблоны"
		self.Text = "Копирование шаблонов"
		self.Size = Size(418, 608)
		self.Icon = Icon("X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico") 
		self.button_create = Button(Text = "Ок")
		self.ControlBox = True
		self.TopMost = True
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.MinimumSize = Size(418, 480)
		self.MaximumSize = Size(418, 480)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.CenterToScreen()
		self.views = ""

		self.lb1 = Label(Text = "Выберите проект с исходными шаблонами:")
		self.lb1.Parent = self
		self.lb1.Size = Size(343, 15)
		self.lb1.Location = Point(30, 80)

		self.lb2 = Label(Text = "Выбранные шаблоны из выбранного проекта будут\nскопированы в активный проект с заменой по наименованию.")
		self.lb2.Parent = self
		self.lb2.Size = Size(343, 40)
		self.lb2.Location = Point(30, 125)

		self.btn_confirm = Button(Text = "Далее")
		self.btn_confirm.Parent = self
		self.btn_confirm.Location = Point(30, 380)
		self.btn_confirm.Enabled = False
		self.btn_confirm.Click += self.OnNext

		self.active_doc = doc
		self.other_docs = self.get_docs()
		self.target_doc = None

		self.cb = ComboBox()
		self.cb.Parent = self
		self.cb.Size = Size(343, 100)
		self.cb.Location = Point(30, 100)
		self.cb.DropDownStyle = ComboBoxStyle.DropDownList
		self.cb.SelectedIndexChanged += self.allow_next
		self.cb.Items.Add("Выберите файл")
		self.cb.Text = "Выберите файл"
		for docel in self.other_docs:
			self.cb.Items.Add(docel.Title)

		self.listbox = ListView()
		self.c_cb = ColumnHeader()
		self.c_cb.Text = ""
		self.c_cb.Width = -2
		self.c_cb.TextAlign = HorizontalAlignment.Center
		self.c_name = ColumnHeader()
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

		self.button_cancel = Button(Text = "Закрыть")
		self.button_cancel.Parent = self
		self.button_cancel.Location = Point(100, 410)
		self.button_cancel.Click += self.OnCancel

		self.templates = []
		self.teplates_names = []

		self.listbox.Hide()
		self.button_ok.Hide()
		self.button_cancel.Hide()

	def get_checked_templates(self):
		list = []
		for i in self.item:
			if i.Checked:
				si = i.SubItems
				template_name = si[1].Text
				for t in self.templates:
					if template_name == t.Name:
						list.append(t)
		return list

	def OnOk(self, sender, args):
		view_templates = []
		view_templates = self.get_checked_templates()
		if len(view_templates) != 0:
			self.copy(self.active_doc, self.target_doc, view_templates)
		self.Close()

	def OnNext(self, sender, args):
		self.cb.Hide()
		self.lb1.Hide()
		self.lb2.Hide()
		self.btn_confirm.Hide()
		self.listbox.Show()
		self.button_ok.Show()
		self.button_cancel.Show()
		for i in self.get_docs():
			if i.Title == self.cb.Text:
				self.target_doc = i
		if self.target_doc != None:
			self.Text = "Выбор копируемых шаблонов"
			self.update(self.active_doc, self.target_doc)
		else:
			self.Hide()
			forms.alert("Не найден документ!")
			self.Show()
			self.Close()

	def allow_next(self, sender, event):
		if self.cb.Text != "Выберите файл":
			self.btn_confirm.Enabled = True
		else:
			self.btn_confirm.Enabled = False

	def update(self, active_doc, target_doc):
		self.view_templates = self.get_templates(target_doc)
		self.listbox.Clear()
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
		self.item = []
		self.templates = []
		self.teplates_names = []
		for template in self.view_templates:
			self.templates.append(template)
			self.teplates_names.append(template.Name)
			self.item.append(ListViewItem())
			self.item[len(self.item)-1].Text = ""
			self.item[len(self.item)-1].Checked = False
			self.item[len(self.item)-1].SubItems.Add(str(template.Name))
			self.listbox.Items.Add(self.item[len(self.item)-1])	

	def get_templates(self, document):
		list = []
		view_collector = DB.FilteredElementCollector(document).OfCategory(DB.BuiltInCategory.OST_Views).ToElements()
		for i in view_collector:
			if i.IsTemplate:
				list.append(i)
		return list

	def OnCancel(self, sender, args):
		self.Close()

	def get_docs(self):
		list = []
		for o_doc in revit.docs:
			if self.active_doc.PathName != o_doc.PathName:
				list.append(o_doc)
		return list

	def copy(self, active_doc, target_doc, elements):
		for element in elements:
			collection = List[DB.ElementId]()	
			collection.Add(element.Id)
			options = DB.CopyPasteOptions()
			options.SetDuplicateTypeNamesHandler(CopyHandler())
			with db.Transaction(name = "copy tempates"):
				self.Hide()
				try:
					DB.ElementTransformUtils.CopyElements(target_doc, collection, active_doc, DB.Transform.Identity, options)
					print("Шаблон «{}» успешно скопирован в проект!".format(str(element.Name)))
				except Exception as e:
					print("Шаблон «{}» невозможно копировать в проект! Операция отменена! {}".format(str(element.Name), str(e)))
				self.Show()

form = CreateWindow()
Application.Run(form)

