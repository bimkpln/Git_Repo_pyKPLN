# -*- coding: utf-8 -*-
"""
RemoveUnplaced

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Удалить неразмещенные"
__doc__ = 'Удалить выбранные неиспользуемые виды' \

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

def logger(result):
	try:
		now = datetime.datetime.now()
		filename = "{}-{}-{}_{}-{}-{}_({})_CROPVIEWS.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), revit.username)
		file = open("Z:\\Отдел BIM\\Перфильев Игорь\\Отчеты_UNIs\\{}".format(filename), "w+")
		text = "unis report\nfile:{}\nversion:{}\nuser:{}\nresult:{};".format(doc.PathName, revit.version, revit.username, result)
		file.write(text.encode('utf-8'))
		file.close()
	except: pass

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "Подрезать виды"
		self.Text = "Выберите виды для подрезки"
		self.Size = Size(418, 608)
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
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

		self.selection = revit.get_selection()

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

		self.Text = "Выберите виды для удаления"
		self.plans = []
		self.plans_names = []
		self.views = []
		self.views_sorted = []
		self.collector_views = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
		self.item = []
		for view in self.collector_views:
			if view.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER).AsString() == None:
				self.views.append(view)
				self.views_sorted.append("{}_{}".format(view.Name, view.Id.ToString()))
		self.views_sorted.sort()
		self.views_sorted.reverse()
		for key in self.views_sorted:
			for view in self.views:
				if key == "{}_{}".format(view.Name, view.Id.ToString()):
					self.plans.append(view)
					self.plans_names.append(view.Name)
					self.item.append(ListViewItem())
					self.item[len(self.item)-1].Text = ""
					self.item[len(self.item)-1].Checked = False
					self.item[len(self.item)-1].SubItems.Add(view.Name)
					self.listbox.Items.Add(self.item[len(self.item)-1])	

	def OnOk(self, sender, args):
		with db.Transaction(name='Обрезка видов'):
			self.define_views()
		self.Close()
	def OnCancel(self, sender, args):
		self.Close()

	def define_views(self):
		for v in self.plans:
			for i in self.item:
				if i.Checked:
					si = i.SubItems
					try:
						if si[1].Text == v.Name:
							doc.Delete(v.Id)
							break
					except Exception as e:
						print(str(e))

form = CreateWindow()
Application.Run(form)
