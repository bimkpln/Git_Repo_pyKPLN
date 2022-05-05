# -*- coding: utf-8 -*-
"""
Create_broach

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Создать развертки по стенам"
__doc__ = 'Создание развертки по стенам на основе выбранного помещения' \


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
from libKPLN_CreationUtils.Library import *

class RoomSelectorFilter (UI.Selection.ISelectionFilter):
	def AllowElement(self, element = DB.Element):
		if element.Category.Id == DB.ElementId(-2000160):
			return True
		return False
	def AllowReference(self, refer, point):
		return False

class Picker(Form):
	def __init__(self):
		global view_template
		global sheet_template
		self.Name = "KPLN_AR_FinishViews"
		self.Text = "KPLN Оформление : Параметры"
		self.Size = Size(400, 140)
		self.MinimumSize = Size(400, 140)
		self.MaximumSize = Size(400, 140)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = True
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.TopMost = True
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")

		self.btn_apply = Button(Text = "Применить")
		self.btn_apply.Location = Point(10, 70)
		self.btn_apply.Parent = self
		self.btn_apply.Click += self.Apply
		self.btn_apply.Enabled = False

		self.pick_sheet = ComboBox()
		self.pick_sheet.Location = Point(10, 10)
		self.pick_sheet.Parent = self
		self.pick_sheet.Size = Size(363, 20)
		self.pick_sheet.DropDownStyle = ComboBoxStyle.DropDownList
		self.pick_sheet.Items.Add("<выберите формат листа из списка>")
		self.pick_sheet.Text = "<выберите формат листа из списка>"
		for view_sheet_type in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType():
			self.pick_sheet.Items.Add(view_sheet_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString())

		self.pick_template = ComboBox()
		self.pick_template.Location = Point(10, 40)
		self.pick_template.Parent = self
		self.pick_template.Size = Size(363, 20)
		self.pick_template.DropDownStyle = ComboBoxStyle.DropDownList

		self.pick_template.Items.Add("<выберите шаблон вида из списка>")
		self.pick_template.Text = "<выберите шаблон вида из списка>"
		for view_type in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views):
			if str(view_type.ViewType) == "Section" and view_type.IsTemplate:
				self.pick_template.Items.Add(str(view_type.Title))

		self.pick_sheet.SelectedIndexChanged += self.CblOnChanged
		self.pick_template.SelectedIndexChanged += self.CblOnChanged
		self.CenterToScreen()

	def CblOnChanged(self, sender, event):
		global view_template
		global sheet_template
		if self.pick_template.Text != "<выберите шаблон вида из списка>" and self.pick_sheet.Text != "<выберите формат листа из списка>":
			self.btn_apply.Enabled = True
			for view_type in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views):
				if str(view_type.ViewType) == "Section" and str(view_type.Title) == self.pick_template.Text:
					view_template = view_type
			for view_sheet_type in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType():
				if view_sheet_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == self.pick_sheet.Text:
					sheet_template = view_sheet_type
		else:
			self.btn_apply.Enabled = False

	def Apply(self, sender, args):
		global script_scipped
		script_scipped = False
		self.Close()

rooms = []
view_template = None
sheet_template = None
script_scipped = True

with db.Transaction(name = "Load families"):
	try:
		doc.LoadFamily("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_LoadUtils\\Families\\SHEETS\\SYS_Автоматическое заполнение_Ось.rfa")
	except: pass

form = Picker()
Application.Run(form)

try:
	filter_openings = RoomSelectorFilter()
	element = uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element, filter_openings, "KPLN : Веберите помещения")
	rooms.append(KPLN_Room(doc.GetElement(element), view_template, sheet_template, KPLN_IdGenerator()))
except: pass

if len(rooms) != 0:
	if not script_scipped:
		for room in rooms:
			room.CreateSectionViews()
			room.CreateSectionSheet()
			room.PlaceSectionViews()

