# -*- coding: utf-8 -*-
"""
CreateSchedules

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Создать спецификации"
__doc__ = 'Создаются для каждого уровня с заменой предыдущих ведомостей\n' \

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

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "Создание спецификаций"
		self.Text = "Выберите уровни"
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

		self.checkbox = CheckBox()
		self.checkbox_gost = CheckBox()
		self.SuspendLayout()
		self.listbox.Dock = DockStyle.Fill
		self.listbox.View = View.Details

		self.listbox.Parent = self
		self.listbox.Size = Size(400, 360)
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

		self.checkbox.Parent = self
		self.checkbox.Location = Point(12, 365)
		self.checkbox.Size = Size(350, 20)
		self.checkbox.Text = "Создать стандартную спецификацию"
		self.checkbox.Checked = False

		self.checkbox_gost.Parent = self
		self.checkbox_gost.Location = Point(12, 385)
		self.checkbox_gost.Size = Size(350, 20)
		self.checkbox_gost.Text = "Создать спецификацию по ГОСТ"
		self.checkbox_gost.Checked = True

		self.button_ok = Button(Text = "Ok")
		self.button_ok.Parent = self
		self.button_ok.Location = Point(10, 410)
		self.button_ok.Click += self.OnOk

		self.button_ok = Button(Text = "Закрыть")
		self.button_ok.Parent = self
		self.button_ok.Location = Point(100, 410)
		self.button_ok.Click += self.OnCancel

		self.Text = "Создать ведомости отделки по уровням"
		self.levels = []
		self.levels_sorted = []
		self.levels_sort = []
		self.levels_sort_elevation = []
		self.levels_names = []
		self.collector_levels = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
		self.item = []
		for level in self.collector_levels:
			self.levels_sort.append(level)
			self.levels_sort_elevation.append(level.Elevation)
		self.levels_sort_elevation.sort()
		self.levels_sort_elevation.reverse()
		for i in self.levels_sort_elevation:
			for level in self.levels_sort:
				if i == level.Elevation:
					self.levels_sorted.append(level)
					break
		for level in self.levels_sorted:
			try:
				self.levels.append(level)
				name = level.Name
				self.levels_names.append(name)
				self.item.append(ListViewItem())
				self.item[len(self.item)-1].Text = ""
				self.item[len(self.item)-1].Checked = False
				self.item[len(self.item)-1].SubItems.Add(name)
				self.listbox.Items.Add(self.item[len(self.item)-1])
			except: pass

	def OnOk(self, sender, args):
		self.define_levels()
		self.Close()

	def define_levels(self):
		global picked_levels
		global checked_regular
		global checked_gost
		for i in self.item:
			if i.Checked:
				si = i.SubItems
				viewname = si[1].Text
				for level in self.levels:
					if viewname == level.Name:
						picked_levels.append(level)
		checked_regular = self.checkbox.Checked
		checked_gost = self.checkbox_gost.Checked

	def OnCancel(self, sender, args):
		self.Close()

def GetParamId(Name):
	room = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement()
	for j in room.Parameters:
		try:
			if j.Definition.Name == Name:
				return j.Id
		except: pass

def NameIsUniq(OST, name, remove):
	elements = DB.FilteredElementCollector(doc).OfCategory(OST).ToElements()
	for element in elements:
		try:
			if element.Name == name:
				if remove:
					if doc.ActiveView.Id.ToString() == element.Id.ToString():
						forms.alert("Невозможно заменить открытую ведомость «{}»!".format(element.Name))
						return False
					else:
						doc.Delete(element.Id)
						doc.Regenerate()
						return True
				doc.Regenerate()
				return False
		except: pass
	doc.Regenerate()
	return True

def CreateSchedule_FinishingGOST(level = None):
	with db.Transaction(name = "Create ViewSchedule"):
		schedule = DB.ViewSchedule.CreateSchedule(doc, DB.ElementId(DB.BuiltInCategory.OST_Rooms))
		doc.Regenerate()
		field_1 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Группа"))
		field_2 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_ГОСТ_Описание стен"))
		field_3 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_ГОСТ_Площадь стен_Текст"))
		field_4 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_ГОСТ_Описание потолков"))
		field_5 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_ГОСТ_Площадь потолков_Текст"))
		field_6 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_ГОСТ_Описание полов"))
		field_7 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_ГОСТ_Площадь полов_Текст"))
		field_8 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_ГОСТ_Описание плинтусов"))
		field_9 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_ГОСТ_Длина плинтусов_Текст"))
		field_10 = DB.SchedulableField(DB.ScheduleFieldType.Instance, DB.ElementId(DB.BuiltInParameter.ROOM_LEVEL_ID))
		schedule.Definition.AddField(field_1)
		schedule.Definition.AddField(field_2)
		schedule.Definition.AddField(field_3)
		schedule.Definition.AddField(field_4)
		schedule.Definition.AddField(field_5)
		schedule.Definition.AddField(field_6)
		schedule.Definition.AddField(field_7)
		schedule.Definition.AddField(field_8)
		schedule.Definition.AddField(field_9)
		schedule.Definition.AddField(field_10)
		#MAIN
		schedule.Definition.ShowHeaders = True
		schedule.Definition.ShowTitle = True
		#1
		sfield_1 = schedule.Definition.GetField(0)
		sfield_1.ColumnHeading = "Помещения"
		sfield_1.GridColumnWidth = 30 / 304.8
		sfield_1.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_1.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#2
		sfield_2 = schedule.Definition.GetField(1)
		sfield_2.ColumnHeading = "Описание"
		sfield_2.GridColumnWidth = 100 / 304.8
		sfield_2.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_2.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#3
		sfield_3 = schedule.Definition.GetField(2)
		sfield_3.ColumnHeading = "Площадь"
		sfield_3.GridColumnWidth = 20 / 304.8
		sfield_3.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_3.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#4
		sfield_4 = schedule.Definition.GetField(3)
		sfield_4.ColumnHeading = "Описание"
		sfield_4.GridColumnWidth = 100 / 304.8
		sfield_4.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_4.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#5
		sfield_5 = schedule.Definition.GetField(4)
		sfield_5.ColumnHeading = "Площадь"
		sfield_5.GridColumnWidth = 20 / 304.8
		sfield_5.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_5.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#6
		sfield_6 = schedule.Definition.GetField(5)
		sfield_6.ColumnHeading = "Описание"
		sfield_6.GridColumnWidth = 100 / 304.8
		sfield_6.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_6.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#7
		sfield_7 = schedule.Definition.GetField(6)
		sfield_7.ColumnHeading = "Площадь"
		sfield_7.GridColumnWidth = 20 / 304.8
		sfield_7.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_7.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#8
		sfield_8 = schedule.Definition.GetField(7)
		sfield_8.ColumnHeading = "Описание"
		sfield_8.GridColumnWidth = 100 / 304.8
		sfield_8.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_8.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#9
		sfield_9 = schedule.Definition.GetField(8)
		sfield_9.ColumnHeading = "Площадь"
		sfield_9.GridColumnWidth = 20 / 304.8
		sfield_9.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_9.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#6
		sfield_10 = schedule.Definition.GetField(9)
		sfield_10.IsHidden = True
		#SORT
		sortfield = DB.ScheduleSortGroupField(DB.ScheduleFieldId(schedule.Definition.GetField(0).FieldId.IntegerValue))
		schedule.Definition.InsertSortGroupField(sortfield, 0)
		schedule.Definition.IsItemized = False
		#FILTER
		if level != None:
			schedule_filter_1 = DB.ScheduleFilter(sfield_10.FieldId, DB.ScheduleFilterType.Equal, level.Id)
			schedule.Definition.AddFilter(schedule_filter_1)
		schedule_filter_2 = DB.ScheduleFilter(sfield_1.FieldId, DB.ScheduleFilterType.GreaterThan, "")
		schedule.Definition.AddFilter(schedule_filter_2)
		#GROUP HEADER
		schedule.GroupHeaders(0, 1, 0, 2, "Стены")
		schedule.GroupHeaders(0, 3, 0, 4, "Потолки")
		schedule.GroupHeaders(0, 5, 0, 6, "Полы")
		schedule.GroupHeaders(0, 7, 0, 8, "Плинтусы")
		#NAME
		if level == None:
			fish = "О_ГОСТ_Ведомость отделки помещений"
		else:
			fish = "О_ГОСТ_Ведомость отделки помещений ({})".format(str(level.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString()))
		if NameIsUniq(DB.BuiltInCategory.OST_Schedules, fish, True):
			try:
				schedule.Name = fish
			except:
				stop = False
				n = 1
				while not stop:
					if n > 100:
						stop = True
					schedule_name = "{}_{}".format(fish, str(n))
					if NameIsUniq(DB.BuiltInCategory.OST_Schedules, schedule_name, True):
						schedule.Name = schedule_name
						stop = True
					n+=1
		else:
			stop = False
			n = 1
			while not stop:
				if n > 100:
					stop = True
				schedule_name = "{}_{}".format(fish, str(n))
				if NameIsUniq(DB.BuiltInCategory.OST_Schedules, schedule_name, True):
					schedule.Name = schedule_name
					stop = True
				n+=1

def CreateSchedule_FinishingDefault(level = None):
	with db.Transaction(name = "Create ViewSchedule"):
		schedule = DB.ViewSchedule.CreateSchedule(doc, DB.ElementId(DB.BuiltInCategory.OST_Rooms))
		doc.Regenerate()
		field_1 = DB.SchedulableField(DB.ScheduleFieldType.Instance, DB.ElementId(DB.BuiltInParameter.ROOM_NAME))
		field_2 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Описание стен"))
		field_3 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Площадь стен_Текст"))
		field_4 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Описание потолков"))
		field_5 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Площадь потолков_Текст"))
		field_6 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Описание полов"))
		field_7 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Площадь полов_Текст"))
		field_8 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Описание плинтусов"))
		field_9 = DB.SchedulableField(DB.ScheduleFieldType.Instance, GetParamId("О_ПОМ_Длина плинтусов_Текст"))
		field_10 = DB.SchedulableField(DB.ScheduleFieldType.Instance, DB.ElementId(DB.BuiltInParameter.ROOM_LEVEL_ID))
		schedule.Definition.AddField(field_1)
		schedule.Definition.AddField(field_2)
		schedule.Definition.AddField(field_3)
		schedule.Definition.AddField(field_4)
		schedule.Definition.AddField(field_5)
		schedule.Definition.AddField(field_6)
		schedule.Definition.AddField(field_7)
		schedule.Definition.AddField(field_8)
		schedule.Definition.AddField(field_9)
		schedule.Definition.AddField(field_10)
		#MAIN
		schedule.Definition.ShowHeaders = True
		schedule.Definition.ShowTitle = True
		#1
		sfield_1 = schedule.Definition.GetField(0)
		sfield_1.ColumnHeading = "Помещения"
		sfield_1.GridColumnWidth = 30 / 304.8
		sfield_1.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_1.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#2
		sfield_2 = schedule.Definition.GetField(1)
		sfield_2.ColumnHeading = "Описание"
		sfield_2.GridColumnWidth = 100 / 304.8
		sfield_2.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_2.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#3
		sfield_3 = schedule.Definition.GetField(2)
		sfield_3.ColumnHeading = "Площадь"
		sfield_3.GridColumnWidth = 20 / 304.8
		sfield_3.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_3.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#4
		sfield_4 = schedule.Definition.GetField(3)
		sfield_4.ColumnHeading = "Описание"
		sfield_4.GridColumnWidth = 100 / 304.8
		sfield_4.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_4.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#5
		sfield_5 = schedule.Definition.GetField(4)
		sfield_5.ColumnHeading = "Площадь"
		sfield_5.GridColumnWidth = 20 / 304.8
		sfield_5.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_5.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#6
		sfield_6 = schedule.Definition.GetField(5)
		sfield_6.ColumnHeading = "Описание"
		sfield_6.GridColumnWidth = 100 / 304.8
		sfield_6.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_6.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#7
		sfield_7 = schedule.Definition.GetField(6)
		sfield_7.ColumnHeading = "Площадь"
		sfield_7.GridColumnWidth = 20 / 304.8
		sfield_7.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_7.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#8
		sfield_8 = schedule.Definition.GetField(7)
		sfield_8.ColumnHeading = "Описание"
		sfield_8.GridColumnWidth = 100 / 304.8
		sfield_8.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_8.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#9
		sfield_9 = schedule.Definition.GetField(8)
		sfield_9.ColumnHeading = "Площадь"
		sfield_9.GridColumnWidth = 20 / 304.8
		sfield_9.HorizontalAlignment = DB.ScheduleHorizontalAlignment.Left
		sfield_9.HeadingOrientation = DB.ScheduleHeadingOrientation.Horizontal
		#10
		sfield_10 = schedule.Definition.GetField(9)
		sfield_10.IsHidden = True
		#SORT
		sortfield = DB.ScheduleSortGroupField(DB.ScheduleFieldId(schedule.Definition.GetField(0).FieldId.IntegerValue))
		schedule.Definition.InsertSortGroupField(sortfield, 0)
		schedule.Definition.IsItemized = True
		#FILTER
		if level != None:
			schedule_filter = DB.ScheduleFilter(sfield_10.FieldId, DB.ScheduleFilterType.Equal, level.Id)
			schedule.Definition.AddFilter(schedule_filter)
		#GROUP HEADER
		schedule.GroupHeaders(0, 1, 0, 2, "Стены")
		schedule.GroupHeaders(0, 3, 0, 4, "Потолки")
		schedule.GroupHeaders(0, 5, 0, 6, "Полы")
		schedule.GroupHeaders(0, 7, 0, 8, "Плинтусы")
		#NAME
		if level == None:
			fish = "О_Ведомость отделки помещений"
		else:
			fish = "О_Ведомость отделки помещений ({})".format(str(level.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString()))
		if NameIsUniq(DB.BuiltInCategory.OST_Schedules, fish, True):
			try:
				schedule.Name = fish
			except:
				stop = False
				n = 1
				while not stop:
					if n > 100:
						stop = True
					schedule_name = "{}_{}".format(fish, str(n))
					if NameIsUniq(DB.BuiltInCategory.OST_Schedules, schedule_name, True):
						schedule.Name = schedule_name
						stop = True
					n+=1
		else:
			stop = False
			n = 1
			while not stop:
				if n > 100:
					stop = True
				schedule_name = "{}_{}".format(fish, str(n))
				if NameIsUniq(DB.BuiltInCategory.OST_Schedules, schedule_name, True):
					schedule.Name = schedule_name
					stop = True
				n+=1

#CreateSchedule_FinishingGOST()
#CreateSchedule_FinishingDefault()
picked_levels = []
checked_regular = False
checked_gost = False
form = CreateWindow()
Application.Run(form)
for level in picked_levels:
	try:
		if checked_regular:
			CreateSchedule_FinishingDefault(level)
		if checked_gost:
			CreateSchedule_FinishingGOST(level)
	except Exception as e:
		forms.alert("Ошибка! Необходимо проверить наличие параметров!\n{}".format(str(e)))
		break