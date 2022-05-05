# -*- coding: utf-8 -*-
"""
ViewSets

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Списки видов"
__doc__ = 'Создание списков видов для печати/экспорта\n' \


"""
KPLN

"""

import re
import webbrowser
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
from rpw.ui.forms import TextInput
import datetime

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "Редактор наборов"
		self.Text = "KPLN Редактор наборов"
		self.Size = Size(800, 608)
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
		self.button_create = Button(Text = "Создать")
		self.button_create.Click += self.create_set
		self.button_remove = Button(Text = "Удалить")
		self.button_remove.Click += self.remove_active
		self.button_save = Button(Text = "Сохранить")
		self.button_save.Click += self.OnApply
		self.button_close = Button(Text = "Закрыть")
		self.button_close.Click += self.OnClose
		self.button_help = Button(Text = "?")
		self.button_help.Click += self.go_to_help
		self.combobox_sets = ComboBox()
		self.combobox_c1 = ComboBox()
		self.combobox_c1.SelectedIndexChanged += self.CbOnChange
		self.combobox_c2 = ComboBox()
		self.combobox_c2.SelectedIndexChanged += self.CbOnChange
		self.combobox_c1_v = ComboBox()
		self.combobox_c2_v = ComboBox()
		self.combobox_c1_v.SelectedIndexChanged += self.CbOnChange_v
		self.combobox_c2_v.SelectedIndexChanged += self.CbOnChange_v
		self.listbox = ListView()
		self.quite = False
		self.names = []
		self.active_set = None
		self.groupbox = GroupBox()
		self.groupbox.Text = "Фильтры:"
		self.groupbox.Size = Size(347, 218)
		self.groupbox.Location = Point(425, 80)
		self.groupbox.Parent = self
		self.groupbox.ForeColor = Color.FromArgb(0,0,0,0)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = True
		self.TopMost = True
		self.MinimumSize = Size(800, 608)
		self.MaximumSize = Size(800, 608)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.CenterToScreen()
		self.sets = []
		self.collector_sheet = []
		self.collector_viewSets = []
		self.sheet_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets)
		self.sortlist_num = []
		self.sortlist_sheets = []
		for sc in self.sheet_collector:
			self.sortlist_sheets.append(sc)
			self.sortlist_num.append(sc.SheetNumber)
		self.sortlist_num.sort()
		for num in self.sortlist_num:
			for sheet in self.sortlist_sheets:
				if num == sheet.SheetNumber:
					self.collector_sheet.append(sheet)
					break
		self.viewSets_collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet)
		for vs in self.viewSets_collector:
			self.collector_viewSets.append(vs)
		self.temp = ""
		self.number = ""
		self.viewname = ""

		self.parameters = []
		self.parameters.append("<без фильтра>")
		self.parameters.append("Номер листа")
		self.parameters.append("Имя листа")
		self.parameters.append("Том")
		self.parameters.append("Дата утверждения листа")
		for b in self.collector_sheet:
			for j in b.Parameters:
				if j.IsShared:
					if str(j.Definition.Name) not in self.parameters:
						self.parameters.append(str(j.Definition.Name))
		self.parameters.sort()
		self.button_create.Parent = self
		self.button_create.Location = Point(425, 9)
		self.button_remove.Parent = self
		self.button_remove.Location = Point(505, 9)
		self.button_save.Parent = self
		self.button_save.Location = Point(425, 535)
		self.button_close.Parent = self
		self.button_close.Location = Point(505, 535)
		self.button_help.Parent = self
		self.button_help.Location = Point(585, 535)

		self.combobox_sets.Parent = self
		self.combobox_sets.Location = Point(425, 40)
		self.combobox_sets.Size = Size(346, 40)
		self.combobox_sets.DropDownWidth = 346
		self.combobox_sets.DropDownStyle = ComboBoxStyle.DropDownList
		#self.combobox_sets.SelectionChangeCommitted += self.OnChange
		self.combobox_sets.SelectedIndexChanged += self.OnChange
		self.cbx_active = False

		self.combobox_c1.Parent = self.groupbox
		self.combobox_c1.Location = Point(10, 50)
		self.combobox_c1.Size = Size(327, 40)
		self.combobox_c1.DropDownWidth = 327
		self.combobox_c1.DropDownStyle = ComboBoxStyle.DropDownList
		self.combobox_c1_v.Parent = self.groupbox
		self.combobox_c1_v.Location = Point(10, 80)
		self.combobox_c1_v.Size = Size(327, 40)
		self.combobox_c1_v.DropDownWidth = 327
		self.combobox_c1_v.DropDownStyle = ComboBoxStyle.DropDownList

		self.combobox_c2.Parent = self.groupbox
		self.combobox_c2.Location = Point(10, 140)
		self.combobox_c2.Size = Size(327, 40)
		self.combobox_c2.DropDownWidth = 327
		self.combobox_c2.DropDownStyle = ComboBoxStyle.DropDownList
		self.combobox_c2_v.Parent = self.groupbox
		self.combobox_c2_v.Location = Point(10, 170)
		self.combobox_c2_v.Size = Size(327, 40)
		self.combobox_c2_v.DropDownWidth = 327
		self.combobox_c2_v.DropDownStyle = ComboBoxStyle.DropDownList

		for p in self.parameters:
			self.combobox_c1.Items.Add(p)
			self.combobox_c2.Items.Add(p)
			self.combobox_c1.Text = "<без фильтра>"
			self.combobox_c2.Text = "<без фильтра>"

		self.get_sets()
		self.check_set_cativity()
		self.update_lists()

		self.c_cb = ColumnHeader()
		self.c_cb.Text = "Включить/№"
		self.c_cb.Width = -2
		self.c_cb.TextAlign = HorizontalAlignment.Center
		self.c_name = ColumnHeader()
		self.c_name.Text = "Наименование листа"
		self.c_name.Width = -2
		self.c_name.TextAlign = HorizontalAlignment.Left

		self.SuspendLayout()
		self.listbox.Dock = DockStyle.Fill
		self.listbox.View = View.Details

		self.listbox.Parent = self
		self.listbox.Size = Size(400, 550)
		self.listbox.Location = Point(10, 10)
		self.listbox.FullRowSelect = True
		self.listbox.GridLines = True
		self.listbox.AllowColumnReorder = True
		self.listbox.Sorting = SortOrder.Ascending
		self.listbox.Columns.Add(self.c_cb)
		self.listbox.Columns.Add(self.c_name)
		self.listbox.LabelEdit = False
		self.listbox.CheckBoxes = True
		self.listbox.MultiSelect = True
		self.cbx_active = True

	def logger(self, name, type, log = "null"):
		try:
			now = datetime.datetime.now()
			filename = "{}-{}-{}_{}-{}-{}_({})_SETMANAGER.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), revit.username)
			file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
			text = "unis report\nfile:{}\nversion:{}\nuser:{} : «{}» - {}\n{};".format(doc.PathName, revit.version, revit.username, name, type, log)
			file.write(text.encode('utf-8'))
			file.close()
		except: pass

	def go_to_help(self, sender, args):
		webbrowser.open('https://www.notion.so/kpln/09048c1546e9459aa6ef16444df72c00')

	def CbOnChange_v(self, sender, event):
		self.update_lists()

	def CbOnChange(self, sender, event):
		if self.combobox_c1.Text == "<без фильтра>":
			self.combobox_c2.Enabled = False
		else:
			self.combobox_c2.Enabled = True
		if sender == self.combobox_c1:
			self.values_1 = []
			self.temptext = ""
			for sheet in self.collector_sheet:
				try:
					self.temptext = sheet.LookupParameter(self.combobox_c1.Text).AsString()
					if self.temptext not in self.values_1:
						self.values_1.append(self.temptext)
				except: pass
			self.combobox_c1_v.Items.Clear()
			self.values_1.sort()

			if len(self.values_1) != 0 and self.combobox_c1.Enabled:
				self.combobox_c1_v.Enabled = True
			else: self.combobox_c1_v.Enabled = False

			for value in self.values_1:
				self.combobox_c1_v.Items.Add(str(value))
		if sender == self.combobox_c2:
			self.values_2 = []
			self.temptext = ""
			for sheet in self.collector_sheet:
				try:
					self.temptext = sheet.LookupParameter(self.combobox_c2.Text).AsString()
					if self.temptext not in self.values_2:
						self.values_2.append(self.temptext)
				except: pass
			self.combobox_c2_v.Items.Clear()
			self.values_2.sort()

			if len(self.values_2) != 0 and self.combobox_c2.Enabled:
				self.combobox_c2_v.Enabled = True
			else: self.combobox_c2_v.Enabled = False

			for value in self.values_2:
				try:
					self.combobox_c2_v.Items.Add(str(value))
				except: pass
		if self.combobox_c1.Enabled: self.combobox_c1_v.Enabled = True
		else: self.combobox_c1_v.Enabled = False
		if self.combobox_c2.Enabled: self.combobox_c2_v.Enabled = True
		else: self.combobox_c2_v.Enabled = False
		if self.combobox_c1.Text == "<без фильтра>":
			self.combobox_c2.Enabled = False
		else:
			self.combobox_c2.Enabled = True
		self.update_lists()

	def OnChange(self, sender, event):
		self.update_sets()
		self.update_lists()

	def OnClose(self, sender, args):
		self.Close()

	def OnApply(self, sender, args):
		log = "[apply]\n"
		log += "      - {}\n".format("self.update_lists()")
		self.apply_changes()
		self.logger(self.active_set.Name, "apply", log)


	def remove_active(self, sender, args):
		log = "[remove]\n"
		name = self.active_set.Name
		if self.active_set != None:
			with db.Transaction(name = "Remove active viewset"):
				try:
					doc.Delete(self.active_set.Id)
					log += "      - {}\n".format("doc.Delete(self.active_set.Id)")
				except: pass
			self.active_set = None			
			self.get_sets()
			log += "      - {}\n".format("self.get_sets()")
			self.update_lists()
			log += "      - {}\n".format("self.update_lists()")
			self.logger(name, "remove", log)

	def create_set(self, sender, args):	
		log = "[create]\n"
		self.names = []
		for v in self.collector_viewSets:
			self.names.append(str(v.Name))
			log += "for v in self.collector_viewSets : self.names.append({})\n".format(str(v.Name))
		value = str(TextInput('Создать новый набор', default="", description = "Введите название для нового набора:", exit_on_close = False))
		if value != "":
			if value not in self.names:
				with db.Transaction(name = "Create new viewset"):
					log += "{}\n".format("with db.Transaction(name = «Create new viewset»)")
					try:
						self.NewSet = DB.ViewSet()
						log += "{}\n".format("self.NewSet = DB.ViewSet()")
						for i in self.item:
							if i.Checked:
								log += "{}\n".format("?i.Checked")
								self.number = str(i.Text)
								si = i.SubItems
								self.viewname = str(si[1].Text)
								log += "search: {} : {}\n".format(self.number, self.viewname)
								for v in self.collector_sheet:
									if str(v.Category.Name) == "Листы":	
										if self.number == str(v.SheetNumber) and self.viewname == str(v.Name):
											#forms.alert(str(v.SheetNumber) + str(v.Name) + " - " + self.number + self.viewname)
											self.NewSet.Insert(v)
											log += "{} - appended\n".format(str(v.Id))
						self.printManager = doc.PrintManager
						log += "{}\n".format("self.printManager = doc.PrintManager")
						self.printManager.PrintRange = DB.PrintRange.Select
						log += "{}\n".format("self.printManager.PrintRange = DB.PrintRange.Select")
						self.viewSheetSetting = self.printManager.ViewSheetSetting
						log += "{}\n".format("self.viewSheetSetting = self.printManager.ViewSheetSetting")
						self.viewSheetSetting.CurrentViewSheetSet.Views = self.NewSet
						log += "{}\n".format("self.viewSheetSetting.CurrentViewSheetSet.Views = self.NewSet")
						self.viewSheetSetting.SaveAs(value)
						log += "{}\n".format("self.viewSheetSetting.SaveAs(value)")
					except: pass
				#self.combobox_sets.Items.Add(str(self.NewSet.Name))
				self.get_sets(name = value)
				log += "      - {}\n".format("self.get_sets(name = self.nvalue)")
				self.update_lists()
				log += "      - {}\n".format("self.update_lists()")
			else:
				log += "{}\n".format("?else")
				#self.Hide()
				forms.alert("Названия наборов не должны повторяться!")
				#self.Show()
		self.logger(value, "create", log)
		

	def apply_changes(self):
		log = "[apply]\n"
		if self.combobox_sets.Text != "<Все листы в проекте>" and self.active_set != None:
			self.nvalue = self.active_set.Name
			log += "self.nvalue = {}\n".format(self.active_set.Name)
			#with db.Transaction(name = "Remove active viewset"):
			#	doc.Delete(self.active_set.Id)
			#self.active_set = None
			for v in self.collector_viewSets:
				with db.Transaction(name = "Create new viewset"):
					log += "{}\n".format("with db.Transaction(name = «Create new viewset»)")
					try:
						log += "{}\n".format("with db.Transaction(name = «Create new viewset»)")
						self.NewSet = DB.ViewSet()
						log += "{}\n".format("self.NewSet = DB.ViewSet()")
						for i in self.item:
							if i.Checked:
								log += "{}\n".format("?i.Checked")
								self.number = str(i.Text)		
								si = i.SubItems
								self.viewname = str(si[1].Text)
								log += "search: {} : {}\n".format(self.number, self.viewname)
								for v in self.collector_sheet:
									if str(v.Category.Name) == "Листы":	
										if self.number == str(v.SheetNumber) and self.viewname == str(v.Name):
											self.NewSet.Insert(v)
											log += "{} - appended\n".format(str(v.Id))
					except: break
					try:
						self.printManager = doc.PrintManager
						log += "{}\n".format("self.printManager = doc.PrintManager")
						self.printManager.PrintRange = DB.PrintRange.Select
						log += "{}\n".format("self.printManager.PrintRange = DB.PrintRange.Select")
						self.viewSheetSetting = self.printManager.ViewSheetSetting
						log += "{}\n".format("self.viewSheetSetting = self.printManager.ViewSheetSetting")
						self.viewSheetSetting.CurrentViewSheetSet = self.active_set
						log += "{}\n".format("self.viewSheetSetting.CurrentViewSheetSet = self.active_set")
						self.viewSheetSetting.CurrentViewSheetSet.Views = self.NewSet
						log += "{}\n".format("self.viewSheetSetting.CurrentViewSheetSet.Views = self.NewSet")
						#self.active_set.Views = self.NewSet
						self.viewSheetSetting.Save()
						log += "{}\n".format("self.viewSheetSetting.Save()")
					except: pass
			self.get_sets(name = self.nvalue)
			log += "      - {}\n".format("self.get_sets(name = self.nvalue)")
			self.update_lists()
			log += "      - {}\n".format("self.update_lists()")

	def get_sets(self, name = "none"):
		self.collector_viewSets = []
		self.viewSets_collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet)
		for vs in self.viewSets_collector:
			self.collector_viewSets.append(vs)
		self.quite = True
		self.combobox_sets.Items.Clear()
		self.combobox_sets.Items.Add("<Все листы в проекте>")
		for set in self.collector_viewSets:
			if str(set.Name) > "":
				self.combobox_sets.Items.Add(str(set.Name))
				self.sets.append([str(set.Name), set])
				#self.active_set = set
		#self.combobox_sets.Text = str(self.active_set.Name)
		if name != "none":
			for set in self.collector_viewSets:
				if set.Name == name:
					self.active_set = set
					self.combobox_sets.Text = str(set.Name)
		else:
			self.active_set = None
			self.combobox_sets.Text = "<Все листы в проекте>"
		self.quite = False

	def update_sets(self):
		if not self.quite:
			for set in self.collector_viewSets:
				if str(set.Name) == self.combobox_sets.Text:
					self.active_set = set

	def check_sheet(self, sheet):
		if self.combobox_c1.Text == "<без фильтра>": return True
		try:
			if self.combobox_c1_v.Enabled and self.combobox_c1.Text != "<без фильтра>":
				if sheet.LookupParameter(self.combobox_c1.Text).AsString() != self.combobox_c1_v.Text: return False
			if self.combobox_c2_v.Enabled and self.combobox_c2.Text != "<без фильтра>":
				if sheet.LookupParameter(self.combobox_c2.Text).AsString() != self.combobox_c2_v.Text: return False
			return True
		except: return False

	def check_set_cativity(self):
		if len(self.sets) != 0:
			self.combobox_sets.Enabled = True
		else:
			self.combobox_sets.Enabled = False

	def update_lists(self):
		if not self.quite:
			self.listbox.Items.Clear()
			self.item = []
			if self.combobox_sets.Text == "<Все листы в проекте>":
				for sheet in self.collector_sheet:
					if str(sheet.Category.Name) == "Листы" and self.check_sheet(sheet):
						self.item.append(ListViewItem())
						self.item[len(self.item)-1].Text = sheet.SheetNumber
						self.item[len(self.item)-1].Checked = False
						self.item[len(self.item)-1].SubItems.Add(sheet.Name)
						self.listbox.Items.Add(self.item[len(self.item)-1])
			elif self.active_set != None:		
				for view in self.active_set.Views:
					if str(view.Category.Name) == "Листы" and self.check_sheet(view):
						self.item.append(ListViewItem())
						self.item[len(self.item)-1].Text = view.SheetNumber
						self.item[len(self.item)-1].Checked = True
						self.item[len(self.item)-1].SubItems.Add(view.Name)
						self.listbox.Items.Add(self.item[len(self.item)-1])

form = CreateWindow()

Application.Run(form)