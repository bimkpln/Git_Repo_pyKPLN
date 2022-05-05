# -*- coding: utf-8 -*-
"""
PrintPDF

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Печать PDF"
__doc__ = 'Пакетная печать PDF стандартных форматов\n' \
          'Принтер автоматически назначается согласно приоритету. Для работы необходим как минимум один из списка ниже;\n\n' \
          'Используемые принтеры по приоритету:\n' \
          '   «PDF24 PDF» - рекомендуемый (если отсутствует в списке создайте заявку в ELMA на установку принтера «PDF 24»)\n' \
          '   «Adobe PDF»\n' \
          '   «PDFCreator»' \


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
from rpw.ui.forms import TextInput, Alert, select_folder

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "KPLN_AR_PDF"
		self.Text = "KPLN Печать PDF"
		self.Size = Size(800, 600)
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
		self.button = []
		self.combobox = []
		self.textbox = []
		self.green = []
		self.label = []
		self.groupbox = GroupBox()
		self.groupbox.Text = "Фильтры:"
		self.groupbox.Size = Size(347, 218)
		self.groupbox.Location = Point(425, 9)
		self.groupbox.Parent = self
		self.groupbox.ForeColor = Color.FromArgb(0,0,0,0)
		self.ControlBox = False
		self.TopMost = True
		self.MinimumSize = Size(800, 600)
		self.MaximumSize = Size(800, 600)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.CenterToScreen()
		self.sheet_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType()
		self.title_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType()


		self.PO_1 = DB.PageOrientationType.Portrait
		self.PO_2 = DB.PageOrientationType.Landscape
		self.Formats = []
		self.Formats.append(["A4",297,210, self.PO_1])
		self.Formats.append(["A4",210,297, self.PO_2])
		self.Formats.append(["A3",420,297, self.PO_1])
		self.Formats.append(["A3",297,420, self.PO_2])
		self.Formats.append(["A2",594,420, self.PO_1])
		self.Formats.append(["A2",420,594, self.PO_2])
		self.Formats.append(["A1",841,594, self.PO_1])
		self.Formats.append(["A1",594,841, self.PO_2])
		self.Formats.append(["A0",1189,841, self.PO_1])
		self.Formats.append(["A0",841,1189, self.PO_2])

		self.temp = []
		self.printable = []
		self.presets = []
		self.log = ""
		self.temptext = ""
		self.label_down = []

		self.listbox = ListBox()
		self.listbox.Parent = self
		self.listbox.Size = Size(400, 500)
		self.listbox.Location = Point(15, 15)
		self.listbox.SelectionMode = SelectionMode.MultiExtended
		self.listbox.Sorted = True
		self.logit = ""

		self.parameters = []
		self.values = []
		self.printer_dictionary = ["PDFCreator", "Adobe PDF", "PDF24 PDF", "PDF24 FAX"]
		self.printers = []

		self.parameters.append("Номер листа")
		self.parameters.append("Имя листа")
		self.parameters.append("Том")
		for b in self.sheet_collector:
			for j in b.Parameters:
				if j.IsShared and j.HasValue:
					if str(j.Definition.Name) not in self.parameters:
						self.parameters.append(str(j.Definition.Name))
		self.parameters.sort()

		for sheet in self.sheet_collector:
			for title in self.title_collector:
				if title.LookupParameter("Имя листа").AsString() == sheet.Name:
					self.temp = []
					self.temp.append(sheet.Name)
					self.temp.append(sheet.LookupParameter("Номер листа").AsString())
					self.temp.append(sheet)
					self.temp.append(title)
					self.temp.append(title.LookupParameter("Ширина листа").AsDouble() * 304.8)
					self.temp.append(title.LookupParameter("Высота листа").AsDouble() * 304.8)
					self.presets.append(self.temp)

		for i in range (0,2):
			self.label.append(Label(Text = "Условие " + str(i + 1)))
			self.label[i].Parent = self.groupbox
			self.label[i].Location = Point(10, 25 + 100 * i)
			self.combobox.append(ComboBox())
			self.combobox[i].Location = Point(10, 55 + 100 * i)
			self.combobox[i].Parent = self.groupbox
			self.combobox[i].Size = Size(327, 40)
			self.textbox.append(ComboBox())
			self.textbox[i].Location = Point(10, 85 + 100 * i)
			self.textbox[i].Parent = self.groupbox
			self.textbox[i].Size = Size(327, 40)
			self.combobox[i].DropDownWidth = 327
			self.textbox[i].DropDownWidth = 327
			self.combobox[i].DropDownStyle = ComboBoxStyle.DropDownList
			self.textbox[i].DropDownStyle = ComboBoxStyle.DropDownList
			self.combobox[i].Enabled = True
			self.textbox[i].Enabled = False
			self.combobox[i].Items.Add("")
			self.combobox[i].Items.Add(".....................................................................................................")
			for z in self.parameters:
				self.combobox[i].Items.Add(z)
			self.combobox[i].SelectionChangeCommitted += self.OnChange
			self.combobox[i].SelectedIndexChanged += self.OnChange
			self.textbox[i].SelectionChangeCommitted += self.OnChange
			self.textbox[i].SelectedIndexChanged += self.OnChange
			self.textbox[i].SelectionChangeCommitted += self.OnChange2
			self.textbox[i].SelectedIndexChanged += self.OnChange2
		self.combobox[1].Enabled = False

		self.button.append(Button())
		self.button[0].Parent = self
		self.button[0].Text = "Печать"
		self.button[0].Location = Point(15, 526)
		self.button[0].Click += self.OnRun

		self.button.append(Button())
		self.button[1].Parent = self
		self.button[1].Text = "Отмена"
		self.button[1].Location = Point(100, 526)
		self.button[1].Click += self.OnClose

		self.button.append(Button())
		self.button[2].Parent = self
		self.button[2].Text = "?"
		self.button[2].Location = Point(185, 526)
		self.button[2].Click += self.go_to_help

		self.label_down.append(Label(Text = "Формат наименования:"))
		self.label_down[0].Parent = self
		self.label_down[0].Location = Point(425, 290)
		self.label_down[0].Size = Size(303, 15)

		self.label_down.append(Label(Text = "Примечание: листы создаются отдельно.\n\nНе рекомендуется использовать совместно с приложением\n«PDF24 - PDF Конструктор».\n\nДля печати подходят только стандартные форматы листов\n(A0, A1, A2, A3, A4, A5, A6) в альбомной и портретной раскладки.\nЕсли элементы выйдут за границы листа то размер листа\nможет быть нераспознан!\n\n[!] - листы с нестандартым масштабом"))
		self.label_down[1].Parent = self
		self.label_down[1].Location = Point(425, 340)
		self.label_down[1].AutoSize = True

		self.label_down.append(Label(Text = "Выберите принтер:"))
		self.label_down[2].Parent = self
		self.label_down[2].Location = Point(425, 240)
		self.label_down[2].Size = Size(303, 15)
		self.alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZабвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ1234567890 «»=№-+.,()_[]"

		self.nametype = ComboBox()
		self.nametype.Parent = self
		self.nametype.Location = Point(435, 310)
		self.nametype.Size = Size(327, 40)
		self.nametype.DropDownWidth = 327
		self.nametype.DropDownStyle = ComboBoxStyle.DropDownList
		self.nametype.Items.Add("[номер листа] + [имя листа].pdf")
		self.nametype.Items.Add("[номер листа].pdf")
		self.nametype.Items.Add("[имя].pdf")
		self.nametype.Items.Add("[год] - [месяц] - [день] + [имя листа].pdf")
		self.nametype.Items.Add("[год] - [месяц] - [день] + [номер листа].pdf")
		self.nametype.Items.Add("[имя] + [год] - [месяц] - [день].pdf")
		self.nametype.Items.Add("[номер листа] + [год] - [месяц] - [день].pdf")
		self.nametype.Items.Add("[номер листа] + [имя листа] + [год] - [месяц] - [день].pdf")
		self.nametype.Items.Add("[год] - [месяц] - [день] + [номер листа] + [имя листа].pdf")
		self.nametype.Text = "[имя].pdf"
		self.path_default = "C:\\"
		self.ticker = 0
		self.ticker_max = 0	
		self.logsheets = 0
		for preset in self.presets:
			for f in range(len(self.Formats)):
				self.FormatH = round(self.Formats[f][1], 1)
				self.FormatW = round(self.Formats[f][2], 1)
				if round(preset[5],1) == self.FormatH and round(preset[4],1) == self.FormatW:
					self.green.append(preset[1] + " - " + preset[0])

		self.printManager = doc.PrintManager
		for printer in self.printer_dictionary:
			try:
				self.printManager.SelectNewPrintDriver(printer)
				self.printers.append(printer)
			except: pass
		self.printselect = ComboBox()
		self.printselect.Parent = self
		self.printselect.Location = Point(435, 260)
		self.printselect.Size = Size(327, 40)
		self.printselect.DropDownWidth = 327
		self.printselect.DropDownStyle = ComboBoxStyle.DropDownList
		if len(self.printers) == 0:
			self.Close()
		for i in self.printers:
			self.printselect.Items.Add(i)
		self.printselect.Text = self.printers[0]

		self.call_update()

	def logger(self, number):
		try:
			now = datetime.datetime.now()
			filename = "{}-{}-{}_{}-{}-{}_{}_ПЕЧАТЬPDF.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), revit.username)
			file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
			text = "unis report\nfile:{}\nversion:{}\nuser:{}, has printed {} sheets;".format(doc.PathName, revit.version, revit.username, str(number))
			file.write(text.encode('utf-8'))
			file.close()
		except: pass

	def go_to_help(self, sender, args):
		webbrowser.open('https://www.notion.so/kpln/PDF-185008efa91045329274caa8a81dea7d')

	def Get_Path(self):
		self.dialog = FolderBrowserDialog()
		if (self.dialog.ShowDialog(self) == DialogResult.OK):
			self.path_default = self.dialog.SelectedPath  

	def Browse(self, sender, event):
		self.constructor = ""
		self.directory.Text = ""
		dialog = FolderBrowserDialog()
		dialog.Description = "Выберите папку для сохранения листов в формате PDF"
		if (dialog.ShowDialog(self) == DialogResult.OK):
			self.directory.Text = dialog.SelectedPath

	def OnRun(self, sender, event):
		self.Text = "Подготовка к печати..."
		self.button[0].Enabled = False
		self.run_script()
		self.button[0].Enabled = False
		self.logger(self.logsheets)
		self.Close()

	def OnClose(self, sender, event):
		self.Close()

	def OnChange(self, sender, event):
		if sender.Text == ".....................................................................................................":
			sender.Text = ""
		if sender == self.combobox[0]:
			self.call_update_activity()
		elif sender == self.combobox[1]:
			self.call_update_activity_2()
		else: self.call_update_activity_3()

	def OnChange2(self, sender, event):
		self.call_update()

	def call_update_activity_3(self):
		if self.combobox[0].Enabled == True and self.combobox[0].Text in self.parameters:
			self.textbox[0].Enabled = True
		else: self.textbox[0].Enabled = False

		if self.textbox[0].Enabled == True and self.textbox[0].Text != "":
			self.combobox[1].Enabled = True
		else:
			self.combobox[1].Enabled = False
			self.combobox[1].Text = ""
			self.textbox[1].Items.Clear()
			self.textbox[1].Enabled = False
		if self.combobox[1].Enabled == True and self.combobox[1].Text in self.parameters:
			self.textbox[1].Enabled = True
		else:
			self.textbox[1].Enabled = False
		self.call_update()

	def call_update_activity_2(self):
		self.values = []
		for sheet in self.sheet_collector:
			try:
				self.temptext = sheet.LookupParameter(self.combobox[1].Text).AsString()
				if self.temptext not in self.values:
					self.values.append(self.temptext)
			except: pass
		self.textbox[1].Items.Clear()
		for value in self.values:
			try:
				self.textbox[1].Items.Add(str(value))
			except: pass

		if self.textbox[0].Enabled:
			self.combobox[1].Enabled = True
		else:
			self.combobox[1].Enabled = False
		if self.combobox[1].Enabled == True:
			self.textbox[1].Enabled = True
		else: self.textbox[1].Enabled = False
		self.call_update()

	def call_update_activity(self):
		self.values = []
		for sheet in self.sheet_collector:
			try:
				self.temptext = str(sheet.LookupParameter(self.combobox[0].Text).AsString())
				if self.temptext not in self.values:
					self.values.append(self.temptext)
			except: pass
		self.textbox[0].Items.Clear()
		try:
			for value in self.values:
				try:
					self.textbox[0].Items.Add(str(value))
				except: pass
		except: pass
		if self.combobox[0].Enabled == True and self.combobox[0].Text in self.parameters:
			self.textbox[0].Enabled = True
		else: self.textbox[0].Enabled = False
		if self.textbox[0].Enabled == True:
			self.combobox[1].Enabled = True
		else:
			self.combobox[1].Enabled = False
			self.combobox[1].Text = ""
			self.textbox[1].Items.Clear()
			self.textbox[1].Enabled = False
		if self.combobox[1].Enabled == True and self.combobox[1].Text in self.parameters:
			self.textbox[1].Enabled = True
		else:
			self.textbox[1].Enabled = False
		self.call_update()

	def call_update(self):
		self.printable = []
		self.listbox.Items.Clear()
		self.corrupted = []
		self.temp = []
		try:
			for preset in self.presets:
				if self.combobox[0].Text in self.parameters and self.combobox[0].Enabled == True:
					if str(preset[2].LookupParameter(self.combobox[0].Text).AsString()) == self.textbox[0].Text:
						if self.combobox[1].Text in self.parameters and self.combobox[1].Enabled == True:
							if str(preset[2].LookupParameter(self.combobox[1].Text).AsString()) == self.textbox[1].Text:
								if preset[1] + " - " + preset[0] not in self.temp:
									self.temp.append(preset[1] + " - " + preset[0])
									self.printable.append(preset)
						else:
							if preset[1] + " - " + preset[0] not in self.temp:
								self.temp.append(preset[1] + " - " + preset[0])
								self.printable.append(preset)
				else:
					if preset[1] + " - " + preset[0] not in self.temp:
						self.temp.append(preset[1] + " - " + preset[0])
						self.printable.append(preset)
		except: pass
		self.temp.sort()
		for temp in self.temp:
			if temp in self.green:
				self.listbox.Items.Add(temp)
			else:
				self.listbox.Items.Add("[!] " + temp)

	def PrintPDF(self, sh, format, path, num, orientation, name, number):
		self.normalized_name = ""
		self.normalized_number = ""
		for char in name:
			if char in self.alphabet:
				self.normalized_name += char
		for char in number:
			if char in self.alphabet:
				self.normalized_number += char
		with db.Transaction(name = "Print PDF"):
			self.printManager = doc.PrintManager
			try:
				self.printManager.SelectNewPrintDriver(self.printselect.Text)
			except: pass
			self.printManager.Apply()
			self.printSetup = self.printManager.PrintSetup
			self.printSetup.CurrentPrintSetting = self.printManager.PrintSetup.CurrentPrintSetting
			try:
				self.printSetup.Delete()
			except: pass
			self.printManager.PrintRange = DB.PrintRange.Select
			self.printManager.Apply()
			self.printManager.Apply()
			self.printManager.CombinedFile = True
			self.printManager.Apply()
			self.printManager.PrintToFile = True
			self.printManager.Apply()
			self.now = datetime.datetime.now()

			if self.nametype.Text == "[номер листа] + [имя листа].pdf":
				self.printManager.PrintToFileName = path + "\\" +  self.normalized_number + "_" + self.normalized_name + ".pdf"
				self.printManager.Apply()
			elif self.nametype.Text == "[номер листа].pdf":
				self.printManager.PrintToFileName = path + "\\" +  self.normalized_number + ".pdf"
				self.printManager.Apply()
			elif self.nametype.Text == "[имя].pdf":
				self.printManager.PrintToFileName = path + "\\" +  self.normalized_name + ".pdf"
				self.printManager.Apply()
			elif self.nametype.Text == "[год] - [месяц] - [день] + [имя листа].pdf":
				self.printManager.PrintToFileName = path + "\\" +  str(self.now.year) +"-"+ str(self.now.month) +"-"+ str(self.now.day) + "_" + self.normalized_name + ".pdf"
				self.printManager.Apply()
			elif self.nametype.Text == "[год] - [месяц] - [день] + [номер листа].pdf":
				self.printManager.PrintToFileName = path + "\\" +  str(self.now.year) +"-"+ str(self.now.month) +"-"+ str(self.now.day) + "_" + self.normalized_number + ".pdf"
				self.printManager.Apply()
			elif self.nametype.Text == "[имя] + [год] - [месяц] - [день].pdf":
				self.printManager.PrintToFileName = path + "\\" +  self.normalized_name + "_" + str(self.now.year) +"-"+ str(self.now.month) +"-"+ str(self.now.day) + ".pdf"
				self.printManager.Apply()
			elif self.nametype.Text == "[номер листа] + [год] - [месяц] - [день].pdf":
				self.printManager.PrintToFileName = path + "\\" +  self.normalized_number + "_" + str(self.now.year) +"-"+ str(self.now.month) +"-"+ str(self.now.day) + ".pdf"
				self.printManager.Apply()
			elif self.nametype.Text == "[номер листа] + [имя листа] + [год] - [месяц] - [день].pdf":
				self.printManager.PrintToFileName = path + "\\" +  self.normalized_number + "_" + self.normalized_name + "_" + str(self.now.year) +"-"+ str(self.now.month) +"-"+ str(self.now.day) + ".pdf"
				self.printManager.Apply()
			elif self.nametype.Text == "[год] - [месяц] - [день] + [номер листа] + [имя листа].pdf":
				self.printManager.PrintToFileName = path + "\\" +  str(self.now.year) +"-"+ str(self.now.month) +"-"+ str(self.now.day) + "_" + self.normalized_number + "_" + self.normalized_name + ".pdf"
				self.printManager.Apply()

			try:
				self.PaperSizes = self.printManager.PaperSizes
				for PS in self.PaperSizes:
					if format == PS.Name:
						self.printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = PS
						self.printManager.Apply()
			except: pass
			
			self.printManager.PrintSetup.CurrentPrintSetting.PrintParameters.ZoomType = DB.ZoomType.Zoom
			self.printManager.Apply()
			self.printManager.PrintSetup.CurrentPrintSetting.PrintParameters.Zoom = 100
			self.printManager.Apply()
			self.printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = orientation
			self.printManager.Apply()

			self.viewSet = DB.ViewSet()
			self.viewSet.Insert(sh)
			self.viewSheetSetting = self.printManager.ViewSheetSetting
			self.viewSheetSetting.CurrentViewSheetSet.Views = self.viewSet
	
			try:
				self.printSetup.SaveAs("temp_" + format)
			except: pass

			self.viewSheetSetting.SaveAs("tempSetName")	
			self.printManager.Apply()
			self.printManager.SubmitPrint()
			self.viewSheetSetting.Delete()
			try:
				self.printSetup.Delete()
			except: pass
			self.logsheets += 1

		self.clearviewsets()

	def run_script(self):
		self.TopMost = False
		self.Get_Path()
		self.ticker_max = len(self.printable)
		self.ticker = 1
		for preset in self.printable:
			self.Text = "Печать [" + str(self.ticker) + " из " + str(self.ticker_max) + "]"
			for f in range(len(self.Formats)):
				self.FormatH = round(self.Formats[f][1], 1)
				self.FormatW = round(self.Formats[f][2], 1)
				if round(preset[5],1) == self.FormatH and round(preset[4],1) == self.FormatW:
					self.PrintPDF(sh = preset[2], format = self.Formats[f][0], path = self.path_default, num = f, orientation = self.Formats[f][3], name = preset[0], number = preset[1])
			self.ticker += 1
		self.Text = "KPLN Печать PDF"

	def clearviewsets(self):
		self.viewSets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet)
		for i in self.viewSets:
			if i.Name == "tempSetName":
				with db.Transaction(name = "clearviewsets"):
					doc.Delete(i.Id)

form = CreateWindow()

Application.Run(form)




