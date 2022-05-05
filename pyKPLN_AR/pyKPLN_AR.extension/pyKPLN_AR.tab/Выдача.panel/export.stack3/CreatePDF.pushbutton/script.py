# -*- coding: utf-8 -*-
"""
PrintPDF

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Печать PDF"
__doc__ = 'Пакетная печать PDF\n' \
          'Доступные принтеры:\n' \
          '   «PDF24 PDF» - рекомендуемый' \
          '   «Adobe PDF»\n' \
          '   «PDFCreator»' \


"""
KPLN

"""
import math
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
from threading import Thread

class FilterParameter():
	def __init__(self, parameter, name, type):
		self.Parameter = parameter
		self.StorageType = type
		self.Name = name
		try:
			self.GUID = parameter.GUID
		except:
			self.GUID = None
		self.Values = []
		self.ValueString = []

	def AddToList(self, v):
		if str(v) > "" and v != None:
			if not str(v) in self.ValueString and self.HasSymbols(str(v)):
				self.Values.append(v)
				self.ValueString.append(str(v))

	def HasSymbols(self, string):
		dict = "0123456789абвгдеёжзиклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-=+,.()[]*//\\_!@#$%^&"
		for i in dict:
			if i in string:
				return True
		return False

	def GetValue(self, sheet_pack):
		if self.Name.startswith("Основные надписи : "):
			element = sheet_pack.Title
		else:
			element = sheet_pack.Sheet
		if self.StorageType == DB.StorageType.Integer:
			try:
				try:
					return element.get_Parameter(self.GUID).AsValueString()
				except:
					return element.LookupParameter(self.Name).AsValueString()
			except: pass
		if self.StorageType == DB.StorageType.Double:
			try:
				try:
					return element.get_Parameter(self.GUID).AsDouble()
				except:
					return element.LookupParameter(self.Name).AsDouble()
			except: pass
		if self.StorageType == DB.StorageType.String:
			try:
				try:
					return element.get_Parameter(self.GUID).AsString()
				except:
					return element.LookupParameter(self.Name).AsString()
			except: pass
		if self.StorageType == DB.StorageType.ElementId:
			try:
				v = element.get_Parameter(self.Parameter.Definition.BuiltInParameter).AsElementId()
				if v != None:
					el = doc.GetElement(v)
					if el != None:
						try:
							return "{} : {}".format(el.FamilyName, el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
						except:
							return "{}".format(el.FamilyName)
				v = element.get_Parameter(self.Parameter.Definition.BuiltInParameter).AsValueString()
				if v != None:
					return str(v)
				v = element.get_Parameter(self.Parameter.Definition.BuiltInParameter).AsString()
				if v != None:
					return str(v)
			except: pass
			return None

	def AddValue(self, element):
		if self.StorageType == DB.StorageType.Integer:
			try:
				try:
					self.AddToList(element.get_Parameter(self.GUID).AsValueString())
				except:
					self.AddToList(element.LookupParameter(self.Name).AsValueString())
			except: pass
		if self.StorageType == DB.StorageType.Double:
			try:
				try:
					self.AddToList(element.get_Parameter(self.GUID).AsDouble())
				except:
					self.AddToList(element.LookupParameter(self.Name).AsDouble())
			except: pass
		if self.StorageType == DB.StorageType.String:
			try:
				try:
					self.AddToList(element.get_Parameter(self.GUID).AsString())
				except:
					self.AddToList(element.LookupParameter(self.Name).AsString())
			except: pass
		if self.StorageType == DB.StorageType.ElementId:
			try:
				v = element.get_Parameter(self.Parameter.Definition.BuiltInParameter).AsElementId()
				if v != None:
					el = doc.GetElement(v)
					if el != None:
						try:
							self.AddToList("{} : {}".format(el.FamilyName, el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()))
						except:
							self.AddToList("{}".format(el.FamilyName))
						return
				v = element.get_Parameter(self.Parameter.Definition.BuiltInParameter).AsValueString()
				if v != None:
					self.AddToList(str(v))
					return
				v = element.get_Parameter(self.Parameter.Definition.BuiltInParameter).AsString()
				if v != None:
					self.AddToList(str(v))
			except: pass

class FilterParameters():
	def __init__(self, PrintableSheets):
		self.Parameters = []
		self.TitleParameters = []
		for j in PrintableSheets[0].Sheet.Parameters:
			if j.UserModifiable:
				self.Parameters.append(FilterParameter(j, j.Definition.Name, j.StorageType))
		for j in PrintableSheets[0].Title.Parameters:
			if j.UserModifiable:
				self.TitleParameters.append(FilterParameter(j, "Основные надписи : {}".format(j.Definition.Name), j.StorageType))
		for parameter in self.Parameters:
			for printable_sheet in PrintableSheets:
				try:
					parameter.AddValue(printable_sheet.Sheet)
				except:
					pass
		for parameter in self.TitleParameters:
			for printable_sheet in PrintableSheets:
				try:
					parameter.AddValue(printable_sheet.Title)
				except:
					pass

class MainWindowFilter():
	def __init__(self, parent, location, parameter_set):
		self.Parent = parent
		self.ParameterSet = parameter_set
		self.PickedParameter = None
		#
		self.Label = Label()
		self.Label.Parent = parent
		self.Label.Location = Point(6, 16 + location)
		self.Label.Text = "Фильтр #1"
		self.Label.BackColor = Color.Transparent
		#
		self.Condition = ComboBox()
		self.Condition.Parent = parent
		self.Condition.Size = Size(120, 21)
		self.Condition.DropDownWidth = 300
		self.Condition.Location = Point(5, 39 + location)
		self.Condition.DropDownStyle = ComboBoxStyle.DropDownList
		self.Condition.FlatStyle = FlatStyle.Flat
		self.AddedParameters = []
		self.Condition.Items.Add("<без фильтра>")
		for i in range(0, len(parameter_set.Parameters)):
			if len(parameter_set.Parameters[i].ValueString) != 0:
				self.Condition.Items.Add(parameter_set.Parameters[i].Name)
				self.AddedParameters.append(parameter_set.Parameters[i])
		for i in range(0, len(parameter_set.TitleParameters)):
			if len(parameter_set.TitleParameters[i].ValueString) != 0:
				self.Condition.Items.Add(parameter_set.TitleParameters[i].Name)
				self.AddedParameters.append(parameter_set.TitleParameters[i])
		self.Condition.Text = "<без фильтра>"
		#
		self.EquationType = ComboBox()
		self.EquationType.Parent = parent
		self.EquationType.Size = Size(96, 21)
		self.EquationType.Location = Point(142, 39 + location)
		self.EquationType.DropDownStyle = ComboBoxStyle.DropDownList
		self.EquationType.Items.Add("Равно")
		self.EquationType.Items.Add("Не равно")
		self.EquationType.Text = "Равно"
		#
		self.Value = ComboBox()
		self.Value.Parent = parent
		self.Value.Size = Size(120, 21)
		self.Value.DropDownWidth = 300
		self.Value.Location = Point(255, 39 + location)
		self.Value.DropDownStyle = ComboBoxStyle.DropDownList
		self.Value.FlatStyle = FlatStyle.Flat
		self.Value.Items.Add("<none>")
		self.Value.Text = "<none>"

		self.Value.SelectedIndexChanged += self.EventValueIndexChanged
		self.EquationType.SelectedIndexChanged += self.EventValueIndexChanged
		self.Condition.SelectedIndexChanged += self.EventIndexChanged

	def EventIndexChanged(self, sender, event):
		if self.Condition.SelectedIndex == 0:
			self.PickedParameter = None
			self.Value.Items.Clear()
			self.Value.Items.Add("<none>")
			self.Value.Text = "<none>"
		else:
			self.PickedParameter = self.AddedParameters[self.Condition.SelectedIndex-1]
			self.Value.Items.Clear()
			self.Value.Items.Add("<none>")
			self.Value.Text = "<none>"
			for i in range(0, len(self.PickedParameter.ValueString)):
				self.Value.Items.Add(self.PickedParameter.ValueString[i])	
			self.Parent.Parent.Update()

	def EventValueIndexChanged(self, sender, event):
		self.Parent.Parent.Update()

	def Disable(self):
		self.Label.Enabled = False
		self.Condition.Enabled = False
		self.EquationType.Enabled = False
		self.Value.Enabled = False

	def Enable(self):
		self.Label.Enabled = True
		self.Condition.Enabled = True
		self.EquationType.Enabled = True
		self.Value.Enabled = True

class MainWindow(Form):
	def __init__(self):
		global sheets
		self.PrintManager = doc.PrintManager
		self.SuspendLayout()
		self.Name = "KPLN_Batch_PDF"
		self.Text = "KPLN : Batch PDF"
		self.DefaultPath = "C:\\"
		self.Size = Size(800, 600)
		self.MinimumSize = Size(800, 600)
		self.MaximumSize = Size(800, 600)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = True
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.TopMost = True
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
		#Create panel with sheet list
		self.Panel = Panel()
		self.Panel.Parent = self
		self.Panel.Size = Size(373, 538)
		self.Panel.Location = Point(12, 12)
		self.Panel.AutoScroll = True
		self.Sheets = []
		self.SortableSheets = []
		self.SheetKeys = []
		for title in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType():
			sheet = doc.GetElement(title.OwnerViewId)
			self.SortableSheets.append(PrintableSheet(sheet, title))
			self.SheetKeys.append("{}_{}_{}".format(sheet.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString(), sheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString(), str(sheet.Id.IntegerValue)))
		self.SheetKeys.sort()
		for i in self.SheetKeys:
			for z in self.SortableSheets:
				if "{}_{}_{}".format(z.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString(), z.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString(), str(z.Sheet.Id.IntegerValue)) == i:
					self.Sheets.append(z)
		self.SheetCount = 0
		#
		self.ParameterSet = FilterParameters(self.Sheets)
		self.GroupFilter = GroupBox()
		self.GroupFilter.Parent = self
		self.GroupFilter.Size = Size(381, 115)
		self.GroupFilter.Location = Point(391, 12)
		self.GroupFilter.Text = "Настройки фильтрации"
		self.filter_01 = MainWindowFilter(self.GroupFilter, 0, self.ParameterSet)
		self.filter_02 = MainWindowFilter(self.GroupFilter, 48, self.ParameterSet)
		#
		self.GroupPrinter = GroupBox()
		self.GroupPrinter.Parent = self
		self.GroupPrinter.Size = Size(381, 100)
		self.GroupPrinter.Location = Point(391, 134)
		self.GroupPrinter.Text = "Настройки принтера"
		self.LabelPrinter = Label()
		self.LabelPrinter.Parent = self.GroupPrinter
		self.LabelPrinter.Location = Point(16, 29)
		self.LabelPrinter.Size = Size(110, 14)
		self.LabelPrinter.Text = "Выберите принтер:"
		self.LabelPrinter.BackColor = Color.Transparent
		self.CBPrinter = ComboBox()
		self.CBPrinter.Parent = self.GroupPrinter
		self.CBPrinter.Size = Size(230, 21)
		self.CBPrinter.Location = Point(130, 26)
		self.CBPrinter.DropDownStyle = ComboBoxStyle.DropDownList
		for printer in ["PDF24 PDF", "PDF24 FAX", "PDF24", "PDF Creator", "Adobe PDF", "PDFCreator", "Adobe", "PDF Adobe"]:
			try:
				self.PrintManager.SelectNewPrintDriver(printer)
				self.CBPrinter.Items.Add(printer)
				self.CBPrinter.Text = printer
			except: pass
		self.LabelName = Label()
		self.LabelName.Parent = self.GroupPrinter
		self.LabelName.Location = Point(16, 59)
		self.LabelName.Size = Size(110, 14)
		self.LabelName.Text = "Выберите формат:"
		self.LabelName.BackColor = Color.Transparent
		self.CBName = ComboBox()
		self.CBName.Parent = self.GroupPrinter
		self.CBName.Size = Size(230, 21)
		self.CBName.Location = Point(130, 56)
		self.CBName.DropDownStyle = ComboBoxStyle.DropDownList
		self.CBName.Items.Add("НОМЕР-ИМЯ.pdf")
		self.CBName.Items.Add("НОМЕР.pdf")
		self.CBName.Items.Add("ИМЯ.pdf")
		self.CBName.Items.Add("ГГ-ММ-ДД_ИМЯ.pdf")
		self.CBName.Items.Add("ГГ-ММ-ДДь_НОМЕР.pdf")
		self.CBName.Items.Add("ИМЯ_ГГ-ММ-ДД.pdf")
		self.CBName.Items.Add("НОМЕР_ГГ-ММ-ДД.pdf")
		self.CBName.Items.Add("НОМЕР-ИМЯ_ГГ-ММ-ДД.pdf")
		self.CBName.Items.Add("ГГ-ММ-ДД_НОМЕР-ИМЯ.pdf")
		self.CBName.Text = "ГГ-ММ-ДД_НОМЕР-ИМЯ.pdf"
		self.PrintBtn = Button(Text="Печать")
		self.PrintBtn.Parent = self
		self.PrintBtn.Location = Point(391, 240)
		if self.CBPrinter.Text != "":
			self.PrintBtn.Enabled = True
		else:
			self.PrintBtn.Enabled = False
		self.PrintBtn.Click += self.StartPrintPDF
		#
		for sheet in self.Sheets:
			if self.Filter(sheet):
				sheet.Panel.Parent = self.Panel
				sheet.Panel.Location = Point(5, 35 * self.SheetCount + 5)
				self.SheetCount += 1
				if sheet.SheetFormat.Name == "Нестандартный":
					sheet.SetStatusError()
				elif sheet.SheetFormat.Width == sheet.TitleFormat.Width and sheet.SheetFormat.Height == sheet.TitleFormat.Height:
					sheet.SetStatusDefault()
				else:
					sheet.SetStatusError()
			else:
				pass
		self.ResumeLayout()

	def Disable(self):
		self.PrintBtn.Enabled = False
		self.CBName.Enabled = False
		self.CBPrinter.Enabled = False
		self.LabelName.Enabled = False
		self.LabelPrinter.Enabled = False
		self.filter_01.Disable()
		self.filter_02.Disable()

	def Enable(self):
		self.PrintBtn.Enabled = True
		self.CBName.Enabled = True
		self.CBPrinter.Enabled = True
		self.LabelName.Enabled = True
		self.LabelPrinter.Enabled = True
		self.filter_01.Enable()
		self.filter_02.Enable()

	def StartPrintPDF(self, sender, args):
		#self.Disable()
		#self.Browse()
		#try:
		#	new_thread = MyThread("PrinterThread", self.Sheets, self.DefaultPath, self.CBName.Text, self.CBPrinter.Text, self)
		#	new_thread.start()
		#	new_thread.run()
		#except Exception as e: print(str(e))
		#self.Refresh()
		#self.Enable()
		self.Disable()
		self.Browse()
		for sheet in self.Sheets:
			if sheet.Status == 1 and sheet.CheckBox.Checked:
				sheet.SetStatusPrinting()
				self.Refresh()
				try:
					sheet.PrintPDF(sheet.Sheet, sheet.TitleFormat, self.DefaultPath, sheet.TitleFormat.Orientation, sheet.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString(), sheet.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString(), self.CBName.Text, self.CBPrinter.Text)
					sheet.SetStatusDone()
					self.Refresh()
				except Exception as e:
					sheet.SetStatusError()
					print(str(e))
				self.Refresh()
		self.Enable()

	def Browse(self):
		dialog = FolderBrowserDialog()
		dialog.Description = "Выберите папку для сохранения листов в формате PDF"
		if (dialog.ShowDialog(self) == DialogResult.OK):
			self.DefaultPath = dialog.SelectedPath

	def Filter(self, sheet):
		try:
			v_1 = self.filter_01.PickedParameter.GetValue(sheet)
		except:
			v_1 = None
		try:
			v_2 = self.filter_02.PickedParameter.GetValue(sheet)
		except:
			v_2 = None
		if self.filter_01.Condition.Text == "<без фильтра>" or self.filter_01.Value.Text == "<none>":
			if self.filter_02.Condition.Text == "<без фильтра>" or self.filter_02.Value.Text == "<none>":
				return True
			elif self.filter_02.EquationType.Text == "Равно":
				if v_2 == self.filter_02.Value.Text:
					return True
			elif self.filter_02.EquationType.Text == "Не равно":
				if v_2 != self.filter_02.Value.Text:
					return True
		elif self.filter_01.EquationType.Text == "Равно":
			if v_1 == self.filter_01.Value.Text:
				if self.filter_02.Condition.Text == "<без фильтра>" or self.filter_02.Value.Text == "<none>":
					return True
				elif v_2 == self.filter_02.Value.Text:
					return True
		elif self.filter_01.EquationType.Text == "Не равно":
			if v_1 != self.filter_01.Value.Text:
				if self.filter_02.Condition.Text == "<без фильтра>" or self.filter_02.Value.Text == "<none>":
					return True
				elif self.filter_02.EquationType.Text == "Равно":
					if v_2 == self.filter_02.Value.Text:
						return True
				elif self.filter_02.EquationType.Text == "Не равно":
					if v_2 != self.filter_02.Value.Text:
						return True
		return False

	def Update(self):
		self.SuspendLayout()
		self.SheetCount = 0
		for sheet in self.Sheets:
			sheet.SetStatusDisable()
		for sheet in self.Sheets:
			if self.Filter(sheet):
				sheet.Panel.Location = Point(5, 35 * self.SheetCount + 5)
				self.SheetCount += 1
				if sheet.SheetFormat.Name == "Нестандартный":
					sheet.SetStatusError()
				elif sheet.SheetFormat.Width == sheet.TitleFormat.Width and sheet.SheetFormat.Height == sheet.TitleFormat.Height:
					sheet.SetStatusDefault()
				else:
					sheet.SetStatusError()
			else:
				pass
		self.ResumeLayout()

class PrintableFormat():
	def __init__(self, name, height, width, orientation):
		self.Name = name
		self.Width = width
		self.Height = height
		self.Orientation = orientation
		self.ToolTip = ToolTip()
		self.ToolTip.ToolTipTitle = "Формат: {}".format(self.Name)

Formats = [PrintableFormat("A4", 297, 210, DB.PageOrientationType.Portrait),
				  PrintableFormat("A4", 210, 297, DB.PageOrientationType.Landscape),
				  PrintableFormat("A3", 420, 297, DB.PageOrientationType.Portrait),
				  PrintableFormat("A3", 297, 420, DB.PageOrientationType.Landscape),
				  PrintableFormat("A2", 594, 420, DB.PageOrientationType.Portrait),
				  PrintableFormat("A2", 420, 594, DB.PageOrientationType.Landscape),
				  PrintableFormat("A1", 841, 594, DB.PageOrientationType.Portrait),
				  PrintableFormat("A1", 594, 841, DB.PageOrientationType.Landscape),
				  PrintableFormat("A0", 1189, 841, DB.PageOrientationType.Portrait),
				  PrintableFormat("A0", 841, 1189, DB.PageOrientationType.Landscape),
				  PrintableFormat("A0x2", 1189, 1682, DB.PageOrientationType.Landscape),
				  PrintableFormat("A0x3", 1189, 2523, DB.PageOrientationType.Landscape),
				  PrintableFormat("A1x3", 841, 1783, DB.PageOrientationType.Landscape),
				  PrintableFormat("A1x4", 841, 2378, DB.PageOrientationType.Landscape),
				  PrintableFormat("A1x5", 841, 2973, DB.PageOrientationType.Landscape),
				  PrintableFormat("A2x3", 594, 1261, DB.PageOrientationType.Landscape),
				  PrintableFormat("A2x4", 594, 1682, DB.PageOrientationType.Landscape),
				  PrintableFormat("A2x5", 594, 2102, DB.PageOrientationType.Landscape),
				  PrintableFormat("A3x3", 420, 891, DB.PageOrientationType.Landscape),
				  PrintableFormat("A3x4", 420, 1189, DB.PageOrientationType.Landscape),
				  PrintableFormat("A3x5", 420, 1486, DB.PageOrientationType.Landscape),
				  PrintableFormat("A3x6", 420, 1783, DB.PageOrientationType.Landscape),
				  PrintableFormat("A3x7", 420, 2080, DB.PageOrientationType.Landscape),
				  PrintableFormat("A4x3", 297, 630, DB.PageOrientationType.Landscape),
				  PrintableFormat("A4x4", 297, 841, DB.PageOrientationType.Landscape),
				  PrintableFormat("A4x5", 297, 1051, DB.PageOrientationType.Landscape),
				  PrintableFormat("A4x6", 297, 1261, DB.PageOrientationType.Landscape),
				  PrintableFormat("A4x7", 297, 1471, DB.PageOrientationType.Landscape),
				  PrintableFormat("A4x8", 297, 1682, DB.PageOrientationType.Landscape),
				  PrintableFormat("A4x9", 297, 1892, DB.PageOrientationType.Landscape),
				  PrintableFormat("A0x2", 1682, 1189, DB.PageOrientationType.Portrait),
				  PrintableFormat("A0x3", 2523, 1189, DB.PageOrientationType.Portrait),
				  PrintableFormat("A1x3", 1783, 841, DB.PageOrientationType.Portrait),
				  PrintableFormat("A1x4", 2378, 841, DB.PageOrientationType.Portrait),
				  PrintableFormat("A1x5", 2973, 841, DB.PageOrientationType.Portrait),
				  PrintableFormat("A2x3", 1261, 594, DB.PageOrientationType.Portrait),
				  PrintableFormat("A2x4", 1682, 594, DB.PageOrientationType.Portrait),
				  PrintableFormat("A2x5", 2102, 594, DB.PageOrientationType.Portrait),
				  PrintableFormat("A3x3", 891, 420, DB.PageOrientationType.Portrait),
				  PrintableFormat("A3x4", 1189, 420, DB.PageOrientationType.Portrait),
				  PrintableFormat("A3x5", 1486, 420, DB.PageOrientationType.Portrait),
				  PrintableFormat("A3x6", 1783, 420, DB.PageOrientationType.Portrait),
				  PrintableFormat("A3x7", 2080, 420, DB.PageOrientationType.Portrait),
				  PrintableFormat("A4x3", 630, 297, DB.PageOrientationType.Portrait),
				  PrintableFormat("A4x4", 841, 297, DB.PageOrientationType.Portrait),
				  PrintableFormat("A4x5", 1051, 297, DB.PageOrientationType.Portrait),
				  PrintableFormat("A4x6", 1261, 297, DB.PageOrientationType.Portrait),
				  PrintableFormat("A4x7", 1471, 297, DB.PageOrientationType.Portrait),
				  PrintableFormat("A4x8", 1682, 297, DB.PageOrientationType.Portrait),
				  PrintableFormat("A4x9", 1892, 297, DB.PageOrientationType.Portrait)]

Image_Sheet_Portrait = Image.FromFile('Z:\\pyRevit\\Source\\SP.png')
Image_Sheet_Portrait_Errored = Image.FromFile('Z:\\pyRevit\\Source\\SPA.png')
Image_Sheet_Landscape = Image.FromFile('Z:\\pyRevit\\Source\\SL.png')
Image_Sheet_Landscape_Errored = Image.FromFile('Z:\\pyRevit\\Source\\SLA.png')

Image_Background_Default = Image.FromFile('Z:\\pyRevit\\Source\\BG.png')
Image_Background_Checked = Image.FromFile('Z:\\pyRevit\\Source\\BG_Checked.png')
Image_Background_Error = Image.FromFile('Z:\\pyRevit\\Source\\BG_Error.png')
Image_Background_Inactive = Image.FromFile('Z:\\pyRevit\\Source\\BG_Inactive.png')


class PrintableSheet():
	def __init__(self, sheet, title):
		global Image_Sheet_Portrait
		global Image_Sheet_Portrait_Errored
		global Image_Sheet_Landscape
		global Image_Sheet_Landscape_Errored
		self.Status = 0
		self.Title = title
		self.Sheet = sheet
		self.StringOrientation = ""
		self.Name = sheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString()
		self.Number = sheet.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER).AsString()
		self.FactWidth = int(round(math.fabs(sheet.Outline.Max.U - sheet.Outline.Min.U) * 304.8,0))
		self.FactHeight = int(round(math.fabs(sheet.Outline.Max.V - sheet.Outline.Min.V) * 304.8,0))
		self.TitleBox = title.Symbol.get_BoundingBox(None)
		self.TitleWidth = int(round(math.fabs(self.TitleBox.Max.X - self.TitleBox.Min.X) * 304.8,0))
		self.TitleHeight = int(round(math.fabs(self.TitleBox.Max.Y - self.TitleBox.Min.Y) * 304.8,0))
		self.SheetFormat = self.GetFormat(self.FactWidth, self.FactHeight)
		self.TitleFormat = self.GetFormat(self.TitleWidth, self.TitleHeight)

		#
		if self.TitleFormat != None:
			if self.TitleFormat.Orientation == DB.PageOrientationType.Portrait:
				self.Image_Default = Image_Sheet_Portrait
				self.Image_Errored = Image_Sheet_Portrait_Errored
				self.StringOrientation = "Портретный"
			else:
				self.Image_Default = Image_Sheet_Landscape
				self.Image_Errored = Image_Sheet_Landscape_Errored
				self.StringOrientation = "Альбомный"
			#
			self.Panel = Panel()
			self.Panel.BackgroundImage = Image.FromFile('Z:\\pyRevit\\Source\\BG.png')
			self.Panel.Size = Size(344, 30)
			self.Pending = PictureBox()
			self.Pending.Visible = False
			self.Pending.Location = Point(0,0)
			self.Pending.Parent = self.Panel
			self.Pending.Size = Size(344, 30)
			self.Pending.Image = Image.FromFile('Z:\\pyRevit\\Source\\BG_Pending.gif')
			self.PictureBox = PictureBox()
			self.PictureBox.BackColor = Color.Transparent
			self.PictureBox.Parent = self.Panel
			self.PictureBox.Location = Point(36,3)
			self.PictureBox.Size = Size(24, 24)
			self.PictureBox.Image = self.Image_Default
			self.PictureBox.Cursor = Cursors.Hand
			self.PictureBox.Click += self.EventOpenSheet
			if self.SheetFormat.Width == self.TitleFormat.Width and self.SheetFormat.Height == self.TitleFormat.Height:
				if self.SheetFormat.Name == "Нестандартный":
					self.TitleFormat.ToolTip.SetToolTip(self.PictureBox, "{}x{}h ({}) : Неудалось подобрать формат листа!".format(str(self.TitleWidth), str(self.TitleHeight), self.StringOrientation))
				else:
					self.TitleFormat.ToolTip.SetToolTip(self.PictureBox, "{}x{}h ({})".format(str(self.TitleWidth), str(self.TitleHeight), self.StringOrientation))
			else:
				self.TitleFormat.ToolTip.SetToolTip(self.PictureBox, "{}x{}h ({}) : Элементы выходят за границы печатаемой области!".format(str(self.TitleWidth), str(self.TitleHeight), self.StringOrientation))
			self.LabelTitle = Label()
			self.LabelTitle.AutoSize = False
			self.LabelTitle.Size = Size(274, 28)
			self.LabelTitle.Parent = self.Panel
			self.LabelTitle.BackColor = Color.Transparent
			self.LabelTitle.Location = Point(67,1)
			self.LabelTitle.Text = "{} - {}".format(str(self.Number), str(self.Name))
			self.CheckBox = CheckBox()
			self.CheckBox.BackColor = Color.Transparent
			self.CheckBox.Parent = self.Panel
			self.CheckBox.Location = Point(11,4)
			self.CheckBox.Text = ""
			self.CheckBox.Checked = True
			self.CheckBox.CheckedChanged += self.EventCheckedChanged

	def EventOpenSheet(self, sender, event):
		try:
			uidoc.ActiveView = self.Sheet
		except: pass

	def EventCheckedChanged(self, sender, event):
		if self.Status == 1 or self.Status == 5:
			if self.CheckBox.Checked:
				self.SetStatusDefault()
			else:
				self.SetStatusDisabled()

	def SetStatusDisable(self):
		self.Status = 0
		self.Panel.Visible = False

	def SetStatusDefault(self):
		self.Status = 1
		self.Panel.Visible = True
		global Image_Background_Default
		self.Pending.Visible = False
		self.PictureBox.Image = self.Image_Default
		self.Panel.BackgroundImage = Image_Background_Default
		self.LabelTitle.Enabled = True
		self.CheckBox.Enabled = True	
	
	def SetStatusError(self):
		self.Status = 2
		self.Panel.Visible = True
		global Image_Background_Error
		self.Pending.Visible = False
		self.PictureBox.Image = self.Image_Errored
		self.Panel.BackgroundImage = Image_Background_Error
		self.LabelTitle.Enabled = False
		self.CheckBox.Enabled = False
		self.CheckBox.Checked = False

	def SetStatusPrinting(self):
		self.Status = 3
		self.Panel.Visible = True
		global Image_Background_Default
		self.Pending.Visible = True
		self.PictureBox.Image = self.Image_Default
		self.Panel.BackgroundImage = Image_Background_Default
		self.LabelTitle.Enabled = False
		self.CheckBox.Enabled = False

	def SetStatusDone(self):
		self.Status = 4
		self.Panel.Visible = True
		global Image_Background_Checked
		self.Pending.Visible = False
		self.PictureBox.Image = self.Image_Default
		self.Panel.BackgroundImage = Image_Background_Checked
		self.LabelTitle.Enabled = False
		self.CheckBox.Enabled = False
		self.CheckBox.Checked = False

	def SetStatusDisabled(self):
		self.Status = 5
		self.Panel.Visible = True
		global Image_Background_Inactive
		self.Pending.Visible = False
		self.PictureBox.Image = self.Image_Default
		self.Panel.BackgroundImage = Image_Background_Inactive
		self.LabelTitle.Enabled = False
		self.CheckBox.Enabled = True
		
	def GetFormat(self, width, height):
		global Formats
		for printable_format in Formats:
			if math.fabs(printable_format.Width - width) < 2 and math.fabs(printable_format.Height - height) < 2:
				return printable_format
		if width >= height:
			return PrintableFormat("Нестандартный", width, height, DB.PageOrientationType.Landscape)
		else:
			return PrintableFormat("Нестандартный", width, height, DB.PageOrientationType.Portrait)

	def GetNormalizedString(self, string):
		dict = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZабвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ1234567890 «»=№-+.,()_[]"
		normalized_string = ""
		for char in string:
			if char in dict:
				normalized_string += char
		return normalized_string

	def GetCurrentDate(self):
		now = datetime.datetime.now()
		return "{}-{}-{}".format(str(now.year), str(now.month), str(now.day))

	def GetPath(self, path, number, name, setup):
		if setup == "НОМЕР-ИМЯ.pdf":
			return "{}\\{}_{}.pdf".format(path, number, name)
		elif setup == "НОМЕР.pdf":
			return "{}\\{}.pdf".format(path, number)
		elif setup == "ИМЯ.pdf":
			return "{}\\{}.pdf".format(path, name)
		elif setup == "ГГ-ММ-ДД_ИМЯ.pdf":
			return "{}\\{}_{}.pdf".format(path, self.GetCurrentDate(), name)
		elif setup == "ГГ-ММ-ДД_НОМЕР.pdf":
			return "{}\\{}_{}.pdf".format(path, self.GetCurrentDate(), number)
		elif setup == "ИМЯ_ГГ-ММ-ДД.pdf":
			return "{}\\{}_{}.pdf".format(path, name, self.GetCurrentDate())
		elif setup == "НОМЕР_ГГ-ММ-ДД.pdf":
			return "{}\\{}_{}.pdf".format(path, number, self.GetCurrentDate())
		elif setup == "НОМЕР-ИМЯ_ГГ-ММ-ДД.pdf":
			return "{}\\{}_{}_{}.pdf".format(path, name, number, self.GetCurrentDate())
		elif setup == "ГГ-ММ-ДД_НОМЕР-ИМЯ.pdf":
			return "{}\\{}_{}_{}.pdf".format(path, self.GetCurrentDate(), name, number)
	
	def GetPaperSize(self, format, print_manager):
		try:
			paper_sizes = print_manager.PaperSizes
			for ps in paper_sizes:
				if format == ps.Name:
					return ps
		except Exception as e: print(str(e))
		return False

	def PrintPDF(self, sheet, format, path, page_orientation, name, number, date_setup, printer):
		safe_name = self.GetNormalizedString(name)
		safe_number = self.GetNormalizedString(number)
		with db.Transaction(name = "pyKPLN"):
			print_manager = doc.PrintManager
			try:
				print_manager.SelectNewPrintDriver(printer)
			except: pass
			print_manager.Apply()
			print_setup = print_manager.PrintSetup
			print_setup.CurrentPrintSetting = print_manager.PrintSetup.CurrentPrintSetting
			try:
				print_setup.Delete()
			except: pass
			print_manager.PrintRange = DB.PrintRange.Select
			print_manager.Apply()
			print_manager.CombinedFile = True
			print_manager.Apply()
			print_manager.PrintToFile = True
			print_manager.Apply()
			print_manager.PrintToFileName = self.GetPath(path, safe_number, safe_name, date_setup)
			print_manager.Apply()			
			print_manager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = self.GetPaperSize(format.Name, print_manager)
			print_manager.Apply()
			print_manager.PrintSetup.CurrentPrintSetting.PrintParameters.ZoomType = DB.ZoomType.Zoom
			print_manager.Apply()
			print_manager.PrintSetup.CurrentPrintSetting.PrintParameters.Zoom = 100
			print_manager.Apply()
			print_manager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = page_orientation
			print_manager.Apply()
			view_set = DB.ViewSet()
			view_set.Insert(sheet)
			view_sheet_setting = print_manager.ViewSheetSetting
			view_sheet_setting.CurrentViewSheetSet.Views = view_set
			try:
				print_setup.SaveAs("sys_" + format.Name)
			except: pass
			view_sheet_setting.SaveAs("sys_setting")	
			print_manager.Apply()
			doc.Regenerate()
			print_manager.SubmitPrint()
			view_sheet_setting.Delete()
			doc.Regenerate()
			try:
				print_setup.Delete()
			except: pass

form = MainWindow()
Application.Run(form)