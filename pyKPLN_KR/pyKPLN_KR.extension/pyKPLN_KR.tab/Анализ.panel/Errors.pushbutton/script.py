# -*- coding: utf-8 -*-
"""
SelectByType

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Изолировать"
__doc__ = 'Выбирает все ошибки на активном виде (по типу)' \

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
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
                              Group, ElementId, Transaction, OverrideGraphicSettings,\
                              Color, FillPatternElement, FillPatternTarget
uiapp = __revit__.Application
from System.Collections.Generic import List
import System.Windows
from Microsoft.Win32 import OpenFileDialog
from HTMLParser import HTMLParser
from collections import Counter
from itertools import groupby
from operator import itemgetter
import codecs

class MyHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.temp_data = []
        self.data1 = []

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self.data1.append(self.temp_data)
            self.temp_data = []

    def handle_data(self, data):
        if data != "\n" and data != "\r\n" and data not in known_errors:
            pass
            self.temp_data.append(data)

def extractId(list1):
    return [i for j in list1 for i in id_expression.findall(j)]

def openFile():
    try:
        dialog = OpenFileDialog()
        dialog.Filter = "All Files|*.*"
        if dialog.ShowDialog() == True:
            selectedFile = dialog.FileName
        return selectedFile
    except:
        pass

def overrideElements(ids, color=(255,0,0), pattern="Solid fill"):
    """override elements"""
    ogs = OverrideGraphicSettings()
    # set color
    ogs.SetProjectionFillColor(Color(*color))
    # set fill pattern
    fillPattern = FillPatternElement
    fillPatternId = fillPattern.GetFillPatternElementByName(doc,
                        FillPatternTarget.Drafting, pattern).Id
    ogs.SetProjectionFillPatternId(fillPatternId)
    for id in ids:
        doc.ActiveView.SetElementOverrides(id, ogs)

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "Анализ"
		self.Text = "Выбрать по типу"
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
		self.c_name.Text = "Тип"
		self.c_name.Text = "Тип"
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

		self.Text = "Выберите типы ошибок"
		self.item = []

		self.warnings = doc.GetWarnings()
		self.description_uniq = []
		for w in self.warnings:
			self.description = str(w.GetDescriptionText())
			if not self.in_list(self.description, self.description_uniq):
				self.description_uniq.append(self.description)
		for des in self.description_uniq:
			try:
				self.item.append(ListViewItem())
				self.item[len(self.item)-1].Text = ""
				self.item[len(self.item)-1].Checked = False
				self.item[len(self.item)-1].SubItems.Add(des)
				self.listbox.Items.Add(self.item[len(self.item)-1])
			except: pass

	def in_list(self, element, list):
		for i in list:
			if element == i:
				return True
		return False

	def OnOk(self, sender, args):
		self.select_errors(self.define_errors())
		self.Close()

	def define_errors(self):
		warnings = doc.GetWarnings()
		list_of_errors = []
		for i in self.item:
			if i.Checked:
				si = i.SubItems
				viewname = si[1].Text
				for w in warnings:
					if str(w.GetDescriptionText()) == viewname:
						for e in w.GetFailingElements():
							list_of_errors.append(e)
		return list_of_errors

	def OnCancel(self, sender, args):
		self.Close()

	def select_errors(self, defined_elements):
		if len(defined_elements) != 0:
			selection_items = []
			active_view = doc.ActiveView.Id
			collector = DB.FilteredElementCollector(doc, active_view).WhereElementIsNotElementType().ToElementIds()
			for collector_element in collector:
				try:
					for picked_element in defined_elements:
						if str(collector_element) == str(picked_element):
							selection_items.append(collector_element)
							break
				except: pass
			if len(defined_elements) != 0:
				selection = uidoc.Selection
				collection = List[DB.ElementId]([element for element in selection_items])
				selection.SetElementIds(collection)
uiapp = __revit__.Application
version =  uiapp.VersionNumber
reportFromFile = None
if int(version) >= 2018:
	form = CreateWindow()
	Application.Run(form)
else:
	known_errors = ("  ", "td&gt;  ", "&amp;")
	version =  uiapp.VersionNumber
	reportFromFile = None
	reportFromFile = openFile()
	if reportFromFile:
		parser = MyHTMLParser()
		test = []
		with codecs.open(reportFromFile, 'r', "utf-16") as f:
			for line in f:
				test.append(line)
				pass
				parser.feed(line)

		parser.data1.pop(0)
		parser.data1.append(parser.temp_data)
		report_data = parser.data1
		parser.close()

		report_errors = [i.pop(0) for i in report_data]
		group_errors = Counter(report_errors).most_common()

		id_expression = re.compile(" (?:id|Код) (.*)  ")
		elements_id = map(extractId, report_data)

		# all warning elements id
		allelements = [ElementId(int(item)) for sublist in elements_id for item in sublist]
		element_to_isolate = List[ElementId](allelements)

		curview = uidoc.ActiveGraphicalView
		with Transaction(doc, 'Isolate Warnings') as t:
			t.Start()
			# isolate
			curview.IsolateElementsTemporary(element_to_isolate)
			# override
			overrideElements(element_to_isolate)
			t.Commit()
	else:
		warnings = doc.GetWarnings()
		war_elements = [w.GetAdditionalElements() for w in warnings]
		print(war_elements)