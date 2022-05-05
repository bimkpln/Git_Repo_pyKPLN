# -*- coding: utf-8 -*-
"""
arm_element

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Расход\nстали"
__doc__ = '' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from rpw.ui import Pick
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import TextInput, Alert, select_folder
import datetime
import System
from System.Windows.Forms import *
from System.Drawing import *

class CreateWindow(Form):
	def __init__(self):
		global value
		self.Name = "KPLN Расход стали"
		self.Text = "KPLN Расход стали"
		self.Size = Size(418, 608)
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
		self.ControlBox = True
		self.TopMost = True
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.MinimumSize = Size(418, 180)
		self.MaximumSize = Size(418, 180)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.CenterToScreen()
		self.BackColor = Color.FromArgb(255, 255, 255)

		self.Value = "None"

		self.lbl1 = Label(Text = 'Расход стали кг/куб.м:')
		self.lbl1.Parent = self
		self.lbl1.Location = Point(20, 10)
		self.lbl1.Size = Size(150, 15)
		self.lbl2 = Label(Text = 'Шаг значений:')
		self.lbl2.Parent = self
		self.lbl2.Location = Point(20, 70)
		self.lbl2.Size = Size(150, 15)

		self.cb_main = ComboBox()
		self.cb_main.Parent = self
		self.cb_main.Location = Point(20, 30)
		self.cb_main.Size = Size(150, 15)
		self.cb_main.SelectedIndexChanged += self.update_value
		self.cb_main.DropDownStyle = ComboBoxStyle.DropDownList

		self.step_range = ["5", "10", "20", "25", "50", "100"]
		self.cb_step = ComboBox()
		self.cb_step.Parent = self
		self.cb_step.Location = Point(20, 90)
		self.cb_step.Size = Size(150, 15)
		self.cb_step.DropDownStyle = ComboBoxStyle.DropDownList
		for i in self.step_range:
			self.cb_step.Items.Add(i)
		self.cb_step.SelectedIndexChanged += self.update_range
		self.Settings = str(doc.ProjectInformation.LookupParameter("KR_ARM_SETTINGS").AsString()).split("_:_")
		try:
			self.cb_step.Text = self.Settings[0]
			if self.cb_step.Text == "":
				self.cb_step.Text = "100"
		except:
			self.cb_step.Text = self.step_range[0]

		self.btn = Button(Text = "Ок")
		self.btn.Parent = self
		self.btn.Location = Point(230, 29)
		self.btn.Size = Size(150, 23)
		self.btn.Click += self.OnOk

		self.btn_cancel = Button(Text = "Отмена")
		self.btn_cancel.Parent = self
		self.btn_cancel.Location = Point(230, 89)
		self.btn_cancel.Size = Size(150, 23)
		self.btn_cancel.Click += self.OnCancel

		try:
			self.cb_main.Text = value
		except: pass

	def OnOk(self, sender, args):
		self.Close()

	def OnCancel(self, sender, args):
		self.Value = "None"
		self.Close()

	def update_value(self, sender, event):
		self.Value = self.cb_main.Text

	def update_range(self, sender, event):
		try:
			if self.cb_step.Text and self.cb_step.Text != "":
				with db.Transaction(name = "save"):
					doc.ProjectInformation.LookupParameter("KR_ARM_SETTINGS").Set(self.cb_step.Text)
				self.cb_main.Items.Clear()
				self.value_range = []
				for i in range(100, 801):
					if i % int(self.cb_step.Text) == 0:
						self.value_range.append("{}".format(str(i)))
				for i in self.value_range:
					self.cb_main.Items.Add(i)
				self.cb_main.Text = self.value_range[0]
				self.Value = self.value_range[0]
		except: pass

class ElementSelectionFilter (UI.Selection.ISelectionFilter):
	def AllowElement(self, element = DB.Element):
		if (element.Category.Id == DB.ElementId(-2001320) or element.Category.Id == DB.ElementId(-2000032) or element.Category.Id == DB.ElementId(-2000011) or element.Category.Id == DB.ElementId(-2001300) or element.Category.Id == DB.ElementId(-2001330)) and (element.Name.startswith("00") or element.Symbol.FamilyName.startswith("00")):
			return True
		return False
	def AllowReference(self, refer, point):
		return False

def HideElements(mview):
	els = []
	collector_mainview = DB.FilteredElementCollector(doc, mview.Id).WhereElementIsNotElementType().ToElements()
	for element in collector_mainview:
		if not ((element.Category.Id == DB.ElementId(-2001320) or element.Category.Id == DB.ElementId(-2000032) or element.Category.Id == DB.ElementId(-2000011) or element.Category.Id == DB.ElementId(-2001300) or element.Category.Id == DB.ElementId(-2001330)) and element.Name.startswith("00")):
			if element.CanBeHidden(mview):
				els.append(element)
	collection = List[DB.ElementId]([element.Id for element in els])
	with db.Transaction(name = "HideElements"):
		try:
			mview.HideElements(collection)
		except: pass

def set_color(element, view):
	col_value = element.LookupParameter("КР_К_Армирование").AsString()
	try:
		col_value = 800 - int(col_value[:(len(col_value)-8)])
	except:
		col_value = 0
	fillPatternElement = None
	dict = ["Solid fill", "Сплошная заливка", "сплошная заливка", "solid fill"]
	for name in dict:
		try:
			fillPatternElement = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Drafting, name)
			if fillPatternElement != None:
				break
		except: pass
	if fillPatternElement != None:
		error_settings = DB.OverrideGraphicSettings()
		try:
			if col_value <= 200:
				x = col_value/2
				red = 255
				blue = (100 - math.floor(math.fabs(x))) * 255 / 100
				green = math.floor(math.fabs(x)) * 255 / 100
			elif col_value <= 400:
				x = (col_value - 200)/2
				red = (100 - math.floor(math.fabs(x))) * 255 / 100
				blue = 0
				green = 255
			elif col_value <= 600:
				x = (col_value - 400)/2
				red = 0
				blue = math.floor(math.fabs(x)) * 255 / 100
				green = 255
			elif col_value <= 800:
				x = (col_value - 600)/2
				red = 0
				blue = 255
				green = (100 - math.floor(math.fabs(x))) * 255 / 100
			color = DB.Color(red, green, blue)
			settings = DB.OverrideGraphicSettings()
			settings.SetProjectionFillColor(color)
			settings.SetProjectionLineWeight(1)
			settings.SetHalftone(False)
			settings.SetProjectionFillPatternId(fillPatternElement.Id)
			settings.SetSurfaceTransparency(0)
			view.SetElementOverrides(element.Id, settings)
		except Exception as e: pass
#MAIN VIEW
mv_name = "SYS_KR_ARM"
if uidoc.ActiveView.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() != mv_name:
	collector_viewFamily = DB.FilteredElementCollector(doc).OfClass(DB.View3D).WhereElementIsNotElementType()
	bool = False
	for i in collector_viewFamily:
		if i.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() == mv_name:
			zview = i
			bool = True
			uidoc.ActiveView = zview
			HideElements(zview)
	if not bool:
		collector_viewFamilyType = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType()
		with db.Transaction(name = "CreateView"):
			for i in collector_viewFamilyType:
				if i.ViewFamily == DB.ViewFamily.ThreeDimensional:
					viewFamilyType = i
					break
			zview = DB.View3D.CreateIsometric(doc, viewFamilyType.Id)
			zview.get_Parameter(DB.BuiltInParameter.VIEW_NAME).Set(mv_name)
		uidoc.ActiveView = zview
		HideElements(zview)
else:
	zview = uidoc.ActiveView
#CHECK PROJECT PARAMETERS
try:
	group = "КОНСТРУКТИВ - Дополнительные"
	elements_param = "КР_К_Армирование"
	room_params_type = "Text"
	param_found = False
	common_parameters_file = "Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt"
	app = doc.Application
	category_set_elements = app.Create.NewCategorySet()
	insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Walls)
	category_set_elements.Insert(insert_cat_elements)
	insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Floors)
	category_set_elements.Insert(insert_cat_elements)
	insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_StructuralColumns)
	category_set_elements.Insert(insert_cat_elements)
	insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_StructuralFoundation)
	category_set_elements.Insert(insert_cat_elements)
	insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_StructuralFraming)
	category_set_elements.Insert(insert_cat_elements)
	originalFile = app.SharedParametersFilename
	app.SharedParametersFilename = common_parameters_file
	SharedParametersFile = app.OpenSharedParameterFile()
	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	while it.MoveNext():
		d_Definition = it.Key
		d_Name = it.Key.Name
		d_Binding = it.Current
		d_catSet = d_Binding.Categories		
		if d_Name == elements_param:
			if d_Binding.GetType() == DB.InstanceBinding:
				if str(d_Definition.ParameterType) == "Text":
					if d_Definition.VariesAcrossGroups:
						if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Walls)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Floors)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.StructuralColumns)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.StructuralFoundation)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_StructuralFraming)):
							param_found == True
	with db.Transaction(name = "AddSharedParameter"):
		for dg in SharedParametersFile.Groups:
			if dg.Name == group:
				if not param_found:
					externalDefinition = dg.Definitions.get_Item(elements_param)
					newIB = app.Create.NewInstanceBinding(category_set_elements)
					doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
					doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)

	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	with db.Transaction(name = "SetAllowVaryBetweenGroups"):
		while it.MoveNext():
			if not param_found:
				d_Definition = it.Key
				d_Name = it.Key.Name
				d_Binding = it.Current
				if d_Name == elements_param:
					d_Definition.SetAllowVaryBetweenGroups(doc, True)
except: pass
try:
	group = "SYSTEM"
	project_param = "KR_ARM_SETTINGS"
	room_params_type = "Text"
	param_found = False
	common_parameters_file = "Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt"
	category_set_elements = app.Create.NewCategorySet()
	insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_ProjectInformation)
	category_set_elements.Insert(insert_cat_elements)
	originalFile = app.SharedParametersFilename
	app.SharedParametersFilename = common_parameters_file
	SharedParametersFile = app.OpenSharedParameterFile()
	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	while it.MoveNext():
		d_Definition = it.Key
		d_Name = it.Key.Name
		d_Binding = it.Current
		d_catSet = d_Binding.Categories		
		if d_Name == project_param:
			if d_Binding.GetType() == DB.InstanceBinding:
				if str(d_Definition.ParameterType) == "Text":
					if d_Definition.VariesAcrossGroups:
						if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_ProjectInformation)):
							param_found == True
	with db.Transaction(name = "AddSharedParameter"):
		for dg in SharedParametersFile.Groups:
			if dg.Name == group:
				if not param_found:
					externalDefinition = dg.Definitions.get_Item(project_param)
					newIB = app.Create.NewInstanceBinding(category_set_elements)
					doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
					doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
except: pass
#START LOOP
value = "100"
loop = True
try:
	while loop:
		otype = UI.Selection.ObjectType.Element
		filter_element = ElementSelectionFilter()
		reference = uidoc.Selection.PickObject(otype, filter_element, "Выберите элемент для установки армирования")
		element = doc.GetElement(reference)
		form = CreateWindow()
		Application.Run(form)
		value = form.Value
		if value != "None":
			with db.Transaction(name = "WriteValue"):
				element.LookupParameter("КР_К_Армирование").Set(value + "кг/куб.м")
				set_color(element, zview)
except: 
	pass