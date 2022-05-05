# -*- coding: utf-8 -*-
"""
KPLN:DIV:MASSSETTER

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Выбрать + Заполнить (Форма)"
__doc__ = 'Заполнить параметры по пересечению с выбранной формообразующей' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

import wpf
from System.Windows import Application, Window
from System.Collections.ObjectModel import ObservableCollection

import re
import System
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert
import datetime
import webbrowser

class WPFCategory():
	def __init__(self, category):
		self.IsChecked = False
		self.Category = category

class MassSelectorFilter (UI.Selection.ISelectionFilter):
	def AllowElement(self, element = DB.Element):
		if element.Category.Id.IntegerValue == -2003400:
			return True
		else:
			return False
	def AllowReference(self, refer, point):
		return False

class CategoryForm(Window):
	def __init__(self):
		wpf.LoadComponent(self, 'Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\ДИВ.pulldown\\Выбиратор (формы).pushbutton\\PickCategories.xaml')
		self.Categories = []
		for c in doc.Settings.Categories:
			if not c.IsTagCategory and c.CategoryType == DB.CategoryType.Model:
				self.Categories.append(WPFCategory(c))
		self.Categories.sort(key=lambda x: x.Category.Name, reverse=False)
		self.lb.ItemsSource = self.Categories

	def OnBtnApply(self, sender, e):
		global cats
		for i in self.lb.ItemsSource:
			if i.IsChecked:
				cats.append(i.Category)
		self.Close()

	def UnChecked(self, sender, e):
		for i in self.lb.SelectedItems:
			i.IsChecked = False
		System.Windows.Data.CollectionViewSource.GetDefaultView(self.lb.ItemsSource).Refresh()

	def Checked(self, sender, e):
		for i in self.lb.SelectedItems:
			i.IsChecked = True
		System.Windows.Data.CollectionViewSource.GetDefaultView(self.lb.ItemsSource).Refresh()

class ParametersForm(Window):
	def __init__(self, parameters):
		wpf.LoadComponent(self, 'Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\ДИВ.pulldown\\Выбиратор (формы).pushbutton\\Parameters.xaml')
		self.cbxParameters.ItemsSource = parameters

	def OnBtnApply(self, sender, e):
		global pickedParameter
		global pickedValue
		pickedParameter = self.cbxParameters.SelectedItem
		pickedValue = self.tbxValue.Text
		self.Close()

def PickElement():
	element = None
	try:
		pick_filter = MassSelectorFilter()
		ref_element = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element, pick_filter, "Выберите элемент : Формы")
		try:
			element = doc.GetElement(ref_element)
		except: pass
	except:
	    pass
	return element

def GetGeometry(element):
	solids = []
	if element.Category.Id.IntegerValue == -2000160:
		SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
		calculator = DB.SpatialElementGeometryCalculator(doc, DB.SpatialElementBoundaryOptions())
		results = calculator.CalculateSpatialElementGeometry(element)
		room_solid = results.GetGeometry()
		if room_solid != None:
			solids.append(room_solid)
	else:
		options = DB.Options()
		options.DetailLevel = DB.ViewDetailLevel.Fine
		options.IncludeNonVisibleObjects = True
		for i in element.get_Geometry(options):
			try:
				if type(i) == DB.Solid:
					if i.SurfaceArea != 0:
						solids.append(i)
			except: pass
			try:
				for g in i.GetInstanceGeometry():
					if type(g) == DB.Solid:
						if g.SurfaceArea != 0:
							solids.append(g)
			except: pass
	return solids

def boolean_intersection(a, b):
	try:
		s_intersect = DB.BooleanOperationsUtils.ExecuteBooleanOperation(a, b, DB.BooleanOperationsType.Intersect)
		if abs(s_intersect.Volume) > 0.0000001:
			return True
		else:
			s_union = DB.BooleanOperationsUtils.ExecuteBooleanOperation(a, b, DB.BooleanOperationsType.Union)
			area = abs(a.SurfaceArea + b.SurfaceArea - s_union.SurfaceArea)
			if (area < 0.0000001) or (a.Edges.Size + b.Edges.Size == s_union.Edges.Size):
				return False
			else:
				s_difference = DB.BooleanOperationsUtils.ExecuteBooleanOperation(a, b, DB.BooleanOperationsType.Difference)
				area = abs(a.SurfaceArea + b.SurfaceArea - s_difference.SurfaceArea)
				if (area < 0.0000001) and (a.Edges.Size + b.Edges.Size == s_difference.Edges.Size):
					return False
				else:
					return True
	except: return False

cats = []
form = PickElement()
if form != None:
	CategoryForm().ShowDialog()
	params = []
	hash = []
	elements = []
	for c in cats:
		for i in DB.FilteredElementCollector(doc).OfCategoryId(c.Id).WhereElementIsNotElementType().ToElements():
			for j in i.Parameters:
				if j.StorageType == DB.StorageType.String and not j.Definition.Name in hash:
					params.append(j)
					hash.append(j.Definition.Name)
			defined = False
			for a in GetGeometry(i):
				if defined: break
				for b in GetGeometry(form):
					if defined: break
					if boolean_intersection(a, b):
						elements.append(i)
						defined = True
						break


	params.sort(key=lambda x: x.Definition.Name, reverse=False)
	pickedParameter = None
	pickedValue = None
	ParametersForm(params).ShowDialog()
	if pickedParameter != None and pickedValue != None:
		with db.Transaction(name = "Запись в параметр"):
			for i in elements:
				try:
					for j in i.Parameters:
						if j.IsShared == pickedParameter.IsShared and j.Definition.Name == pickedParameter.Definition.Name and j.StorageType == DB.StorageType.String:
							j.Set(pickedValue)
				except :
					pass