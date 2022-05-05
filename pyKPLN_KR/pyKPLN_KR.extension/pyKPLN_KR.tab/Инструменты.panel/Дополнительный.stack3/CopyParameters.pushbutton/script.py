# -*- coding: utf-8 -*-
"""
КопированиеСвойствКР

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Копировать свойства"
__doc__ = 'Присваивает выбранным элементам значения параметров выбранного элемента' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *

class AnySelectorFilter (UI.Selection.ISelectionFilter):
	def AllowElement(self, element = DB.Element):
		return True
	def AllowReference(self, refer, point):
		return False

class ClassicParameter():
	def __init__(self, guid, value):
		self.Value = value
		self.Name = guid

	def TryToSet(self, parameter, element):
		try:
			if not parameter.IsShared:
				element.LookupParameter(self.Name).Set(self.Value)
		except: pass

class SharedParameter():
	def __init__(self, id, value, type):
		self.StorageType = type
		self.Value = value
		self.GUID = id

	def TryToSet(self, parameter):
		try:
			if self.GUID == parameter.GUID:
				parameter.Set(self.Value)
		except: pass

def GetSelected():
	elements = []
	selection = ui.Selection()
	for element in selection:
		try:
			elements.append(element.unwrap())
		except: 
			pass
	return elements

def PickElement():
	element = None
	try:
		pick_filter = AnySelectorFilter()
		ref_element = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element, pick_filter, "KPLN : Элемент - основу")
		try:
			element = doc.GetElement(ref_element)
		except: pass
	except: pass
	return element

c_elements = GetSelected()
if len(c_elements) != 0:
	copy_from = PickElement()
	if copy_from != None:
		values = []
		def_values = []
		for j in copy_from.Parameters:
			if j.IsShared:
				try:
					if j.UserModifiable and j.HasValue:
						if j.StorageType == DB.StorageType.String:
							values.append(SharedParameter(j.GUID, j.AsString(), j.StorageType))
						elif j.StorageType == DB.StorageType.Double:
							values.append(SharedParameter(j.GUID, j.AsDouble(), j.StorageType))
						elif j.StorageType == DB.StorageType.Integer:
							values.append(SharedParameter(j.GUID, j.AsInteger(), j.StorageType))
						elif j.StorageType == DB.StorageType.ElementId:
							values.append(SharedParameter(j.GUID, j.AsElementId(), j.StorageType))
				except: pass
			else:
				try:
					if j.StorageType == DB.StorageType.String:
						def_values.append(ClassicParameter(j.Definition.Name, copy_from.LookupParameter(j.Definition.Name).AsString()))
					elif j.StorageType == DB.StorageType.Double:
						def_values.append(ClassicParameter(j.Definition.Name, copy_from.LookupParameter(j.Definition.Name).AsDouble()))
					elif j.StorageType == DB.StorageType.Integer:
						def_values.append(ClassicParameter(j.Definition.Name, copy_from.LookupParameter(j.Definition.Name).AsInteger()))
					elif j.StorageType == DB.StorageType.ElementId:
						def_values.append(ClassicParameter(j.Definition.Name, copy_from.LookupParameter(j.Definition.Name).AsElementId()))
				except: pass

		with db.Transaction(name = "Create view sections"):
			for element in c_elements:
				for j in element.Parameters:
					for copy_parameter in values:
						copy_parameter.TryToSet(j)
					for copy_parameter in def_values:
						copy_parameter.TryToSet(j, element)
else: forms.alert("Не выбрано ни одного элемента!")