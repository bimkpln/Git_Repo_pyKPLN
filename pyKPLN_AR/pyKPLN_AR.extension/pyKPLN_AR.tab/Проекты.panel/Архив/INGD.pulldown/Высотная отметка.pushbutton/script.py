# -*- coding: utf-8 -*-
"""
OpeningsHeight

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Высотная отметка"
__doc__ = 'Запись значений высоты проема (относительная и абсолютная)\n' \
		  '      «00_Отметка_Относительная» - высота проема относительно уровня ч.п. 1-го этажа\n' \
		  '      «00_Отметка_Абсолютная» - высота проема относительно ч.п. связанного уровня\n' \

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
from rpw.ui.forms import TextInput, Alert
from rpw.ui.forms import CommandLink, TaskDialog
import datetime

def GetHeight(element):
	values = ["Высота", "INGD_Длина"]
	for j in element.Parameters:
		if j.StorageType == DB.StorageType.Double:
			if j.Definition.Name in values:
				return j.AsDouble()
	return None

def get_description(length_feet):
	comma = "."
	if length_feet >= 0:
		sign = "+"
	else:
		sign = "-"
	length_feet_abs = math.fabs(length_feet)
	length_meters = int(round(length_feet_abs * 304.8 / 5, 0) * 5)
	length_string = str(length_meters)
	if len(length_string) == 7:
		value = length_string[:4] + comma + length_string[4:]
	elif len(length_string) == 6:
		value = length_string[:3] + comma + length_string[3:]
	elif len(length_string) == 5:
		value = length_string[:2] + comma + length_string[2:]
	elif len(length_string) == 4:
		value = length_string[:1] + comma + length_string[1:]
	elif len(length_string) == 3:
		value = "0{}".format(comma) + length_string
	elif len(length_string) == 2:
		value = "0{}0".format(comma) + length_string
	elif len(length_string) == 1:
		value = "0{}00".format(comma) + length_string
	value = sign + value
	return value

collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
params = ["00_Отметка_Относительная", "00_Отметка_Абсолютная"]
params_exist = [False, False]
default_offset_bp = 0.00
#ПРОВЕРКА НАЛИЧИЯ ПАРАМЕТРОВ
try:
	app = doc.Application
	category_set_elements = app.Create.NewCategorySet()
	category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Windows))
	category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Doors))
	category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_MechanicalEquipment))
	originalFile = app.SharedParametersFilename
	app.SharedParametersFilename = "Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\lib\\ФОП_Scripts.txt"
	SharedParametersFile = app.OpenSharedParameterFile()
	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	while it.MoveNext():
		d_Definition = it.Key
		d_Name = it.Key.Name
		d_Binding = it.Current
		d_catSet = d_Binding.Categories	
		for param, bool in zip(params, params_exist):
			if d_Name == param:
				if d_Binding.GetType() == DB.InstanceBinding:
					if str(d_Definition.ParameterType) == "Text":
						if d_Definition.VariesAcrossGroups:
							if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Windows)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Doors)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_MechanicalEquipment)):
								bool == True
	with db.Transaction(name = "AddSharedParameter"):
		for dg in SharedParametersFile.Groups:
			if dg.Name == "АРХИТЕКТУРА - Дополнительные":
				for param, bool in zip(params, params_exist):
					if not bool:
						externalDefinition = dg.Definitions.get_Item(param)
						newIB = app.Create.NewInstanceBinding(category_set_elements)
						doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
						doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)

	map = doc.ParameterBindings
	it = map.ForwardIterator()
	it.Reset()
	with db.Transaction(name = "SetAllowVaryBetweenGroups"):
		while it.MoveNext():
			for param in params:
				d_Definition = it.Key
				d_Name = it.Key.Name
				d_Binding = it.Current
				if d_Name == param:
					d_Definition.SetAllowVaryBetweenGroups(doc, True)
except Exception as e:
	print(str(e))

project_base_point = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().FirstElement()
bp_height = project_base_point.get_BoundingBox(None).Min.Z
#ОТНОСИТЕЛЬНАЯ ОТМЕТКА
with db.Transaction(name='Local height'):
	for element in collector_elements:
		try:
			fam_name = element.Symbol.FamilyName 
			if fam_name.startswith("199_Отверстие в стене прямоугольное") or fam_name == "199_AR_OSW":
				if element.LookupParameter("Расширение границ") != None:
					bound_expand = element.LookupParameter("Расширение границ").AsDouble()
					base_height = element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
					value = get_description(base_height - bound_expand)
					element.LookupParameter("00_Отметка_Относительная").Set(value)
				else:
					base_height = element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
					value = get_description(base_height)
					element.LookupParameter("00_Отметка_Относительная").Set(value)
			if fam_name.startswith("199_Отверстие в стене круглое") or fam_name == "199_AR_ORW":
				if element.LookupParameter("Расширение границ") != None:
					bound_expand = element.LookupParameter("Расширение границ").AsDouble()
					base_height = element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
					element_height = GetHeight(element)/2
					value = get_description(base_height + element_height - bound_expand)
					element.LookupParameter("00_Отметка_Относительная").Set(value)
				else:
					base_height = element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
					element_height = GetHeight(element)/2
					value = get_description(base_height + element_height)
					element.LookupParameter("00_Отметка_Относительная").Set(value)
		except Exception as e:
			print(str(e))

#АБСОЛЮТНАЯ ОТМЕТКА
collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
collector_viewFamilyType = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType()
for i in collector_viewFamilyType:
	if i.ViewFamily == DB.ViewFamily.ThreeDimensional:
		viewFamilyType = i
		break
with db.Transaction(name='Global height'):
	for element in collector_elements:
		try:
			fam_name = element.Symbol.FamilyName 
			if fam_name.startswith("199_Отверстие в стене прямоугольное") or fam_name == "199_AR_OSW":
				down = element.LookupParameter("offset_down").AsDouble()
				b_box = element.get_BoundingBox(None)
				boundingBox_Z_min = b_box.Min.Z + down - default_offset_bp
				value = get_description(boundingBox_Z_min - bp_height)
				element.LookupParameter("00_Отметка_Абсолютная").Set(value)
			if fam_name.startswith("199_Отверстие в стене круглое"):
				down = element.LookupParameter("offset_down").AsDouble()
				up = element.LookupParameter("offset_up").AsDouble()
				b_box = element.get_BoundingBox(None)
				boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
				value = get_description(boundingBox_Z_center - bp_height)
				element.LookupParameter("00_Отметка_Абсолютная").Set(value)
			if fam_name.startswith("501_Гильза_АР_Стена") or fam_name == "199_AR_ORW":#temp20200515
				down = element.LookupParameter("offset_down").AsDouble()
				up = element.LookupParameter("offset_up").AsDouble()
				b_box = element.get_BoundingBox(None)
				boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
				value = get_description(boundingBox_Z_center - bp_height)
				element.LookupParameter("00_Отметка_Абсолютная").Set(value)
		except:
			try:
				fam_name = element.Symbol.FamilyName 
				if fam_name.startswith("199_Отверстие в стене прямоугольное") or fam_name == "199_AR_OSW":
					down = element.LookupParameter("SYS_OFFSET_DOWN").AsDouble()
					b_box = element.get_BoundingBox(None)
					boundingBox_Z_min = b_box.Min.Z + down - default_offset_bp
					value = get_description(boundingBox_Z_min - bp_height)
					element.LookupParameter("00_Отметка_Абсолютная").Set(value)
				if fam_name.startswith("199_Отверстие в стене круглое"):
					down = element.LookupParameter("SYS_OFFSET_DOWN").AsDouble()
					up = element.LookupParameter("SYS_OFFSET_UP").AsDouble()
					b_box = element.get_BoundingBox(None)
					boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
					value = get_description(boundingBox_Z_center - bp_height)
					element.LookupParameter("00_Отметка_Абсолютная").Set(value)
				if fam_name.startswith("501_Гильза_АР_Стена") or fam_name == "199_AR_ORW":
					down = element.LookupParameter("SYS_OFFSET_DOWN").AsDouble()
					up = element.LookupParameter("SYS_OFFSET_UP").AsDouble()
					b_box = element.get_BoundingBox(None)
					boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
					value = get_description(boundingBox_Z_center - bp_height)
					element.LookupParameter("00_Отметка_Абсолютная").Set(value)
			except:
				pass