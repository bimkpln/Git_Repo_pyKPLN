# -*- coding: utf-8 -*-
"""
MarkARM

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Маркировать элементы"
__doc__ = 'С установкой значения расхода стали' \

"""
Архитекурное бюро KPLN

"""
import math
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import revit, DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

def LoadFamilies():
	load_list = ["Z:\\pyRevit\\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\\SYS_KR_ARM_COLUMNS.rfa",
				 "Z:\\pyRevit\\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\\SYS_KR_ARM_CONSTRUCT.rfa",
				 "Z:\\pyRevit\\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\\SYS_KR_ARM_FLOORS.rfa",
				 "Z:\\pyRevit\\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\\SYS_KR_ARM_WALLS.rfa",
				 "Z:\\pyRevit\\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\\SYS_KR_ARM_FOUNDATION.rfa"]
	with db.Transaction(name = "LoadData"):
		for i in load_list:
			try:
				doc.LoadFamily(i)
			except: pass

def Run():
	collector_floor_tags = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_FloorTags)
	for i in collector_floor_tags:
		if i.FamilyName == "SYS_KR_ARM_FLOORS":
			floor_tag_type = i
	collector_wall_tags = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_WallTags)
	for i in collector_wall_tags:
		if i.FamilyName == "SYS_KR_ARM_WALLS":
			wall_tag_type = i
	collector_foundation_tags = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_StructuralFoundationTags)
	for i in collector_foundation_tags:
		if i.FamilyName == "SYS_KR_ARM_FOUNDATION":
			foundation_tag_type = i
	collector_columns_tags = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_StructuralColumnTags)
	for i in collector_columns_tags:
		if i.FamilyName == "SYS_KR_ARM_COLUMNS":
			column_tag_type = i
	collector_framing_tags = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_StructuralFramingTags)
	for i in collector_framing_tags:
		if i.FamilyName == "SYS_KR_ARM_CONSTRUCT":
			framing_tag_type = i
	with db.Transaction(name = "TagElements"):
		for element in selection:
			try:
				options = DB.Options()
				options.ComputeReferences = False
				options.IncludeNonVisibleObjects = False
				geometry = element.get_Geometry(options)
				bbox = geometry.GetBoundingBox()
				max = bbox.Max
				point = DB.XYZ((bbox.Max.X + bbox.Min.X)/2, (bbox.Max.Y + bbox.Min.Y)/2, (bbox.Max.Z + bbox.Min.Z)/2)
				tag = doc.Create.NewTag(view, element, False, DB.TagMode.TM_ADDBY_CATEGORY, DB.TagOrientation.Horizontal, point)
				for type in [floor_tag_type, wall_tag_type, foundation_tag_type, column_tag_type, framing_tag_type]:
					try:
						tag.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM).Set(tagtype.Id)
					except: pass
			except: pass

selection = revit.get_selection()
view = doc.ActiveView

commands = [CommandLink('Да', return_value='Yes'), CommandLink('Отмена', return_value='Cancel')]
dialog = TaskDialog('Маркировка элементов возможна только на заблокированном виде. Заблокировать вид и замаркировать элементы?',
					title = "Маркировка элементов",
					title_prefix=False,
					content="",
					commands=commands,
					footer='',
					show_close=False)

if view.IsLocked:
	try:
		LoadFamilies()
		Run()
	except:
		Alert("Не удалось замаркировать элементы", title="Маркировка", header = "Неизвестная ошибка")
else:
	result = dialog.show()
	if result == "Yes":
		with db.Transaction(name = "LockView"):
			try:
				view.SaveOrientationAndLock()
			except Exception as e:
				Alert("Не удалось заблокировать ориентацию текущего вида! ({})".format(str(e)), title="Маркировка", header = "Неизвестная ошибка")
		try:
			LoadFamilies()
			Run()
		except:
			Alert("Не удалось замаркировать элементы", title="Маркировка", header = "Неизвестная ошибка")