# -*- coding: utf-8 -*-
"""
INGD_Level

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "INGD_Этаж"
__doc__ = 'Заполнение параметра INGD_Номер этажа по значению из связанного уровня для всех элементов проекта НА АКТИВНОМ ВИДЕ' \
"""
"""
from pyrevit.framework import clr
import re
import os
from rpw import doc, uidoc, DB, UI, db, ui
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction
from System.Collections.Generic import *
from System.Windows.Forms import *
from System.Drawing import *
import System
from System import Guid
import datetime
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

def GetLevelParameter(el, par):
	bp = None
	try: bp = doc.GetElement(el.LevelId).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(el.Host.LevelId).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(el.SuperComponent.LevelId).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(el.SuperComponent.Host.LevelId).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(el.SuperComponent).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(el.SuperComponent.Host).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(doc.GetElement(el.GroupId).LevelId).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(doc.GetElement(el.Host.GroupId).LevelId).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(doc.GetElement(el.SuperComponent.GroupId).LevelId).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(doc.GetElement(el.SuperComponent.Host.GroupId).LevelId).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(el.get_Parameter(DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId()).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(doc.GetElement(el.HostId).get_Parameter(DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId()).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(el.get_Parameter(DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId()).LookupParameter(par)
	except : pass
	if bp != None: return bp
	#
	try: bp = doc.GetElement(el.Host).LookupParameter(par)
	except : pass
	if bp != None: return bp
	return bp

with db.Transaction(name = "ING"):
	for element in DB.FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements():
		if element != None:
			try:
				for p in ["INGD_Номер этажа"]:
					bp = GetLevelParameter(element, p)
					if bp != None:
						par = element.LookupParameter(p)
						if par.CanBeAssociatedWithGlobalParameters():
							par.Set(bp.AsString())
			except : pass