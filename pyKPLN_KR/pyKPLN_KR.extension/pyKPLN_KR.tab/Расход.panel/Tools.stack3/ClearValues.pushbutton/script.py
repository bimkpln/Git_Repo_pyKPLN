# -*- coding: utf-8 -*-
"""
ClearData

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Отчистить выбранные"
__doc__ = 'Удаление значений расхода стали у выбранных элементов' \
          '-' \

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

def set_color(element, view):
	fillPatternElement = None
	dict = ["Solid fill", "Сплошная заливка", "сплошная заливка", "solid fill"]
	for name in dict:
		try:
			fillPatternElement = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Drafting, name)
			if fillPatternElement != None:
				break
		except: pass
	error_settings = DB.OverrideGraphicSettings()
	settings = DB.OverrideGraphicSettings()	
	settings.SetProjectionLineWeight(1)
	if fillPatternElement != None:
		settings.SetProjectionFillPatternId(fillPatternElement.Id)
	color = DB.Color(150, 150, 150)
	settings.SetSurfaceTransparency(10)
	settings.SetProjectionLineWeight(6)
	settings.SetHalftone(False)
	settings.SetProjectionFillColor(color)
	view.SetElementOverrides(element.Id, settings)


	pass
mv_name = "SYS_KR_ARM"
selection = revit.get_selection()
with db.Transaction(name = "ClearData"):
	for element in selection:
		try:
			element.LookupParameter("КР_К_Армирование").Set("")
			try:
				if doc.ActiveView.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() == mv_name:
					bool = False
					if doc.ActiveView.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() == mv_name:
						set_color(element, doc.ActiveView)
				else: pass
			except: pass
		except: pass
