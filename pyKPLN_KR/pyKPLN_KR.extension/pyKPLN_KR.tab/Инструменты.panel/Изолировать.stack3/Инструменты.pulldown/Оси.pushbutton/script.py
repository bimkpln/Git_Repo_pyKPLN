# -*- coding: utf-8 -*-
"""
Fixgrids

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Оси"
__doc__ = 'Выравнивание положения выбранных осей относительно начала координат;\n\n' \
          ' - Значения округляется до 5 мм.\n' \

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

selection = revit.get_selection()
grids = []

for element in selection:
	if element.Category.Name == "Оси":
		grids.append(element)

with db.Transaction(name = "Fix grid location"):
	if len(grids) > 0:
		for grid in grids:
			#gridLocation = grid.Location
			pt_1 = grid.Curve.GetEndPoint(0)
			pt_2 = grid.Curve.GetEndPoint(1)
			StartX = round(pt_1.X * 304.8 / 50, 1) * 50 / 304.8
			StartY = round(pt_1.Y * 304.8 / 50, 1) * 50 / 304.8
			EndX = round(pt_2.X * 304.8 / 50, 1) * 50 / 304.8
			EndY = round(pt_2.Y * 304.8 / 50, 1) * 50 / 304.8
			pt_start = DB.XYZ(StartX, StartY, pt_1.Z)
			pt_end = DB.XYZ(EndX, EndY, pt_2.Z)
			#grid.Curve = grid.Curve.CreateBound(pt_start, pt_end)
			pt_translate = DB.XYZ(StartX - pt_1.X, StartY - pt_1.Y, 0)
			DB.ElementTransformUtils.MoveElement(doc, grid.Id, pt_translate)
	else: forms.alert("Не выбрано ни одного подходящего элемента!")
