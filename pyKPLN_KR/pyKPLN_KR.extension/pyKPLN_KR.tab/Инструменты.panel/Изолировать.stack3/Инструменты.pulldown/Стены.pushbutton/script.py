# -*- coding: utf-8 -*-
"""
Fixwalls

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Стены"
__doc__ = 'Выравнивание положения выбранных стен относительно начала координат;\n\n' \
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
walls = []

for element in selection:
	if element.Category.Name == "Стены":
		walls.append(element)

with db.Transaction(name = "Fix wall location"):
	if len(walls) > 0:
		for wall in walls:
			wallLocation = wall.Location
			pt_1 = wallLocation.Curve.GetEndPoint(0)
			pt_2 = wallLocation.Curve.GetEndPoint(1)
			StartX = round(pt_1.X * 304.8 / 50, 1) * 50 / 304.8
			StartY = round(pt_1.Y * 304.8 / 50, 1) * 50 / 304.8
			EndX = round(pt_2.X * 304.8 / 50, 1) * 50 / 304.8
			EndY = round(pt_2.Y * 304.8 / 50, 1) * 50 / 304.8
			pt_start = DB.XYZ(StartX, StartY, pt_1.Z)
			pt_end = DB.XYZ(EndX, EndY, pt_2.Z)
			wallLocation.Curve = wallLocation.Curve.CreateBound(pt_start, pt_end)
	else: forms.alert("Не выбрано ни одного подходящего элемента!")
