# -*- coding: utf-8 -*-
"""
CheckGrids

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Оси"
__doc__ = 'Проверка расположения осей относительно друг друга' \

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

class GridTuple():
	def __init__(self, grid_1, grid_2, difference):
		self.Grid_1 = grid_1
		self.Grid_2 = grid_2
		self.Difference = difference

class GridApplied():
	def __init__(self, element, curve, s_p, a_dir, a_dir_p, a_deg):
		self.Element = element
		self.Line = curve
		self.Line_p = DB.Line.CreateBound(DB.XYZ(a_dir_p.X * 1000, a_dir_p.Y * 1000, 0), DB.XYZ(-a_dir_p.X * 1000, -a_dir_p.Y * 1000, 0))
		self.s_p = s_p
		self.a_dir = a_dir
		self.a_dir_p = a_dir_p
		self.a_deg = a_deg

	def IsEqual(self, grid):
		if str(grid.a_deg) == str(self.a_deg) or str(grid.a_deg) == str(self.a_deg) or str(grid.a_deg) == str(self.a_deg+180) or str(grid.a_deg) == str(self.a_deg-180):
			return True
		return False

def AddToGroups(grid):
	for group in groups:
		if group[0].IsEqual(grid):
			group.append(grid)
			return True
	groups.append([grid])
	lines.append(grid.Line_p)
	return False

def ShowAlert(header, content, footer):
	TD = UI.TaskDialog("KPLN Координация : Мониторинг")
	TD.TitleAutoPrefix = False
	TD.MainInstruction = header
	TD.MainContent = content
	TD.FooterText = footer
	TD.Show()

output = script.get_output()
output.set_title("KPLN : Проверка координат осей")
grids = []
groups = []
lines = []
Added = []
Tuples = []

for element in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements():
	s_p = DB.XYZ(element.Curve.GetEndPoint(0).X, element.Curve.GetEndPoint(0).Y, 0)
	a_dir = element.Curve.Direction
	a_dir_p = DB.XYZ(a_dir.Y, -a_dir.X, 0)
	a_deg = math.degrees(a_dir.AngleTo(DB.XYZ(1, 0, 0)))
	grids.append(GridApplied(element, element.Curve, s_p, a_dir, a_dir_p, a_deg))

for grid in grids:
	AddToGroups(grid)

for group, line in zip(groups, lines):
	for i in range(0, len(group)):
		for z in range(0, len(group)):
			if i != z:
				pt_1 = line.Project(group[i].s_p).XYZPoint
				pt_2 = line.Project(group[z].s_p).XYZPoint
				try:
					line_res = DB.Line.CreateBound(pt_1, pt_2)
					fact = line_res.Length * 304.8
					rounded = round(line_res.Length * 304.8, 0)
					difference = round(math.fabs(fact - rounded), 5)
					if difference != 0 and str(group[i].Element.Id.IntegerValue) not in Added and str(group[z].Element.Id.IntegerValue) not in Added:
						Tuples.append(GridTuple(group[i].Element, group[z].Element, difference))
						Added.append(str(group[i].Element.Id.IntegerValue))
						continue
					if difference != 0 and str(group[i].Element.Id.IntegerValue) not in Added and str(group[z].Element.Id.IntegerValue) not in Added:
						Added.append(str(group[z].Element.Id.IntegerValue))
						Tuples.append(GridTuple(self, group[i].Element, group[z].Element, difference))
				except Exception as e: print(str(e))
if len(Tuples) != 0:
	for pair in Tuples:
		output.print_md("Погрешность между осями «{}/{}» и «{}/{}» - {}мм;".format(pair.Grid_1.Name, output.linkify(pair.Grid_1.Id), pair.Grid_2.Name, output.linkify(pair.Grid_2.Id), str(pair.Difference)))
else:
	ShowAlert(":  )", "Расположение параллельных осей относительно друг друга в порядке!", "")
