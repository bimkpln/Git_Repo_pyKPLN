# -*- coding: utf-8 -*-
"""
CheckGrids

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Стены"
__doc__ = 'Проверка расположения стен относительно ближайших осей' \

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

class GridTripple():
	def __init__(self, wall, grid, difference):
		self.Wall = wall
		self.Grid = grid
		self.Difference = difference

class WallApplied():
	def __init__(self, wall, curve, s_p, a_dir, a_dir_p, deg):
		self.Wall = wall
		self.Line = curve
		self.Line_p = DB.Line.CreateBound(DB.XYZ(a_dir_p.X * 1000, a_dir_p.Y * 1000, 0), DB.XYZ(-a_dir_p.X * 1000, -a_dir_p.Y * 1000, 0))
		self.s_p = s_p
		self.a_dir = a_dir
		self.a_dir_p = a_dir_p
		self.a_deg = deg
		self.Grid = self.GetNearestGrid()
		self.Difference = 0.00
		self.GetDifference()


	def GetLength(self, pt_1, pt_2):
		dx = math.fabs(pt_1.X - pt_2.X)
		dy = math.fabs(pt_1.Y - pt_2.Y)
		return math.sqrt(dx*dx+dy*dy)


	def GetDifference(self):
		if self.Grid != False:
			try:
				pt_1 = self.Line_p.Project(self.Grid.Curve.GetEndPoint(0)).XYZPoint
				pt_2 = self.Line_p.Project(self.s_p).XYZPoint
				length = self.GetLength(pt_1, pt_2) * 304.8
				rounded = round(length, 0)
				self.Difference = round(math.fabs(length - rounded), 5)
			except Exception as e:
				print(str(e))
				pass
	def GetNearestGrid(self):
		grids = []
		tuples_l = []
		tuples_l_s = []
		tuples_g = []
		tuples_g_s = []
		for grid in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements():
			dir = grid.Curve.Direction
			deg = math.degrees(dir.AngleTo(DB.XYZ(1, 0, 0)))
			if self.IsEqual(deg, self.a_deg):
				grids.append(grid)

		for grid in grids:
			try:
				pt_1 = self.Line_p.Project(grid.Curve.GetEndPoint(0)).XYZPoint
				pt_2 = self.Line_p.Project(self.s_p).XYZPoint
				length = DB.Line.CreateBound(pt_1, pt_2).Length
				tuples_g.append(grid)
				tuples_l.append(length)
				tuples_l_s.append(length)
			except Exception as e:
				tuples_g.append(grid)
				tuples_l.append(0)
				tuples_l_s.append(0)
		tuples_l_s.sort()
		for sorted_length in tuples_l_s:
			for grid, length in zip(tuples_g, tuples_l):
				if sorted_length == length:
					tuples_g_s.append(grid)
		try:
			return tuples_g_s[0]
		except Exception as e: 			
			return False
		return False

	def IsEqual(self, deg, deg_self):
		if str(deg) == str(deg_self) or str(deg) == str(deg_self +180) or str(deg) == str(deg_self-180):
			return True
		return False

def set_color(ids, warning_ids, views):
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
		settings.SetCutFillPatternId(fillPatternElement.Id)
	for view in views:
		for element in DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType():
			try:
				if element.Id.IntegerValue in ids:
					color = DB.Color(255, 0, 0)
					line_color = DB.Color(255, 0, 0)
					settings.SetHalftone(False)
					settings.SetProjectionLineWeight(6)
					settings.SetCutLineWeight(6)
				elif element.Id.IntegerValue in warning_ids:
					color = DB.Color(255, 168, 0)
					line_color = DB.Color(255, 168, 0)
					settings.SetHalftone(False)
					settings.SetProjectionLineWeight(3)
					settings.SetCutLineWeight(3)
				else:
					color = DB.Color(255, 255, 255)
					line_color = DB.Color(160, 160, 160)
					settings.SetHalftone(True)
					settings.SetProjectionLineWeight(1)
					settings.SetCutLineWeight(1)		
				
				settings.SetSurfaceTransparency(0)				
				settings.SetSurfaceForegroundPatternColor(color)				
				settings.SetProjectionLineColor(line_color)  				
				settings.SetCutForegroundPatternColor(color) 				
				settings.SetCutLineColor(line_color)
				
				view.SetElementOverrides(element.Id, settings)
				
			except:
				pass

def GetView(level):
	for view in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements():
		name = "SYS.В_{}_Проверка стен".format(level.Name)
		try:
			if view.Name == name:
				return view
		except Exception as e:
			return False
	with db.Transaction(name = "CreateView"):
		for view_fam_type in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType():
			if view_fam_type.ViewFamily == DB.ViewFamily.FloorPlan:
				viewFamilyType = view_fam_type
				break
		new_view = DB.ViewPlan.Create(doc, viewFamilyType.Id, level.Id)
		new_view.get_Parameter(DB.BuiltInParameter.VIEW_NAME).Set(name)
		print("Создан проверочный план «{}»".format(name))
	return new_view

def ShowAlert(header, content, footer):
	TD = UI.TaskDialog("KPLN : Проверка модели")
	TD.TitleAutoPrefix = False
	TD.MainInstruction = header
	TD.MainContent = content
	TD.FooterText = footer
	TD.Show()

output = script.get_output()
output.set_title("KPLN : Проверка координат стен")
walls = []
lines = []
Added = []
Tuples = []
ColoredElements = []
ColoredElementsIds = []
ColoredElementsIdsWarning = []
Views = []

max = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements().Count
current = 0
with forms.ProgressBar(title='Проверка стен',indeterminate=False ,cancellable=False, step=1) as pb1:
	for wall in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements():
		current += 1
		pb1.update_progress(current, max_value = max)
		try:
			if type(wall.Location.Curve) == DB.Line:
				s_p = DB.XYZ(wall.Location.Curve.GetEndPoint(0).X, wall.Location.Curve.GetEndPoint(0).Y, 0)
				a_dir = wall.Location.Curve.Direction
				a_dir_p = DB.XYZ(a_dir.Y, -a_dir.X, 0)
				a_deg = math.degrees(a_dir.AngleTo(DB.XYZ(1, 0, 0)))
				walls.append(WallApplied(wall, wall.Location.Curve, s_p, a_dir, a_dir_p, a_deg))
		except: pass
max = len(walls)
current = 0
with forms.ProgressBar(title='Формирование списка',indeterminate=False ,cancellable=False, step=1) as pb2:
	for wall in walls:
		current += 1
		pb2.update_progress(current, max_value = max)
		try:
			if wall.Grid != False:
				if wall.Difference != 0:
					if wall.Difference == 0.5:
						ColoredElementsIdsWarning.append(wall.Wall.Id.IntegerValue)
					else:
						ColoredElementsIds.append(wall.Wall.Id.IntegerValue)
			else:
				print("{}  - стена имеет некорректную привязку к ближайшей оси ({} : <{}>)".format(output.linkify(wall.Wall.Id), str(wall.Wall.WallType.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()), str(wall.Wall.WallType.FamilyName)))
		except Exception as e: print(str(e))
max = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements().Count
current = 0
with forms.ProgressBar(title='Подготовка видов',indeterminate=False ,cancellable=False, step=1) as pb3:
	for level in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements():
		current += 1
		pb3.update_progress(current, max_value = max)
		v = GetView(level)
		if v != False:
			Views.append(GetView(level))
with forms.ProgressBar(title='Применение настроек графики',indeterminate=True ,cancellable=False, step=1) as pb4:
	with db.Transaction(name = "SetColor"):
		try:		
			set_color(ColoredElementsIds, ColoredElementsIdsWarning, Views)
		except Exception as e:
			ShowAlert(":  (", str(e), "")