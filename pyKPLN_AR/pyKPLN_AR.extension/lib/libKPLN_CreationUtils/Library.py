# -*- coding: utf-8 -*-
import math
import time
import datetime
import os
import re

from rpw import doc, uidoc, DB, UI, db, ui, revit as Revit
from pyrevit import script
from pyrevit import forms
from pyrevit import revit, DB, UI
from pyrevit.framework import clr
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *

import System
from System.Windows.Forms import *
from System.Drawing import *
from System import Guid

class RoomSelectorFilter (UI.Selection.ISelectionFilter):
	def AllowElement(self, element = DB.Element):
		if element.Category.Id == DB.ElementId(-2000160):
			return True
		return False
	def AllowReference(self, refer, point):
		return False

class KPLN_Tuple_CreateCornerMark():
	def __init__(self, edge_1, edge_2, symbol, view, char):
		self.Edge_1 = edge_1
		self.Edge_2 = edge_2
		try:
			self.Point = self.GetIntersectPoint()
			self.Dir = (self.Edge_1.Direction.Normalize() + self.Edge_2.Direction.Normalize())/2
			self.Dir = self.Dir.Normalize().Negate()
			self.Angle = -math.atan2(self.Dir.Y, self.Dir.X) / 0.01745329251994
			if self.Angle < 0:
				self.Angle += 360
			annotation = doc.Create.NewFamilyInstance(self.Point, symbol, view)
			n = 0
			if self.Angle < 22.5 or self.Angle > 337.5:
				n = 0
			elif self.Angle <= 67.5 and self.Angle >= 22.5:
				n = 1
			elif self.Angle < 112.5 and self.Angle > 67.5:
				n = 2
			elif self.Angle <= 157.5 and self.Angle >= 112.5:
				n = 3
			elif self.Angle < 202.5 and self.Angle > 157.5:
				n = 4
			elif self.Angle <= 247.5 and self.Angle >= 202.5:
				n = 5
			elif self.Angle < 292.5 and self.Angle > 247.5:
				n = 6
			elif self.Angle <= 337.5 and self.Angle >= 292.5:
				n = 7
			annotation.LookupParameter("SYS_ANGLE").Set(n)
			annotation.LookupParameter("Значение").Set(char)
		except Exception as e: print("KPLN_Tuple_CreateCornerMark: {}".format(str(e)))

	def GetIntersectPoint(self):
		try:
			pts_1 = [self.Edge_1.GetEndPoint(0), self.Edge_1.GetEndPoint(1)]
			pts_2 = [self.Edge_2.GetEndPoint(0), self.Edge_2.GetEndPoint(1)]
			for point_1 in pts_1:
				for point_2 in pts_2:
					pt = pts_2[0]
					if str(point_1) ==  str(point_2):
						return point_1
			return pt
		except Exception as e: print("KPLN_Tuple_CreateCornerMark : GetIntersectPoint: {}".format(str(e)))
		return None

class KPLN_Flat():
	def __init__(self, rooms):
		self.Rooms = rooms
		self.Edges = []
		for room in self.Rooms:
			self.Edges.append(room.OptimizedLoop)
		self.CommonLoop = self.GetCommonLoop(self.Edges)
		self.DrawLines(self.CommonLoop)

	def GetIntersectPointProtected(self, ln_1, ln_2):
		pt_1 = ln_1.GetEndPoint(1)
		pt_2 = ln_2.GetEndPoint(0)
		if str(pt_1) == str(pt_2):
			return pt_2
		elif self.PointOnLine(pt_1, ln_2):
			return pt_2
		elif self.PointOnLine(pt_2, ln_1):
			return pt_1
		max_x = max([pt_1.X, pt_2.X])
		min_x = min([pt_1.X, pt_2.X])
		max_y = max([pt_1.Y, pt_2.Y])
		min_y = min([pt_1.Y, pt_2.Y])
		x = math.fabs(max_x - min_x)
		y = math.fabs(max_y - min_y)
		distance = math.sqrt(x*x + y*y)
		if distance <= 0.00:
			return pt_2
		d_dir = pt_2 - pt_1
		d_dir.Normalize()
		if str(d_dir) == str(ln_1.Direction) or str(d_dir.Negate()) == str(ln_1.Direction):
			return pt_1
		if str(d_dir) == str(ln_2.Direction) or str(d_dir.Negate()) == str(ln_2.Direction):
			return pt_2
		angle_b = ln_1.Direction.AngleTo(ln_2.Direction)
		angle_y = ln_1.Direction.AngleTo(d_dir)
		add_range = round(distance, 10) * round((math.sin(angle_b + angle_y)), 10) / round(math.sin(angle_b), 10)
		multiplier = DB.XYZ(ln_1.Direction.X * add_range, ln_1.Direction.Y * add_range, 0.0)
		intersection_point = pt_1 + multiplier
		return intersection_point

	def PointOnLine(self, pt, ln):
		try:
			dir_1 = ln.Direction
			dir_2 = DB.Line.CreateBound(ln.GetEndPoint(0), pt).Direction
			dir_3 = DB.Line.CreateBound(pt, ln.GetEndPoint(1)).Direction
			if str(dir_1) == str(dir_2) or str(dir_2) == str(dir_3):
				return True
			return False
		except Exception as e:
			return True

	def DrawLines(self, loop):
		with db.Transaction(name = "DrawLines"):
			for i in loop:
				try:
					doc.Create.NewDetailCurve(doc.ActiveView, i)
				except: pass

	def GetCommonLoop(self, loop_list):
		common_loop = [loop_list[0][0]]
		for loop in loop_list:
			current_line = common_loop[-1]
			for i in range(0, len(loop)):
				line = loop[i]
				if i != len(loop)-1:
					next_line = loop[i+1]
				else:
					next_line = loop[0]
				angle = self.GetAngle(current_line, line)
				if angle < 0 and math.fabs(angle) > 30:
					common_loop.append(DB.Line.CreateBound(current_line.GetEndPoint(0), line.GetEndPoint(1)))
				else:
					pass
		return common_loop

	def GetAngle(self, line_1, line_2):
		dir_1 = math.degrees(line_1.Direction.AngleTo(DB.XYZ(0,1,0)))
		dir_2 = math.degrees(line_2.Direction.AngleTo(DB.XYZ(0,1,0)))
		return dir_2 - dir_1

	def GetClockwiseLoop(self, previous_line, loop, exist):
		lines = []
		next_line = None
		lines_distance = []
		lines_distance_sorted = []
		dir_1 = math.degrees(previous_line.Direction.AngleTo(DB.XYZ(0,1,0)))
		for i in range(0, len(loop)):
			loop_line = loop[i]
			if str(loop_line.GetEndPoint(0)) == str(previous_line.GetEndPoint(0)) and str(loop_line.GetEndPoint(1)) == str(previous_line.GetEndPoint(1)):
				if i != len(loop)-1:
					next_line = loop[i+1]
				else:
					next_line = loop[0]
			elif str(loop_line.GetEndPoint(0)) != str(previous_line.GetEndPoint(1)) and str(loop_line.GetEndPoint(1)) != str(previous_line.GetEndPoint(1)):
				dir_2 = math.degrees(loop_line.Direction.AngleTo(DB.XYZ(0,1,0)))
				angle = dir_2 - dir_1
				if angle < 0 and math.fabs(angle) > 15:
					try:
						lines.append(DB.Line.CreateBound(previous_line.GetEndPoint(0), self.GetIntersectPointProtected(previous_line, loop_line)))
						distance = loop_line.Distance(previous_line.GetEndPoint(1))
						lines_distance.append(distance)
						lines_distance_sorted.append(distance)
					except: pass
		lines_distance_sorted.sort()
		for i in range(0, len(lines_distance)):
			if lines_distance[i] == lines_distance_sorted[0]:
				if not self.InList(lines[i], exist):
					return lines[i]
		return next_line

	def InList(self, line, list):
		for i in list:
			if str(i.GetEndPoint(0)) == str(line.GetEndPoint(0)) and str(i.GetEndPoint(1)) == str(line.GetEndPoint(1)):
				return True
		return False

class KPLN_Room():
	def __init__(self, element = DB.Element):
		self.Room = element
		self.OptimizedLoop = self.GetOptimizedEdges()
		self.FloopPlan = None
		self.FloorSheet = None
		self.SectionViews = []
		self.SectionViewsDict = []
		self.SectionSheet = None
		self.Dict = ["А", "Б", "В", "Г", "Д", "Е", "Ж", "З", "И", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Э", "Ю", "Я",
			   "АА", "ББ", "ВВ", "ГГ", "ДД", "ЕЕ", "ЖЖ", "ЗЗ", "ИИ", "КК", "ЛЛ", "ММ", "НН", "ОО", "ПП", "РР", "СС", "ТТ", "УУ", "ФФ", "ХХ", "ЦЦ", "ЧЧ", "ШШ", "ЩЩ", "ЭЭ", "ЮЮ", "ЯЯ"]

	def	BaseLevelHeight(self):
		elevation = doc.GetElement(self.Room.LevelId).Elevation - 200 / 304.8
		return elevation

	def	UpperLevelHeight(self):
		elevation = doc.GetElement(self.Room.LevelId).Elevation + self.Room.get_Parameter(DB.BuiltInParameter.ROOM_HEIGHT).AsDouble() + 200 / 304.8
		return elevation

	def CreateSectionsAndSheets(self, view_template = DB.Element, sheet_template = DB.Element):
		print("\nПОМЕЩЕНИЕ № {}".format(self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()))
		try:
			#GET FAMILY TYPE
			for view_fam_type in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType():
				if view_fam_type.ViewFamily == DB.ViewFamily.Section:
					viewFamilyType = view_fam_type
			#GET Z BOUNDS
			minZ = self.BaseLevelHeight()
			maxZ = self.UpperLevelHeight()
			#CREATE SECTIONS
			self.SectionViews = []
			for i in range(0, len(self.OptimizedLoop)):
				line = self.OptimizedLoop[i]
				view = self.CreateSection(line, viewFamilyType, minZ, maxZ, view_template)
				view.ViewTemplateId = DB.ElementId(-1)
				view.Scale = view_template.Scale
				self.ResetGraphics(view)
				self.SectionViews.append(view)
				#SET NAME
				attempt = 0
				char_1 = self.Dict[i]
				if i != len(self.OptimizedLoop)-1:
					char_2 = self.Dict[i+1]
				else:
					char_2 = self.Dict[0]
				self.SectionViewsDict.append([char_1, char_2])
				name = "АР.О_ЧП({})_Развертка {}-{} №{}".format(self.GetElevationDescription(doc.GetElement(self.Room.LevelId).Elevation), char_1, char_2, self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
				if self.ViewNameIsUniq(name):
					view.get_Parameter(DB.BuiltInParameter.VIEW_NAME).Set(name)
					doc.Regenerate()
				else:
					while True:
						attempt += 1
						name = "АР.О_ЧП({})_Развертка {}-{} №{} ({})".format(self.GetElevationDescription(doc.GetElement(self.Room.LevelId).Elevation), char_1, char_2, self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString(), str(attempt))					
						if self.ViewNameIsUniq(name):
							view.get_Parameter(DB.BuiltInParameter.VIEW_NAME).Set(name)
							doc.Regenerate()
							break
				print("\tСоздан разрез: «{}»".format(name))
				#CROP VIEW
				pt_1 = DB.XYZ(line.GetEndPoint(0).X, line.GetEndPoint(0).Y, minZ)
				pt_2 = DB.XYZ(line.GetEndPoint(1).X, line.GetEndPoint(1).Y, minZ)
				pt_3 = DB.XYZ(line.GetEndPoint(0).X, line.GetEndPoint(0).Y, maxZ)
				pt_4 = DB.XYZ(line.GetEndPoint(1).X, line.GetEndPoint(1).Y, maxZ)
				plane = view.GetCropRegionShapeManager().GetCropShape()[0].GetPlane()
				pt_pj_1 = pt_1 + self.GetDistanceToPlane(plane, pt_1) * plane.Normal
				pt_pj_2 = pt_2 + self.GetDistanceToPlane(plane, pt_2) * plane.Normal
				pt_pj_3 = pt_3 + self.GetDistanceToPlane(plane, pt_3) * plane.Normal
				pt_pj_4 = pt_4 + self.GetDistanceToPlane(plane, pt_4) * plane.Normal
				line_1 = DB.Line.CreateBound(pt_pj_1, pt_pj_2)
				line_2 = DB.Line.CreateBound(pt_pj_2, pt_pj_4)
				line_3 = DB.Line.CreateBound(pt_pj_4, pt_pj_3)
				line_4 = DB.Line.CreateBound(pt_pj_3, pt_pj_1)
				loop = DB.CurveLoop()
				loop.Append(line_1)
				loop.Append(line_2)
				loop.Append(line_3)
				loop.Append(line_4)
				view.GetCropRegionShapeManager().SetCropShape(loop)
			doc.Regenerate()
			#CREATE SHEET
			self.SectionSheet = DB.ViewSheet.Create(doc, sheet_template.Id)		
			attempt = 0
			name = "АР.О_ЧП({})_Развертки по помещению «{}» №{}".format(self.GetElevationDescription(doc.GetElement(self.Room.LevelId).Elevation), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString(), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
			if self.SheetNameIsUniq(name):
				self.SectionSheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).Set(name)
			else:
				while True:
					attempt += 1
					name = "АР.О_ЧП({})_Развертки по помещению «{}» №{} ({})".format(self.GetElevationDescription(doc.GetElement(self.Room.LevelId).Elevation), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString(), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString(), str(attempt))
					if self.SheetNameIsUniq(name):
						self.SectionSheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).Set(name)
						break
			print("\tСоздан лист: «{}»".format(name))
			doc.Regenerate()
			#PLACE VIEWS ON SHEETS
			width = math.fabs(self.SectionSheet.Outline.Max.U - self.SectionSheet.Outline.Min.U)
			height = math.fabs(self.SectionSheet.Outline.Max.V - self.SectionSheet.Outline.Min.V)
			view_group = []
			view_group_char = []
			view_groups = []
			view_groups_chars = []
			view_group_width = 0
			for view, chars in zip(self.SectionViews, self.SectionViewsDict):
				view_group_width += math.fabs(view.Outline.Max.U - view.Outline.Min.U)
				if view_group_width < width - 0.2:
					view_group.append(view)
					view_group_char.append(chars)
				elif len(view_group) != 0:
					view_group_width = 0
					view_groups.append(view_group)
					view_groups_chars.append(view_group_char)
					view_group = []
					view_group_char = []
					view_group.append(view)
					view_group_char.append(chars)
				else:
					view_group.append(view)
					view_group_char.append(chars)
			if len(view_group) != 0:
				view_groups.append(view_group)
				view_groups_chars.append(view_group_char)
			number_of_groups = len(view_groups)
			num = 0
			group_number = 0
			width = math.fabs(self.SectionSheet.Outline.Max.U - self.SectionSheet.Outline.Min.U)
			height = math.fabs(self.SectionSheet.Outline.Max.V - self.SectionSheet.Outline.Min.V)
			for group, char_g in zip(view_groups, view_groups_chars):
				group_number += 1
				num += 1
				position = self.SectionSheet.Outline.Min.V + height / (number_of_groups + 1) * (number_of_groups + 1 - num)
				sheet_V_center = (self.SectionSheet.Outline.Min.V + self.SectionSheet.Outline.Max.V)/2
				sheet_U_center = (self.SectionSheet.Outline.Min.U + self.SectionSheet.Outline.Max.U)/2
				views_U = 0.0
				views_V = 0.0
				for view in group:
					views_U += math.fabs(view.Outline.Max.U - view.Outline.Min.U)
					views_V += math.fabs(view.Outline.Max.V - view.Outline.Min.V)
				pl_U = sheet_U_center - views_U/2
				pl_V = sheet_V_center - views_V/2
				for i in range(0, len(group)):
					view = group[i]
					annotation = doc.Create.NewFamilyInstance(DB.XYZ(pl_U, position + math.fabs(view.Outline.Max.V - view.Outline.Min.V)/2,0), self.GetAnnotationSymbol_Axis(), self.SectionSheet)
					annotation.LookupParameter("Высота").Set(math.fabs(view.Outline.Max.V - view.Outline.Min.V))
					annotation.LookupParameter("Значение").Set(char_g[i][0])
					pl_U += math.fabs(view.Outline.Max.U - view.Outline.Min.U)/2
					DB.Viewport.Create(doc, self.SectionSheet.Id, view.Id, DB.XYZ(pl_U, position, 0))		
					pl_U += math.fabs(view.Outline.Max.U - view.Outline.Min.U)/2
					if i == len(group)-1:
						annotation = doc.Create.NewFamilyInstance(DB.XYZ(pl_U,position + math.fabs(view.Outline.Max.V - view.Outline.Min.V)/2,0), self.GetAnnotationSymbol_Axis(), self.SectionSheet)
						annotation.LookupParameter("Высота").Set(math.fabs(view.Outline.Max.V - view.Outline.Min.V))
						annotation.LookupParameter("Значение").Set(char_g[i][1])
					view.ViewTemplateId = view_template.Id
					doc.Regenerate()
					view.get_Parameter(DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER).Set("{}-{}".format(char_g[i][0], char_g[i][1]))
		except Exception as e: print("CreateSectionsAndSheets: " + str(e))

	def GetDistanceToPlane(self, plane, point):
		line_to_origin = DB.Line.CreateBound(point, plane.Origin)
		distance_to_origin = line_to_origin.Length
		angle_a = line_to_origin.Direction.AngleTo(plane.Normal)
		distance = math.fabs(distance_to_origin * math.cos(angle_a))
		return distance

	def CreateSection(self, line, fam_type, minZ, maxZ, view_template):
		converted_line = line.CreateReversed()
		p = converted_line.GetEndPoint(0)
		q = converted_line.GetEndPoint(1)
		v = q - p

		w = v.GetLength()
		h = maxZ - minZ
		d = 0.1
		offset = 0.1 * w

		min = DB.XYZ( -w, minZ - offset, -offset )
		max = DB.XYZ( w, maxZ + offset, 0 )

		midpoint = p + 0.5 * v
		walldir = v.Normalize()
		up = DB.XYZ.BasisZ
		viewdir = walldir.CrossProduct(up)
 
		t = DB.Transform.Identity
		t.Origin = midpoint
		t.BasisX = walldir
		t.BasisY = up
		t.BasisZ = viewdir

		sectionBox = DB.BoundingBoxXYZ()
		sectionBox.Transform = t
		#
		sectionBox.Min = min
		sectionBox.Max = max
		view_section = DB.ViewSection.CreateSection(doc, fam_type.Id, sectionBox)
		#
		view_section.get_Parameter(DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR).Set(5)
		view_section.Scale = view_template.Scale
		doc.Regenerate()
		return view_section

	def CreateFloorPlanAndSheet(self, view_template = DB.Element, sheet_template = DB.Element):
		print("\nПОМЕЩЕНИЕ № {}".format(self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()))
		#GET FAMILY TYPE
		for view_fam_type in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType():
			if view_fam_type.ViewFamily == DB.ViewFamily.FloorPlan:
				viewFamilyType = view_fam_type
		#CREATE FLOOR PLAN
		self.FloopPlan = DB.ViewPlan.Create(doc, viewFamilyType.Id, self.Room.LevelId)
		doc.Regenerate()
		#SET VIEW NAME
		attempt = 1
		name = "АР.О_ЧП({})_Разверточный план «{}» №{}".format(self.GetElevationDescription(doc.GetElement(self.Room.LevelId).Elevation), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString(), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
		if self.ViewNameIsUniq(name):
			self.FloopPlan.get_Parameter(DB.BuiltInParameter.VIEW_NAME).Set(name)
			doc.Regenerate()
		else:
			while True:
				name = "АР.О_ЧП({})_Разверточный план «{}» №{} ({})".format(self.GetElevationDescription(doc.GetElement(self.Room.LevelId).Elevation), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString(), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString(), str(attempt))
				if self.ViewNameIsUniq(name):
					self.FloopPlan.get_Parameter(DB.BuiltInParameter.VIEW_NAME).Set(name)
					doc.Regenerate()
					break
				attempt += 1
		print("\tСоздан план: «{}»".format(name))
		#SET TEMPLATE
		self.FloopPlan.ViewTemplateId = view_template.Id
		#SET CROP BOX
		self.FloopPlan.CropBoxActive = True
		try:
			new_crop_shape = self.GetExpandedLoop(self.OptimizedLoop, 1.2)
			self.FloopPlan.GetCropRegionShapeManager().SetCropShape(new_crop_shape)
		except:
			try:
				new_crop_shape = self.GetExpandedLoop(self.OptimizedLoop, 0.6)
				self.FloopPlan.GetCropRegionShapeManager().SetCropShape(new_crop_shape)
			except:
				try:
					new_crop_shape = self.GetExpandedLoop(self.OptimizedLoop, 0.4)
					self.FloopPlan.GetCropRegionShapeManager().SetCropShape(new_crop_shape)
				except:
					try:
						new_crop_shape = self.GetExpandedLoop(self.OptimizedLoop, 0.2)
						self.FloopPlan.GetCropRegionShapeManager().SetCropShape(new_crop_shape)
					except:
						new_crop_shape = self.GetExpandedLoop(self.OptimizedLoop, 0.0001)
						self.FloopPlan.GetCropRegionShapeManager().SetCropShape(new_crop_shape)
		self.FloopPlan.CropBoxVisible = False
		for i in range(0, len(self.OptimizedLoop)):
			if i != len(self.OptimizedLoop)-1:
				KPLN_Tuple_CreateCornerMark(self.OptimizedLoop[i], self.OptimizedLoop[i+1], self.GetAnnotationSymbol_Rose(), self.FloopPlan, self.Dict[i+1])
			else:
				KPLN_Tuple_CreateCornerMark(self.OptimizedLoop[i], self.OptimizedLoop[0], self.GetAnnotationSymbol_Rose(), self.FloopPlan, self.Dict[0])
		#CREATE SHEET
		doc.Regenerate()
		self.FloorSheet = DB.ViewSheet.Create(doc, sheet_template.Id)
		#SET SHEET NAME
		attempt = 1
		name = "АР.О_ЧП({})_Разверточный план помещения «{}» №{}".format(self.GetElevationDescription(doc.GetElement(self.Room.LevelId).Elevation), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString(), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
		if self.SheetNameIsUniq(name):
			self.FloorSheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).Set(name)
			doc.Regenerate()
		else:
			while True:
				attempt += 1
				name = "АР.О_ЧП({})_Разверточный план помещения «{}» №{} ({})".format(self.GetElevationDescription(doc.GetElement(self.Room.LevelId).Elevation), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString(), self.Room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString(), str(attempt))
				if self.SheetNameIsUniq(name):
					self.FloorSheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).Set(name)
					doc.Regenerate()
					break
		print("\tСоздан лист: «{}»".format(name))
		#PLACE FLOOR PLAN ON SHEET
		doc.Regenerate()
		pl_V = (self.FloorSheet.Outline.Min.V + self.FloorSheet.Outline.Max.V)/2
		pl_U = (self.FloorSheet.Outline.Min.U + self.FloorSheet.Outline.Max.U)/2
		DB.Viewport.Create(doc, self.FloorSheet.Id, self.FloopPlan.Id, DB.XYZ(pl_U, pl_V, 0))
		doc.Regenerate()

	def ResetGraphics(self, view):
		for cat in view.Document.Settings.Categories:
			if cat.get_AllowsVisibilityControl(view):
				cat.set_Visible(view, False)
		doc.Regenerate()

	def ViewNameIsUniq(self, name):
		for view in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).ToElements():
			try:
				if str(view.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString()) == name:
					return False
			except: pass
		return True

	def SheetNameIsUniq(self, name):
		for sheet in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).ToElements():
			try:
				if str(sheet.get_Parameter(DB.BuiltInParameter.SHEET_NAME).AsString()) == name:
					return False
			except: pass
		return True

	def GetElevationDescription(self, length_feet):
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

	def GetEdges(self):
		floor_face = None
		edge_list_01 = []
		edge_list_02 = []
		shell = self.Room.ClosedShell.GetEnumerator()
		check_loop = DB.CurveLoop()
		while shell.MoveNext():
			if type(shell.Current) == DB.Solid:
				for face in shell.Current.Faces:			
					if str(face.ComputeNormal(DB.UV(0.5, 0.5))) == str(DB.XYZ(0,0,1)) or str(face.ComputeNormal(DB.UV(0.5, 0.5))) == str(DB.XYZ(0,0,-1)):
						floor_face = face
						break
				for edge in floor_face.EdgeLoops[0]:
					check_loop.Append(edge.AsCurve())
					if edge.AsCurve().Length > 100 / 304.8:
						try:
							start_point = DB.XYZ(edge.AsCurve().GetEndPoint(0).X, edge.AsCurve().GetEndPoint(0).Y, 0.0)
							end_point = DB.XYZ(edge.AsCurve().GetEndPoint(1).X, edge.AsCurve().GetEndPoint(1).Y, 0.0)
							line = DB.Line.CreateBound(start_point, end_point)
							edge_list_01.append(line)
						except: pass
				edge_list_01.reverse()
				for edge in edge_list_01:
					if edge.Length > 100 / 304.8:
						edge_list_02.append(edge.CreateReversed())	

				CW = check_loop.IsCounterclockwise(DB.XYZ(0, 0, 1))
				if CW:
					new_loop = []
					edge_list_02.reverse()
					for i in edge_list_02:
						new_loop.append(i.CreateReversed())
					edge_list_02 = new_loop
				return edge_list_02

	def GetIntersectPointProtected(self, ln_1, ln_2):
		pt_1 = ln_1.GetEndPoint(1)
		pt_2 = ln_2.GetEndPoint(0)
		if str(pt_1) == str(pt_2):
			return pt_2
		elif self.PointOnLine(pt_1, ln_2):
			return pt_2
		elif self.PointOnLine(pt_2, ln_1):
			return pt_1
		max_x = max([pt_1.X, pt_2.X])
		min_x = min([pt_1.X, pt_2.X])
		max_y = max([pt_1.Y, pt_2.Y])
		min_y = min([pt_1.Y, pt_2.Y])
		x = math.fabs(max_x - min_x)
		y = math.fabs(max_y - min_y)
		distance = math.sqrt(x*x + y*y)
		if distance <= 0.00:
			return pt_2
		d_dir = pt_2 - pt_1
		d_dir.Normalize()
		if str(d_dir) == str(ln_1.Direction) or str(d_dir.Negate()) == str(ln_1.Direction):
			return pt_1
		if str(d_dir) == str(ln_2.Direction) or str(d_dir.Negate()) == str(ln_2.Direction):
			return pt_2
		angle_b = ln_1.Direction.AngleTo(ln_2.Direction)
		angle_y = ln_1.Direction.AngleTo(d_dir)
		add_range = round(distance, 10) * round((math.sin(angle_b + angle_y)), 10) / round(math.sin(angle_b), 10)
		multiplier = DB.XYZ(ln_1.Direction.X * add_range, ln_1.Direction.Y * add_range, 0.0)
		intersection_point = pt_1 + multiplier
		return intersection_point

	def PointOnLine(self, pt, ln):
		try:
			dir_1 = ln.Direction
			dir_2 = DB.Line.CreateBound(ln.GetEndPoint(0), pt).Direction
			dir_3 = DB.Line.CreateBound(pt, ln.GetEndPoint(1)).Direction
			if str(dir_1) == str(dir_2) or str(dir_2) == str(dir_3):
				return True
			return False
		except Exception as e:
			return True

	def IsParallel(self, line_1, line_2):
		if math.degrees(line_1.Direction.AngleTo(line_2.Direction)) < 10:
			return True
		return False

	def GetExpandedLoop(self, loop, offset):
		new_loop = DB.CurveLoop()
		new_points = []
		for z in range(0, len(loop)):
			p = loop[z].GetEndPoint(0)
			vectors = [loop[z-1].Direction, loop[z].Direction]
			vectors_p = self.GetVector(vectors)
			angle_wide = vectors_p[0].AngleTo(vectors_p[1])/2
			v = DB.XYZ((vectors_p[0].X + vectors_p[1].X)/2, (vectors_p[0].Y + vectors_p[1].Y)/2, (vectors_p[0].Z + vectors_p[1].Z)/2).Normalize()
			distance = offset / math.cos(angle_wide)
			vector = DB.XYZ(v.X * distance, v.Y * distance, v.Z).Negate()
			p += vector
			if len(new_points) == 0:
				new_points.append(p)
			elif str(p) != str(new_points[-1]):
				if z == len(loop)-1:
					try:
						DB.Line.CreateBound(p, new_points[0])
						new_points.append(p)
					except:
						pass
				else:
					try:
						DB.Line.CreateBound(p, new_points[-1])
						new_points.append(p)
					except:
						pass
		for m in range(0, len(new_points)):
			if m != len(new_points)-1:
				new_loop.Append(DB.Line.CreateBound(new_points[m], new_points[m+1]))
			else:
				if str(new_points[m]) != str(new_points[0]):
					new_loop.Append(DB.Line.CreateBound(new_points[m], new_points[0]))
		return new_loop

	def GetVector(self, vectors):
		new_vectors = []
		for vector in vectors:
			new_vectors.append(DB.XYZ(vector.Y, -vector.X, 0))
		return new_vectors

	def GetOptimizedEdges(self):
		edge_list_01 = []
		edge_list_02 = []
		edge_list_03 = []
		edge_list_04 = []
		edge_list_05 = []
		edge_loop = DB.CurveLoop()
		for edge in self.GetEdges():
			if edge.Length > 100 / 304.8:
				edge_list_01.append(edge)
		for i in range(0, len(edge_list_01)):
			if len(edge_list_02) == 0:
				edge_list_02.append(edge_list_01[i])
			elif i == len(edge_list_01)-1:
				if not (self.IsParallel(edge_list_01[i], edge_list_02[-1]) and self.IsParallel(edge_list_01[i], edge_list_02[0])):
					edge_list_02.append(edge_list_01[i])
			else:
				if not self.IsParallel(edge_list_01[i], edge_list_02[-1]):
					edge_list_02.append(edge_list_01[i])
		for i in range(0, len(edge_list_02)):
			try:
				if len(edge_list_03) == 0:
					s_p = self.GetIntersectPointProtected(edge_list_02[i], edge_list_02[-1])
					e_p = self.GetIntersectPointProtected(edge_list_02[i+1], edge_list_02[i])
					line = DB.Line.CreateBound(s_p, e_p)
					edge_list_03.append(line)
				elif i == len(edge_list_02)-1:
					s_p = self.GetIntersectPointProtected(edge_list_03[-1], edge_list_02[i])
					e_p = self.GetIntersectPointProtected(edge_list_03[0], edge_list_02[i])
					line = DB.Line.CreateBound(s_p, e_p)
					edge_list_03.append(line)
				else:
					s_p = self.GetIntersectPointProtected(edge_list_02[i], edge_list_03[-1])
					e_p = self.GetIntersectPointProtected(edge_list_02[i], edge_list_02[i+1])
					line = DB.Line.CreateBound(s_p, e_p)
					edge_list_03.append(line)
			except:
				pass
		for i in range(0, len(edge_list_03)):
			if i != len(edge_list_03)-1:
				line = DB.Line.CreateBound(edge_list_03[i].GetEndPoint(0), edge_list_03[i+1].GetEndPoint(0))
				edge_list_04.append(line)
			else:
				line = DB.Line.CreateBound(edge_list_03[i].GetEndPoint(0), edge_list_03[0].GetEndPoint(0))
				edge_list_04.append(line)
		for i in edge_list_04:
			edge_loop.Append(i)
		edge_loop.Flip()
		for z in edge_loop:
			edge_list_05.append(z)
		return edge_list_05

	def GetAnnotationSymbol_Axis(self):
		uiapp = __revit__.Application
		for instance in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_GenericAnnotation).ToElements():
			if instance.Family.Name == "SYS_Автоматическое заполнение_Ось" and instance.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == "Круг":
				return instance
		try:
			if uiapp.VersionNumber != 2016:
				doc.LoadFamily("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_LoadUtils\\Families\\SHEETS\\SYS_Автоматическое заполнение_Ось.rfa")
			else:
				doc.LoadFamily("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_LoadUtils\\Families_2016\\SHEETS\\SYS_Автоматическое заполнение_Ось.rfa")
		except: pass
		doc.Regenerate()
		for instance in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_GenericAnnotation).ToElements():
			if instance.Family.Name == "SYS_Автоматическое заполнение_Ось" and instance.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == "Круг":
				return instance
			return None

	def GetAnnotationSymbol_Rose(self):
		uiapp = __revit__.Application
		for instance in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_GenericAnnotation).ToElements():
			if instance.Family.Name == "SYS_Автоматическое заполнение_Развертка на плане" and instance.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == "Круг":
				return instance
		try:
			if uiapp.VersionNumber != 2016:
				doc.LoadFamily("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_LoadUtils\\Families\\SHEETS\\SYS_Автоматическое заполнение_Развертка на плане.rfa")
			else:
				doc.LoadFamily("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_LoadUtils\\Families_2016\\SHEETS\\SYS_Автоматическое заполнение_Развертка на плане.rfa")
		except: pass
		doc.Regenerate()
		for instance in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_GenericAnnotation).ToElements():
			if instance.Family.Name == "SYS_Автоматическое заполнение_Развертка на плане" and instance.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == "Круг":
				return instance
			return None