# -*- coding: utf-8 -*-
"""
OpenPlacerByTask

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Расставить\nпо заданию"
__doc__ = 'Валидированная расстановка отверстий по заданию от инженерных разделов.\nС формированием html-отчета и выводом на печать.\n\nВажно: в данной версии скрипта расчитываются только стены на основе линий (криволинейные и арочные будут проигнорированы).' \

"""
Архитекурное бюро KPLN

"""
import math
import time
import os
import System
from pyrevit.framework import clr
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit as Revit
from pyrevit import script
from pyrevit import forms
from pyrevit import revit, DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
import datetime
from System.Windows.Forms import *
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

class OpeningsSelectorFilterRectangle (UI.Selection.ISelectionFilter):
	def _init_(self):
		self.list = []

	def SetList(self, list):
		self.list = list

	def InList(self, element):
		for i in self.list:
			try:
				if i.Id.ToString() == element.Id.ToString():
					return True
			except: pass
		return False

	def AllowElement(self, element = DB.Element):
		if self.InList(element) and element.Category.Id == DB.ElementId(-2001140) and ("199_отверстие в стене прямоугольное" in str(element.Symbol.FamilyName).lower() or "199_отверстие в стене прямоугольное" in str(element.Name).lower()):
			return True
		return False
	def AllowReference(self, refer, point):
		return False

class OpeningsSelectorFilterRound (UI.Selection.ISelectionFilter):
	def _init_(self):
		self.list = []

	def SetList(self, list):
		self.list = list

	def InList(self, element):
		for i in self.list:
			try:
				if i.Id.ToString() == element.Id.ToString():
					return True
			except: pass
		return False

	def AllowElement(self, element = DB.Element):
		if self.InList(element) and element.Category.Id == DB.ElementId(-2001140) and ("199_отверстие в стене круглое" in str(element.Symbol.FamilyName).lower() or "199_отверстие в стене круглое" in str(element.Name).lower()):
			return True
		return False
	def AllowReference(self, refer, point):
		return False

class OpeningsSelectorFilter (UI.Selection.ISelectionFilter):
	def _init_(self):
		self.list = []

	def SetList(self, list):
		self.list = list

	def InList(self, element):
		for i in self.list:
			try:
				if i.Id.ToString() == element.Id.ToString():
					return True
			except: pass
		return False

	def AllowElement(self, element = DB.Element):
		if self.InList(element) and element.Category.Id == DB.ElementId(-2001140) and ("199_отверстие в стене круглое" in str(element.Symbol.FamilyName).lower() or "199_отверстие в стене круглое" in str(element.Name).lower() or "199_отверстие в стене прямоугольное" in str(element.Symbol.FamilyName).lower() or "199_отверстие в стене прямоугольное" in str(element.Name).lower()):
			return True
		return False
	def AllowReference(self, refer, point):
		return False

class Asset():
	def __init__(self, task, wallset, element):
		self.Task = []
		self.Wallset = wallset
		self.Element = element
		if type(task) is list:
			for i in range(0, len(task)): 
				self.Task.append(task[i])
		else:
			self.Task.append(task)

	def ContainsId(self, id):
		if self.Element.Id.ToString() == id.ToString():
			return True
		else:
			return False

	def IsGroup(self):
		if len(self.Task) > 1:
			return True
		else: return False

	def Remove(self):
		with db.Transaction(name = "RemoveElement"):
			try:
				doc.Delete(self.Element.Id)
				self.Element = None
			except: pass

	def SetElement(self, element):
		self.Element = element

	@property
	def Task(self):
		return self.Task

	@property
	def Wallset(self):
		return self.Wallset

	@property
	def Element(self):
		return self.Element

class WallSet():
	def __init__(self, wall, solid, centroid, bbox):
		self.Wall = wall
		self.Solid = solid
		self.Centroid = centroid
		self.BoundingBox = bbox

	@property
	def Wall(self):
		return self.Wall
	@property
	def Solid(self):
		return self.Solid
	@property
	def Centroid(self):
		return self.Centroid
	@property
	def BoundingBox(self):
		return self.BoundingBox

class Matrix():
	def __init__(self, BoundingBox):
		self.matrix = []
		self.minx = BoundingBox.Min.X
		self.maxx = BoundingBox.Max.X
		self.miny = BoundingBox.Min.Y
		self.maxy = BoundingBox.Max.Y
		self.minz = BoundingBox.Min.Z
		self.maxz = BoundingBox.Max.Z
		for x in range(int(self.minx), int(self.maxx)):
			for y in range(int(self.miny), int(self.maxy)):
				for z in range(int(self.minz), int(self.maxz)):
					matrix_1 = []
					Box_1 = DB.BoundingBoxXYZ()
					Box_1.Min = DB.XYZ(40 * x, 40 * y, 40 * z)
					Box_1.Max = DB.XYZ(40 * x + 40, 40 * y + 40, 40 * z + 40)
					boxes_1 = self.ExplodeBBox(Box_1)
					for Box_2 in boxes_1:
						boxes_2 = self.ExplodeBBox(Box_2)
						matrix_2 = []
						for b in boxes_2:
							matrix_2.append([b, []])
						matrix_1.append([Box_2, matrix_2])
					self.matrix.append([Box_1, matrix_1])

	def Flatten(self):
		try:
			op_matrix = []
			for matrix in self.matrix:
				m2 = []
				for matrix_1 in matrix[1]:
					m3 = []
					for matrix_2 in matrix_1[1]:
						if len(matrix_2[1]) != 0:
							m3.append(matrix_2)
					if len(m3) != 0:
						m2.append([matrix_1[0], m3])
				if len(m2) != 0:
					op_matrix.append([matrix[0], m2])
			self.matrix = []
			for i in op_matrix:
				self.matrix.append(i)
			return True
		except: return False

	def BboxIntersect(self, Box_1, Box_2):
		try:
			if Box_2.Max.Z < Box_1.Min.Z or Box_2.Min.Z > Box_1.Max.Z:
				return False
			if Box_2.Max.Y < Box_1.Min.Y or Box_2.Min.Y > Box_1.Max.Y:
				return False
			if Box_2.Max.X < Box_1.Min.X or Box_2.Min.X > Box_1.Max.X:
				return False
			return True
		except: return False

	def Insert(self, task):
		for i in self.matrix:
			box_1 = i[0]
			if self.BboxIntersect(task.BoundingBox, i[0]):
				for z in i[1]:
					box_2 = z[0]
					if self.BboxIntersect(task.BoundingBox, z[0]):
						for x in z[1]:
							if self.BboxIntersect(task.BoundingBox, x[0]):
								x[1].append(task)
		return True

	def TaskIsUniq(self, Task, list):
		try:
			if len(list) == 0:
				return True
			for i in list:
				if str(i.Element.Id) == str(Task.Element.Id):
					return False
			return True
		except: return False

	def GetInserts(self, box):
		list = []
		min_z = box.Min.Z
		max_z = box.Max.Z
		min_y = box.Min.Y
		max_y = box.Max.Y
		min_x = box.Min.X
		max_x = box.Max.X
		for i in self.matrix:
			box_1 = i[0]
			if max_z < box_1.Min.Z or min_z > box_1.Max.Z:
				continue
			if max_y < box_1.Min.Y or min_y > box_1.Max.Y:
				continue
			if max_x < box_1.Min.X or min_x > box_1.Max.X:
				continue
			for z in i[1]:
				box_2 = z[0]
				if max_z < box_2.Min.Z or min_z > box_2.Max.Z:
					continue
				if max_y < box_2.Min.Y or min_y > box_2.Max.Y:
					continue
				if max_x < box_2.Min.X or min_x > box_2.Max.X:
					continue
				for x in z[1]:
					box_3 = x[0]
					if max_z < box_3.Min.Z or min_z > box_3.Max.Z:
						continue
					if max_y < box_3.Min.Y or min_y > box_3.Max.Y:
						continue
					if max_x < box_3.Min.X or min_x > box_3.Max.X:
						continue
					for task in x[1]:
						if self.TaskIsUniq(task, list):
							list.append(task)
		return list

	def ExplodeBBox(self, boundingbox):
		min = boundingbox.Min
		max = boundingbox.Max
		w = math.fabs(max.X - min.X)
		bbox_mini_1 = DB.BoundingBoxXYZ()
		bbox_mini_1.Min = DB.XYZ(min.X, min.Y, min.Z)
		bbox_mini_1.Max = DB.XYZ(min.X + w, min.Y + w, min.Z + w)

		bbox_mini_2 = DB.BoundingBoxXYZ()
		bbox_mini_2.Min = DB.XYZ(min.X + w, min.Y, min.Z)
		bbox_mini_2.Max = DB.XYZ(min.X + 2 * w, min.Y + w, min.Z + w)

		bbox_mini_3 = DB.BoundingBoxXYZ()
		bbox_mini_3.Min = DB.XYZ(min.X, min.Y + w, min.Z)
		bbox_mini_3.Max = DB.XYZ(min.X + w, min.Y + 2 * w, min.Z + w)

		bbox_mini_4 = DB.BoundingBoxXYZ()
		bbox_mini_4.Min = DB.XYZ(min.X + w, min.Y + w, min.Z)
		bbox_mini_4.Max = DB.XYZ(min.X + 2 * w, min.Y + 2 * w, min.Z + w)

		bbox_mini_5 = DB.BoundingBoxXYZ()
		bbox_mini_5.Min = DB.XYZ(min.X, min.Y, min.Z + w)
		bbox_mini_5.Max = DB.XYZ(min.X + w, min.Y + w, min.Z + 2 * w)

		bbox_mini_6 = DB.BoundingBoxXYZ()
		bbox_mini_6.Min = DB.XYZ(min.X + w, min.Y, min.Z + w)
		bbox_mini_6.Max = DB.XYZ(min.X + 2 * w, min.Y + w, min.Z + 2 * w)

		bbox_mini_7 = DB.BoundingBoxXYZ()
		bbox_mini_7.Min = DB.XYZ(min.X, min.Y + w, min.Z + w)
		bbox_mini_7.Max = DB.XYZ(min.X + w, min.Y + 2 * w, min.Z + 2 * w)

		bbox_mini_8 = DB.BoundingBoxXYZ()
		bbox_mini_8.Min = DB.XYZ(min.X + w, min.Y + w, min.Z + w)
		bbox_mini_8.Max = DB.XYZ(min.X + 2 * w, min.Y + 2 * w, min.Z + 2 * w)

		return [bbox_mini_1, bbox_mini_2, bbox_mini_3, bbox_mini_4, bbox_mini_5, bbox_mini_6, bbox_mini_7, bbox_mini_8]

class TaskElement():
	def __init__(self, element, bbox, centroid, solid, round):
		self.Element = element
		self.BoundingBox = bbox
		self.Centroid = centroid
		self.Solid = solid
		self.Intersection = None
		self.Round = round

	def SetIntersection(self, s):
		self.Intersection = s

	@property
	def Round(self):
		return self.Round

	@property
	def Intersection(self):
		return self.Intersection

	@property
	def Element(self):
		return self.Element

	@property
	def BoundingBox(self):
		return self.BoundingBox

	@property
	def Centroid(self):
		return self.Centroid

	@property
	def Solid(self):
		return self.Solid

class MainForm(Form):
	def __init__(self):
		self.Name = "KPLN_KR_Openings"
		self.Text = "KPLN Отверстия"
		self.Size = Size(600, 400)
		self.MinimumSize = Size(600, 400)
		self.MaximumSize = Size(600, 400)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = True
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.TopMost = True
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
		#ELEMENTS
		self.walls = []
		self.tasks = []
		self.created_elements = []
		self.created_elements_post = []
		self.assets = []
		#CURRENT
		self.active_wall = None
		self.step = 0
		self.current_elements = []
		#OUTPUT
		self.output = script.get_output()
		self.output.set_title("KPLN : Отчет по расстановке отверстий")
		self.report_succeeded = []
		self.report_failures = []
		#
		self.gb1 = GroupBox()
		self.gb2 = GroupBox()
		self.button1 = Button()
		self.button2 = Button()
		self.button3 = Button()
		self.button4 = Button()
		self.button5 = Button()
		self.button6 = Button()
		self.button7 = Button()
		self.button8 = Button()
		self.button9 = Button()
		self.label1 = Label()
		self.richTextBox1 = RichTextBox()
		self.pb = PictureBox()
		self.pb.SizeMode = PictureBoxSizeMode.Zoom
		self.pb.BackColor =  System.Drawing.Color.FromArgb(255, 255,255,255)

		self.gb1.Location = Point(12, 214)
		self.gb1.Size = Size(275, 107)
		self.gb1.Text = "Действия"
		self.gb1.Parent = self

		self.gb2.Location = Point(297, 214)
		self.gb2.Size = Size(275, 107)
		self.gb2.Text = "Добавить в отчет"
		self.gb2.Parent = self

		self.button1.Location = Point(6, 19)
		self.button1.Size = Size(263, 23)
		self.button1.Text = "Сгруппировать выбранные"
		self.button1.Parent = self.gb1
		self.button1.Click += self.JoinElements

		self.button2.Location = Point(6, 48)
		self.button2.Size = Size(263, 23)
		self.button2.Text = "Разгруппировать выбранные"
		self.button2.Parent = self.gb1
		self.button2.Click += self.UnjoinElements

		self.button3.Location = Point(6, 77)
		self.button3.Size = Size(263, 23)
		self.button3.Text = "Удалить выбранные"
		self.button3.Parent = self.gb1
		self.button3.Click += self.RemoveElements

		self.button4.Location = Point(6, 19)
		self.button4.Size = Size(263, 23)
		self.button4.Text = "Утвердить"
		self.button4.Parent = self.gb2
		self.button4.Click += self.Result_Succeed

		self.button5.Location = Point(6, 48)
		self.button5.Size = Size(263, 23)
		self.button5.Text = "Оклонить и удалить все"
		self.button5.Parent = self.gb2
		self.button5.Click += self.Result_Failed_Removed

		self.button6.Location = Point(6, 77)
		self.button6.Size = Size(263, 23)
		self.button6.Text = "Сохранить с пометкой"
		self.button6.Parent = self.gb2
		self.button6.Click += self.Result_Failed

		self.button7.Location = Point(12, 327)
		self.button7.Size = Size(75, 23)
		self.button7.Text = "Пропустить"
		self.button7.Parent = self
		self.button7.Click += self.SkipStep

		self.button8.Location = Point(93, 327)
		self.button8.Size = Size(75, 23)
		self.button8.Text = "Завершить"
		self.button8.Parent = self
		self.button8.Click += self.Stop

		self.button9.Location = Point(200, 34)
		self.button9.Size = Size(75, 23)
		self.button9.Text = "Обновить"
		self.button9.Parent = self
		self.button9.Click += self.Refresh

		self.label1.AutoSize = True
		self.label1.Location = Point(294, 9)
		self.label1.Name = "label1"
		self.label1.Size = Size(111, 13)
		self.label1.Text = "Комментарий"
		self.label1.Parent = self

		self.richTextBox1.Location = Point(297, 28)
		self.richTextBox1.Name = "richTextBox1"
		self.richTextBox1.Size = Size(275, 180)
		self.richTextBox1.Text = ""
		self.richTextBox1.Parent = self
		self.richTextBox1.BorderStyle = BorderStyle.Fixed3D

		self.pb.Location = Point(12, 28)
		self.pb.Name = "pictureBox1"
		self.pb.Size = Size(269, 180)
		self.pb.Parent = self
		self.pb.BorderStyle = BorderStyle.Fixed3D
		self.CenterToScreen()
		#GET WALLS
		min_x = 999999
		min_y = 999999
		min_z = 999999
		max_x = -999999
		max_y = -999999
		max_z = -999999
		for wall in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements():
			if wall.Width >= 99 / 304.8:
				options = DB.Options()
				options.ComputeReferences = False
				options.IncludeNonVisibleObjects = False
				wall_geometry = wall.get_Geometry(options)
				for item in wall_geometry:
					if item.GetType() == DB.Solid:
						pt = item.ComputeCentroid()
						wall_bbox = item.GetBoundingBox()
						normalized_box = DB.BoundingBoxXYZ()
						normalized_box.Min = DB.XYZ(wall_bbox.Min.X + pt.X, wall_bbox.Min.Y + pt.Y, wall_bbox.Min.Z + pt.Z)
						normalized_box.Max = DB.XYZ(wall_bbox.Max.X + pt.X, wall_bbox.Max.Y + pt.Y, wall_bbox.Max.Z + pt.Z)
						self.walls.append(WallSet(wall, item, pt, normalized_box))
						if min_x > wall_bbox.Min.X + pt.X:
							min_x = wall_bbox.Min.X + pt.X
						if min_y > wall_bbox.Min.Y + pt.Y:
							min_y = wall_bbox.Min.Y + pt.Y
						if min_z > wall_bbox.Min.Z + pt.Z:
							min_z = wall_bbox.Min.Z + pt.Z
						if max_x < wall_bbox.Max.X + pt.X:
							max_x = wall_bbox.Max.X + pt.X
						if max_y < wall_bbox.Max.Y + pt.Y:
							max_y = wall_bbox.Max.Y + pt.Y
						if max_z < wall_bbox.Max.Z + pt.Z:
							max_z = wall_bbox.Max.Z + pt.Z
		if len(self.walls) > 0:
			minx = min(int(math.copysign(math.ceil(math.fabs(min_x) / 40 ), min_x)), int(math.copysign(math.floor(math.fabs(min_x) / 40 ), min_x)))
			miny = min(int(math.copysign(math.ceil(math.fabs(min_y) / 40), min_y)), int(math.copysign(math.floor(math.fabs(min_y) / 40), min_y)))
			minz = min(int(math.copysign(math.ceil(math.fabs(min_z) / 40), min_z)), int(math.copysign(math.floor(math.fabs(min_z) / 40), min_z)))
			maxx = max(int(math.ceil(math.fabs(max_x) / 40) + math.fabs(minx)), int(math.floor(math.fabs(max_x) / 40) + math.fabs(minx)))
			maxy = max(int(math.ceil(math.fabs(max_y) / 40) + math.fabs(miny)), int(math.floor(math.fabs(max_y) / 40) + math.fabs(miny)))
			maxz = max(int(math.ceil(math.fabs(max_z) / 40) + math.fabs(minz)), int(math.floor(math.fabs(max_z) / 40) + math.fabs(minz)))
			self.globalBox = DB.BoundingBoxXYZ()
			self.globalBox.Min = DB.XYZ(minx, miny, minz)
			self.globalBox.Max = DB.XYZ(maxx, maxy, maxz)
			self.matrix = Matrix(self.globalBox)
			#GET TASKS
			for link in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements():
				try:
					document = link.GetLinkDocument()
					transform = link.GetTransform()
					for ref in DB.FilteredElementCollector(document).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements():
						if ref.Symbol.FamilyName == "501_Задание на отверстие в стене прямоугольное" or ref.Symbol.FamilyName == "501_Задание на отверстие в стене круглое":
							geometry_object = ref.GetGeometryObjectFromReference(DB.Reference(ref))
							geometry_object_transformed = geometry_object.GetTransformed(transform)
							for item in geometry_object_transformed:
								if item.GetType() == DB.Solid:
									task_solid = item
									try:
										pt = task_solid.ComputeCentroid()
									except: continue
									if ref.Symbol.FamilyName == "501_Задание на отверстие в стене круглое":
										round = True
									else:
										round = False
									bbox = task_solid.GetBoundingBox()
									bbox_normalized = DB.BoundingBoxXYZ()
									bbox_normalized.Min = DB.XYZ(bbox.Min.X + pt.X, bbox.Min.Y + pt.Y, bbox.Min.Z + pt.Z)
									bbox_normalized.Max = DB.XYZ(bbox.Max.X + pt.X, bbox.Max.Y + pt.Y, bbox.Max.Z + pt.Z)
									task = TaskElement(ref, bbox_normalized, pt, item, round)
									self.tasks.append(task)
				except: pass
			if len(self.tasks) > 0:
				for task in self.tasks:
					self.matrix.Insert(task)
				self.matrix.Flatten()
		self.Load += self.OnLoad

	def AlignElements(self, element):
		try:
			pt = element.Host.Location.Curve.Evaluate(element.HostParameter, False)
			elementdirection = math.degrees(element.Host.Location.Curve.Direction.AngleTo(DB.XYZ(0, 1, 0))) + 90
			rads = math.radians(elementdirection)
			elementpoint = DB.XYZ(pt.X, pt.Y, 0.0)
			grids = []
			length = 9999.9
			for grid in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements():
				griddirection = math.degrees(grid.Curve.Direction.AngleTo(DB.XYZ(0, 1, 0)))
				griddirectionreversed = math.degrees(grid.Curve.CreateReversed().Direction.AngleTo(DB.XYZ(0, 1, 0)))
				if round(elementdirection, 2) == round(griddirection, 2) or round(elementdirection, 2) == round(griddirectionreversed, 2) or round(elementdirection, 2) - 180.0 == round(griddirection, 2) or round(elementdirection, 2) - 180.0 == round(griddirectionreversed, 2):
					grids.append(grid)
				else:
					pass
			itersectionresult = None
			for grid in grids:
				gc = grid.Curve
				curve = DB.Line.CreateBound(DB.XYZ(gc.GetEndPoint(0).X, gc.GetEndPoint(0).Y, 0.0), DB.XYZ(gc.GetEndPoint(1).X, gc.GetEndPoint(1).Y, 0.0))
				itersectionresult = curve.Project(elementpoint)	
				try:
					if itersectionresult.Distance < length:
						length = itersectionresult.Distance
				except:
					pass
			if itersectionresult != None:
				v = DB.Line.CreateBound(pt, itersectionresult.XYZPoint).Direction.Normalize()
				if length != 999.9:
					rounded_length = round(length * 304.8 / 100, 1) * 100 / 304.8
					def_difference = length - rounded_length
					if self.DistanceEQ(DB.XYZ(elementpoint.X + v.X * def_difference, elementpoint.Y + v.Y * def_difference, elementpoint.Z), itersectionresult.XYZPoint, rounded_length):
						difference = length - rounded_length
					else:
						difference = rounded_length - length
					move_pt = DB.XYZ(v.X * difference, v.Y * difference, 0.0)
					DB.ElementTransformUtils.MoveElement(doc, element.Id, move_pt)
		except:
			pass

	def DistanceEQ(self, point_1, point_2, distance):
		line = DB.Line.CreateBound(point_1, point_2)
		if round(line.Length, 2) == round(distance, 2):
			return True
		else:
			return False

	def OnLoad(self, e):
		self.Disable()
		self.button7.Text = "Начать!"
		self.button7.Enabled = True

	def GetLevel(self, wallset):
		return doc.GetElement(wallset.Wall.LevelId)

	def Result_Succeed(self, sender, args):
		self.Disable()
		self.report_succeeded.append(["Задание утверждено", self.created_elements, self.active_image, self.active_wall, self.GetLevel(self.active_wall), self.richTextBox1.Text])
		self.NextStep()

	def Result_Failed(self, sender, args):
		self.Disable()
		self.report_failures.append(["Утверждено с корректировками", self.created_elements, self.active_image, self.active_wall, self.GetLevel(self.active_wall), self.richTextBox1.Text])
		self.NextStep()

	def Result_Failed_Removed(self, sender, args):
		self.Disable()
		self.report_failures.append(["Не утверждено", self.created_elements, self.active_image, self.active_wall, self.GetLevel(self.active_wall), self.richTextBox1.Text])
		with db.Transaction(name = "RemoveElements"):
			for item in self.created_elements:
				try:
					doc.Delete(item.Id)
				except: pass
		self.NextStep()

	def JoinElements(self, sender, args):
		pass
		self.Hide()
		try:
			join_elements = []
			filter_openings = OpeningsSelectorFilterRectangle()
			filter_openings.SetList(self.created_elements)
			elements = uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element, filter_openings, "KPLN : Веберите отверстия для объединения")
			if len(elements) == 1:
				forms.alert("Для объединения необходимо выбрать несколько элементов!")
			else:
				for item in elements:
					try:
						join_elements.append(doc.GetElement(item))
					except: pass
			if len(join_elements) > 1:
				self.PlaceOpenCombined(self.active_wall, join_elements)
		except: pass
		self.Show()

	def UnjoinElements(self, sender, args):
		self.Hide()
		try:
			unjoin_elements = []
			filter_openings = OpeningsSelectorFilterRectangle()
			filter_openings.SetList(self.created_elements)
			element = doc.GetElement(uidoc.Selection.PickObject(UI.Selection.ObjectType.Element, filter_openings, "KPLN : Веберите отверстие для разделения"))
			for asset in self.assets:
				try:
					if asset.ContainsId(element.Id):
						if asset.IsGroup():
							self.PlaceUnjoined(asset)
						else:
							forms.alert("Элемент не возможно разделить!")
				except: 
					pass
		except: 
			pass
		self.Show()

	def RemoveElements(self, sender, args):
		self.Hide()
		try:
			filter_openings = OpeningsSelectorFilter()
			filter_openings.SetList(self.created_elements)
			elements = uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element, filter_openings, "KPLN : Веберите отверстия для удаления")
			for asset in self.assets:
				try:
					for reference in elements:
						element = doc.GetElement(reference)
						if asset.ContainsId(element.Id):
							asset.Remove()
				except:
					pass
		except: 
			pass
		self.Show()

	def Refresh(self, sender, args):
		self.UpdateView(doc.ActiveView)

	def CloseForm(self):
		self.Close()

	def Stop(self, sender, args):
		pass
		self.Hide()
		global dialog
		result = dialog.show()
		self.Show()
		if result == "1":
			self.report_succeeded.append(["Задание утверждено", self.created_elements, self.active_image, self.active_wall, self.GetLevel(self.active_wall), self.richTextBox1.Text])
		elif result == "2":
			self.report_failures.append(["Не утверждено", self.created_elements, self.active_image, self.active_wall, self.GetLevel(self.active_wall), self.richTextBox1.Text])
			for asset in self.assets:
				asset.Remove()
		elif result == "3":
			self.report_failures.append(["Утверждено с корректировками", "Удалены", self.active_image, self.active_wall, self.GetLevel(self.active_wall), self.richTextBox1.Text])
		r_results = []
		r_ids = []
		r_images = []
		r_walls = []
		r_levels = []
		r_comments = []
		for i in self.report_failures:
			try:
				r_ids.append(i[1])
				r_results.append(i[0])
				r_images.append(i[2])
				r_walls.append(i[3])
				r_levels.append(i[4])
				r_comments.append(i[5])
			except: pass
		for i in self.report_succeeded:
			try:
				r_ids.append(i[1])
				r_results.append(i[0])
				r_images.append(i[2])
				r_walls.append(i[3])
				r_levels.append(i[4])
				r_comments.append(i[5])
			except: pass
		self.SetOffset()
		self.PrintReport(r_results, r_ids, r_images, r_walls, r_levels, r_comments)
		self.CloseForm()

	def Disable(self):
		self.richTextBox1.Enabled = False
		self.button1.Enabled = False
		self.button2.Enabled = False
		self.button3.Enabled = False
		self.button4.Enabled = False
		self.button5.Enabled = False
		self.button6.Enabled = False
		self.button7.Enabled = False
		self.button8.Enabled = False
		self.button9.Enabled = False

	def Enable(self):
		self.richTextBox1.Enabled = True
		self.button1.Enabled = True
		self.button2.Enabled = True
		self.button3.Enabled = True
		self.button4.Enabled = True
		self.button5.Enabled = True
		self.button6.Enabled = True
		self.button7.Enabled = True
		self.button8.Enabled = True
		self.button9.Enabled = True

	def HasOpenings(self, task):
		bbox_1 = task.BoundingBox
		options = DB.Options()
		options.ComputeReferences = True
		options.IncludeNonVisibleObjects = True
		for element in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements():
			if element.Symbol.FamilyName == "199_Отверстие в стене прямоугольное" or element.Symbol.FamilyName == "199_Отверстие в стене круглое":
				for item in element.get_Geometry(options):
					bbox_2 = item.GetInstanceGeometry().GetBoundingBox()
					if self.BboxIntersect(bbox_1, bbox_2):
						return True
		return False

	def BboxIntersect(self, Box_1, Box_2):
		try:
			if Box_2.Max.Z < Box_1.Min.Z or Box_2.Min.Z > Box_1.Max.Z:
				return False
			if Box_2.Max.Y < Box_1.Min.Y or Box_2.Min.Y > Box_1.Max.Y:
				return False
			if Box_2.Max.X < Box_1.Min.X or Box_2.Min.X > Box_1.Max.X:
				return False
			return True
		except: return False

	def GetIntersection(self, wall_solid, element_solid):
		try:
			intersection_result = DB.BooleanOperationsUtils.ExecuteBooleanOperation(wall_solid, element_solid, DB.BooleanOperationsType.Intersect)
			if abs(intersection_result.Volume) > 0.0001:	
				return intersection_result
			return False
		except: return False

	def GetIntersections(self, wallset):
		list = []
		try:
			tasks = self.matrix.GetInserts(wallset.BoundingBox)
			for task in tasks:
				intersection = self.GetIntersection(wallset.Solid, task.Solid)
				if intersection != False:
					task.SetIntersection(intersection)
					if not self.HasOpenings(task) and intersection.Volume >= task.Solid.Volume * 0.5:
						list.append(task)
		except: pass
		return list


	def CalculateProperties(self, geometry, wall):
		points = []
		if type(geometry) == DB.Solid:
			for edge in geometry.Edges:
				curve = edge.AsCurve()
				points.append(curve.GetEndPoint(0))
				points.append(curve.GetEndPoint(1))
		else:
			for s in geometry:
				for edge in s.Edges:
					curve = edge.AsCurve()
					points.append(curve.GetEndPoint(0))
					points.append(curve.GetEndPoint(1))
		curve = wall.Location.Curve
		projected_points = []
		length = []
		length_sorted = []
		lines = []
		lines_sorted = []
		for point in points:
			projected_points.append(curve.Project(point).XYZPoint)
		for pt_1 in projected_points:
			for pt_2 in projected_points:
				if str(pt_1) != (str(pt_2)):
					try:
						line = DB.Line.CreateBound(pt_1, pt_2)
						length.append(line.Length)
						length_sorted.append(line.Length)
						lines.append(line)
						lines_sorted.append(line)
					except: pass
		length_sorted.sort()
		for i in range(0, len(length_sorted)):
			for z in range(0, len(length)):
				if length_sorted[i] == length[z]:
					lines_sorted.append(lines[z])
		largest_line = lines_sorted[-1]
		cp = DB.XYZ((largest_line.GetEndPoint(0).X + largest_line.GetEndPoint(1).X)/2, (largest_line.GetEndPoint(0).Y + largest_line.GetEndPoint(1).Y)/2, (largest_line.GetEndPoint(0).Z + largest_line.GetEndPoint(1).Z)/2)
		return [length_sorted[-1], cp]

	def GetCreationValues(self, intersection, centroid, wall):
		intersection_box = intersection.GetBoundingBox()
		height = round(math.fabs(intersection_box.Max.Z - intersection_box.Min.Z) / 100 * 304.8, 1) * 100 / 304.8
		calculations = self.CalculateProperties(intersection, wall)
		width = round(calculations[0] / 100 * 304.8, 1) * 100 / 304.8
		center_point = DB.XYZ(calculations[1].X, calculations[1].Y, (centroid.Max.Z + centroid.Min.Z)/2)
		project_level_elevation = doc.GetElement(wall.LevelId).Elevation
		elevation = round((center_point.Z - project_level_elevation - height/2) / 500 * 304.8, 1) * 500 / 304.8
		return [center_point, height, width, elevation]

	def Force_Stop(self):
		r_results = []
		r_ids = []
		r_images = []
		r_walls = []
		r_levels = []
		r_comments = []
		for i in self.report_failures:
			try:
				r_ids.append(i[1])
				r_results.append(i[0])
				r_images.append(i[2])
				r_walls.append(i[3])
				r_levels.append(i[4])
				r_comments.append(i[5])
			except: pass
		for i in self.report_succeeded:
			try:
				r_ids.append(i[1])
				r_results.append(i[0])
				r_images.append(i[2])
				r_walls.append(i[3])
				r_levels.append(i[4])
				r_comments.append(i[5])
			except: pass
		self.SetOffset()
		self.PrintReport(r_results, r_ids, r_images, r_walls, r_levels, r_comments)

	def CreateInstance(self, point, symbol, wallset, h, w, elevation, task):
		try:
			level = doc.GetElement(wallset.Wall.LevelId)
			with db.Transaction(name = "Create instance"):
				instance = doc.Create.NewFamilyInstance(point, symbol, wallset.Wall, level, DB.Structure.StructuralType.NonStructural)
				try:
					instance.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(elevation)
				except: pass
				instance.LookupParameter("Высота").Set(h)
				try:
					instance.LookupParameter("Ширина").Set(w)
				except: pass
				if type(task) == type([]):
					types = []
					for t in task:
						task_type = str(t.Element.Symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
						if not self.InList(task_type, types):
							types.append(task_type)
					instance.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(", ".join(types))
				else:
					instance.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(str(task.Element.Symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()))
				self.AlignElements(instance)
			created_asset = Asset(task, wallset, instance)
			self.assets.append(created_asset)
			self.created_elements.append(instance)
			self.created_elements_post.append(instance)
			return created_asset
		except:
			return False

	def get_description(self, length_feet):
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

	def SetHeightDescription(self, elements):
		collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
		local_param = "00_Отметка_Относительная"
		global_param = "00_Отметка_Абсолютная"
		host_param = "00_Основа_Элемента"
		family_rectangle = "199_Отверстие в стене прямоугольное"
		family_round = "199_Отверстие в стене круглое"
		common_parameters_file = "Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt"
		cpf_group = "АРХИТЕКТУРА - Дополнительные"
		default_offset_bp = 0.00
		try:
			room_params_type = "Text"
			param_found = [False, False, False]
			app = doc.Application
			category_set_elements = app.Create.NewCategorySet()
			insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Windows)
			category_set_elements.Insert(insert_cat_elements)
			insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Doors)
			category_set_elements.Insert(insert_cat_elements)
			insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_MechanicalEquipment)
			category_set_elements.Insert(insert_cat_elements)
			originalFile = app.SharedParametersFilename
			app.SharedParametersFilename = common_parameters_file
			SharedParametersFile = app.OpenSharedParameterFile()
			map = doc.ParameterBindings
			it = map.ForwardIterator()
			it.Reset()
			while it.MoveNext():
				d_Definition = it.Key
				d_Name = it.Key.Name
				d_Binding = it.Current
				d_catSet = d_Binding.Categories	
				for param, bool in zip([local_param, global_param, host_param], param_found):
					if d_Name == param:
						if d_Binding.GetType() == DB.InstanceBinding:
							if str(d_Definition.ParameterType) == "Text":
								if d_Definition.VariesAcrossGroups:
									if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Windows)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Doors)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_MechanicalEquipment)):
										bool == True
			with db.Transaction(name = "AddSharedParameter"):
				for dg in SharedParametersFile.Groups:
					if dg.Name == cpf_group:
						for param, bool in zip([local_param, global_param, host_param], param_found):
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
					for param in [local_param, global_param, host_param]:
						d_Definition = it.Key
						d_Name = it.Key.Name
						d_Binding = it.Current
						if d_Name == param:
							d_Definition.SetAllowVaryBetweenGroups(doc, True)
		except: pass
		#ОТНОСИТЕЛЬНАЯ ОТМЕТКА
		with db.Transaction(name='Local height'):
			for element in collector_elements:
				try:
					fam_name = element.Symbol.FamilyName 
					if fam_name.startswith(family_rectangle):
						base_height = element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
						value = "Низ на отм. " + self.get_description(base_height) + " мм от ур.ч.п."
						element.LookupParameter(local_param).Set(value)
					if fam_name.startswith(family_round):
						base_height = element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
						element_height = element.LookupParameter("Радиус").AsDouble()
						value = "Центр на отм. " + self.get_description(base_height + element_height) + " мм от ур.ч.п."
						element.LookupParameter(local_param).Set(value)
				except: pass

		#АБСОЛЮТНАЯ ОТМЕТКА
		collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
		collector_viewFamilyType = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType()
		for i in collector_viewFamilyType:
			if i.ViewFamily == DB.ViewFamily.ThreeDimensional:
				viewFamilyType = i
				break
		with db.Transaction(name='Global height'):
			zview = DB.View3D.CreateIsometric(doc, viewFamilyType.Id)
			base_point_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_ProjectBasePoint).ToElements()
			for b_point in base_point_collector:
				default_offset_bp = b_point.get_BoundingBox(zview).Max.Z
			for element in collector_elements:
				try:
					fam_name = element.Symbol.FamilyName 
					if fam_name.startswith(family_rectangle):
						down = element.LookupParameter("offset_down").AsDouble()
						b_box = element.get_BoundingBox(zview)
						boundingBox_Z_min = b_box.Min.Z + down - default_offset_bp
						value = "Низ на отм. " + self.get_description(boundingBox_Z_min) + " мм"
						element.LookupParameter(global_param).Set(value)
					if fam_name.startswith(family_round):
						down = element.LookupParameter("offset_down").AsDouble()
						up = element.LookupParameter("offset_up").AsDouble()
						b_box = element.get_BoundingBox(zview)
						boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
						value = "Центр на отм. " + self.get_description(boundingBox_Z_center) + " мм"
						element.LookupParameter(global_param).Set(value)
				except:
					pass
		try:
			doc.Delete(zview.Id)
		except:
			pass
		#ОСНОВА
		collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
		with db.Transaction(name='Запись параметров'):
			for element in collector_elements:
				try:
					name = element.Symbol.FamilyName
					if fam_name.startswith(family_rectangle) or fam_name.startswith(family_round):
						host = element.Host.Name
						parameter = element.LookupParameter(host_param).Set(host)
				except: pass

	def SetOffset(self):
		try:
			self.SetHeightDescription(self.created_elements_post)
			with db.Transaction(name = "ChangeElements"):
				for element in self.created_elements_post:
					try:
						offset_down =math.fabs( element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble())
						level = doc.GetElement(element.LevelId)
						level_elevation = level.Elevation
						offset_up = 3000 / 304.8 - element.LookupParameter("Высота").AsDouble() - offset_down - 250 / 304.8
						levels_all = []
						levels_height_all = []
						levels_height_all_sorted = []
						levels_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
						for lvl in levels_collector:
							levels_all.append(lvl)
							levels_height_all.append(lvl.Elevation)
							levels_height_all_sorted.append(lvl.Elevation)
						levels_height_all_sorted.sort()
						for i in range(0, len(levels_height_all_sorted)):
							if level_elevation == levels_height_all_sorted[i]:
								try:
									offset_up = math.fabs(levels_height_all_sorted[i+1] - level_elevation - element.LookupParameter("Высота").AsDouble() - offset_down - 250 / 304.8)
								except: pass
								element.LookupParameter("offset_down").Set(offset_down)
								element.LookupParameter("offset_up").Set(offset_up)
					except: pass
		except: pass

	def PlaceOpenSingle(self, wallset, task):
		global family_symbol
		global family_symbol_round
		try:
			with db.Transaction(name = "Activate symbol"):
				family_symbol.Activate()
		except: pass
		try:
			if task.Round:
				offset = round(task.Element.LookupParameter("Расширение границ").AsDouble() / 100 * 304.8, 1) * 100 / 304.8
				height = round(task.Element.LookupParameter("КП_Р_Высота").AsDouble() / 100 * 304.8, 1) * 100 / 304.8
				width = round(task.Element.LookupParameter("КП_Р_Ширина").AsDouble() / 100 * 304.8, 1) * 100 / 304.8
				elevation = round((task.Element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble() - offset) / 100 * 304.8, 1) * 100 / 304.8
				self.CreateInstance(task.Centroid, family_symbol_round, wallset, height, width, elevation, task)
			else:
				offset = round(task.Element.LookupParameter("Расширение границ").AsDouble() / 100 * 304.8, 1) * 100 / 304.8
				height = round(task.Element.LookupParameter("КП_Р_Высота").AsDouble() / 100 * 304.8, 1) * 100 / 304.8
				width = round(task.Element.LookupParameter("КП_Р_Ширина").AsDouble() / 100 * 304.8, 1) * 100 / 304.8
				elevation = round((task.Element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble() - offset) / 100 * 304.8, 1) * 100 / 304.8
				self.CreateInstance(task.Centroid, family_symbol, wallset, height, width, elevation, task)
		except:
			pass

	def InList(self, item, list):
		for i in list:
			if i == item:
				return True
		return False

	def PlaceUnjoined(self, picked_asset):
		tasks = []
		old_assets = []
		for asset in self.assets:
			try:
				if picked_asset.Element.Id.ToString() == asset.Element.Id.ToString():
					tasks = picked_asset.Task
					wallset = picked_asset.Wallset
					picked_asset.Remove()
					break
				else:
					old_assets.append(asset)
			except:
				pass
		self.assets = []
		for asset in old_assets:
			self.assets.append(asset)
		for task in tasks:
			try:
				self.PlaceOpenSingle(wallset, task)
			except: 
				pass

	def AssetInList(self, asset, list):
		for a in list:
			try:
				if asset.Element.Id.ToString() == a.Element.Id.ToString():
					return True
			except:
				return True
		return False

	def PlaceOpenCombined(self, wallset, instances):
		global family_symbol
		global family_symbol_round
		try:
			with db.Transaction(name = "Activate symbol"):
				family_symbol.Activate()
				family_symbol_round.Activate()
		except: pass
		old_assets = []
		union_assets = []
		min_x = 9999999
		min_y = 9999999
		min_z = 9999999
		max_x = -9999999
		max_y = -9999999
		max_z = -9999999
		for asset in self.assets:
			for instance in instances:
				if asset.ContainsId(instance.Id):
					union_assets.append(asset)
		for asset in self.assets:
			if not self.AssetInList(asset, union_assets):
				old_assets.append(asset)
		self.assets = []
		for asset in old_assets:
			self.assets.append(asset)
		intersections = []
		tasks = []
		for asset in union_assets:
			task_list = asset.Task
			for t in task_list:
				tasks.append(t)
				intersections.append(t.Intersection)
				bb = t.BoundingBox
				if max_z < bb.Max.Z: max_z = bb.Max.Z
				if min_z > bb.Min.Z: min_z = bb.Min.Z
		height = round(math.fabs(max_z - min_z) / 100 * 304.8, 1) * 100 / 304.8
		calculations_result = self.CalculateProperties(intersections, wallset.Wall)
		width = round(calculations_result[0] / 100 * 304.8, 1) * 100 / 304.8
		calc_point = calculations_result[1]
		point = DB.XYZ(calc_point.X, calc_point.Y, (max_z + min_z)/2)
		elevation = round((point.Z - doc.GetElement(wallset.Wall.LevelId).Elevation - height/2) / 500 * 304.8, 1) * 500 / 304.8
		with db.Transaction(name = "RemoveOpenings"):
			for remove_instance in instances:
				try:
					doc.Delete(remove_instance.Id)
				except: pass
		self.CreateInstance(point, family_symbol, wallset, height, width, elevation, tasks)

	def SkipStep(self, sender, args):
		with db.Transaction(name = "RemoveOpenings"):
			for item in self.created_elements:
				try:
					doc.Delete(item.Id)
				except: pass
		self.button7.Text = "Пропустить"
		self.Disable()
		self.richTextBox1.Text = ""
		self.assets = []
		self.created_elements = []
		for asset in self.assets:
			asset.Remove()
		self.Text = "KPLN Стены : {} из {}".format(self.step, len(self.walls))
		v = len(self.walls)
		with forms.ProgressBar(title='Wall {value} of {max_value}') as pb:
			while True:
				pb.update_progress(self.step, max_value = v)
				try:
					tasks = self.GetIntersections(self.walls[self.step])
					if len(tasks) > 0 and self.step < len(self.walls)-1:
						self.active_wall = self.walls[self.step]
						self.Text = "KPLN Стены : {} из {}".format(str(self.step), str(len(self.walls)))
						for t in tasks:
							self.PlaceOpenSingle(self.active_wall, t)
						self.ActivateView(self.active_wall)
						self.step += 1
						self.Text = "KPLN Стены : {} из {}".format(str(self.step), str(len(self.walls)))
						self.Enable()
						return
					else:
						self.step += 1
						self.Text = "KPLN Стены : {} из {}".format(str(self.step), str(len(self.walls)))
						if self.step >= len(self.walls)-1:
							self.Force_Stop()
							break
						continue
				except: 
					self.Force_Stop()
					break
				if self.step >= len(self.walls):
					self.Force_Stop()
					break
		self.CloseForm()

	def NextStep(self):
		self.Disable()
		self.richTextBox1.Text = ""
		self.assets = []
		self.created_elements = []
		self.Text = "KPLN Стены : {} из {}".format(self.step, len(self.walls))
		v = len(self.walls)
		with forms.ProgressBar(title='Wall {value} of {max_value}') as pb:
			while True:
				pb.update_progress(self.step, max_value = v)
				try:
					tasks = self.GetIntersections(self.walls[self.step])
					if len(tasks) > 0 and self.step < len(self.walls)-1:
						self.active_wall = self.walls[self.step]
						self.Text = "KPLN Стены : {} из {}".format(str(self.step), str(len(self.walls)))
						for t in tasks:
							self.PlaceOpenSingle(self.active_wall, t)
						self.ActivateView(self.active_wall)
						self.step += 1
						self.Text = "KPLN Стены : {} из {}".format(str(self.step), str(len(self.walls)))
						self.Enable()
						return
					else:
						self.step += 1
						self.Text = "KPLN Стены : {} из {}".format(str(self.step), str(len(self.walls)))
						if self.step >= len(self.walls)-1:
							self.Force_Stop()
							break
						continue
				except: 
					self.Force_Stop()
					break
				if self.step >= len(self.walls):
					self.Force_Stop()
					break
		self.CloseForm()

	def ActivateView(self, wallset):
		try:
			if uidoc.ActiveView.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() != "SYS_OPENINGS":
				collector_viewFamily = DB.FilteredElementCollector(doc).OfClass(DB.View3D).WhereElementIsNotElementType()
				bool = False
				for view in collector_viewFamily:
					if view.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() == "SYS_OPENINGS":
						zview = view
						bool = True
						uidoc.ActiveView = zview
				if not bool:
					collector_viewFamilyType = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType()
					for view in collector_viewFamilyType:
						if view.ViewFamily == DB.ViewFamily.ThreeDimensional:
							viewFamilyType = view
							break
					with db.Transaction(name = "CreateView"):
						zview = DB.View3D.CreateIsometric(doc, viewFamilyType.Id)
						zview.get_Parameter(DB.BuiltInParameter.VIEW_NAME).Set("SYS_OPENINGS")
					return False
			a_view = uidoc.ActiveView
			with db.Transaction(name = "Reset"):
				a_view.CropBoxActive = False
			with db.Transaction(name = "ModifyView"):
				pt = DB.XYZ(0,0,0)
				clipbox = DB.BoundingBoxXYZ()
				clipbox.Min = DB.XYZ(wallset.BoundingBox.Min.X + pt.X - 15, wallset.BoundingBox.Min.Y + pt.Y - 15, wallset.BoundingBox.Min.Z + pt.Z - 20)
				clipbox.Max = DB.XYZ(wallset.BoundingBox.Max.X + pt.X + 15, wallset.BoundingBox.Max.Y + pt.Y + 15, wallset.BoundingBox.Max.Z + pt.Z)
				a_view.SetSectionBox(clipbox)
				eye_position = DB.XYZ((wallset.BoundingBox.Min.X + wallset.BoundingBox.Max.X)/2, (wallset.BoundingBox.Min.Y + wallset.BoundingBox.Max.Y)/2, (wallset.BoundingBox.Min.Z + wallset.BoundingBox.Max.Z)/2)
				forward_direction = self.VectorFromHorizVertAngles(135, -30)
				up_direction = self.VectorFromHorizVertAngles(135, -30 + 90 )
				orientation = DB.ViewOrientation3D(eye_position, up_direction, forward_direction)
				a_view.SetOrientation(orientation)
			with db.Transaction(name = "Colorify"):
				for workset in DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset).ToWorksets():
					try:
						a_view.SetWorksetVisibility(workset.Id, DB.WorksetVisibility.Visible)
					except: pass
				a_view.DetailLevel = DB.ViewDetailLevel.Fine
				self.set_color(wallset.Wall, a_view)
				self.ZoomBbox(wallset, a_view)
				self.UpdateView(a_view)
		except: 
			print("except!")
			return False
		return True

	def set_color(self, wall, view):
		collector_elements = DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements()
		fillPatternElement = None
		dict = ["Solid fill", "Сплошная заливка", "сплошная заливка", "solid fill"]
		for name in dict:
			try:
				fillPatternElement = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Drafting, name)
				if fillPatternElement != None:
					break
			except: pass
		for built_in_category in [DB.BuiltInCategory.OST_DuctCurves, DB.BuiltInCategory.OST_PipeCurves, DB.BuiltInCategory.OST_Conduit, DB.BuiltInCategory.OST_CableTray, DB.BuiltInCategory.OST_FlexDuctCurves]:
			try:
				category = DB.Category.GetCategory(doc, built_in_category)
				color = DB.Color(255, 0, 0)
				settings = DB.OverrideGraphicSettings()
				settings.SetProjectionLineColor(DB.Color(0, 0, 0))
				settings.SetProjectionLineWeight(4)
				settings.SetHalftone(False)
				settings.SetProjectionFillPatternId(fillPatternElement.Id)
				settings.SetSurfaceTransparency(10)
				view.SetCategoryOverrides(category.Id, settings)
			except: pass
		for element in collector_elements:
			if str(wall.Id) != str(element.Id):
				color = DB.Color(255, 255, 255)
				settings = DB.OverrideGraphicSettings()
				settings.SetProjectionFillColor(color)
				settings.SetProjectionLineWeight(1)
				settings.SetHalftone(True)
				settings.SetProjectionFillPatternId(fillPatternElement.Id)
				settings.SetSurfaceTransparency(80)
				view.SetElementOverrides(element.Id, settings)
			else:
				color = DB.Color(10, 250, 0)
				settings = DB.OverrideGraphicSettings()
				settings.SetProjectionFillColor(color)
				settings.SetProjectionLineWeight(4)
				settings.SetHalftone(False)
				settings.SetProjectionFillPatternId(fillPatternElement.Id)
				settings.SetSurfaceTransparency(0)
				view.SetElementOverrides(element.Id, settings)

	def ZoomBbox(self, wallset, view):
		views = uidoc.GetOpenUIViews()
		for v in views:
			if str(v.ViewId) == str(view.Id):
				v.ZoomAndCenterRectangle(wallset.BoundingBox.Min, wallset.BoundingBox.Max)

	def VectorFromHorizVertAngles(self, angleHorizD, angleVertD):
		degToRadian = math.pi * 2 / 360
		angleHorizR = angleHorizD * degToRadian
		angleVertR = angleVertD * degToRadian
		a = math.cos(angleVertR)
		b = math.cos(angleHorizR)
		c = math.sin(angleHorizR)
		d = math.sin(angleVertR)
		return DB.XYZ(a*b, a*c, d)

	def UpdateView(self, view):
		log_username = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')
		system_username = log_username[len(log_username) - 1]
		now = datetime.datetime.now()
		time = "{}{}{}-{}_{}_{}".format(now.year, now.month, now.day, now.hour, now.minute, now.second)
		jpeg_folder = "{}_{}".format(system_username, "{}{}{}({})_{}".format(now.year, now.month, now.day, now.hour, system_username))
		file_name = doc.Title.replace('.rvt', '')
		file_name += "({})".format(time)
		views = List[DB.ElementId]([view.Id])
		export_options = DB.ImageExportOptions()
		export_options.ZoomType = DB.ZoomFitType.FitToPage
		export_options.PixelSize = 400
		export_options.FilePath = "Z:\\pyRevit\\Applications\\Reports\\{}\\{}.jpg".format(jpeg_folder, file_name)
		directory = "Z:\\pyRevit\\Applications\\Reports\\{}".format(jpeg_folder)
		try:
			os.stat(directory)
		except:
			os.mkdir(directory)
		export_options.FitDirection = DB.FitDirectionType.Horizontal
		export_options.ImageResolution = DB.ImageResolution.DPI_600
		export_options.ExportRange = DB.ExportRange.VisibleRegionOfCurrentView
		export_options.HLRandWFViewsFileType = DB.ImageFileType.JPEGLossless;
		export_options.ShadowViewsFileType = DB.ImageFileType.JPEGLossless;
		try:
			doc.ExportImage(export_options)
			self.pb.Image = Image.FromFile("Z:\\pyRevit\\Applications\\Reports\\{}\\{}.jpg".format(jpeg_folder, file_name))
			self.active_image = "Z:\\pyRevit\\Applications\\Reports\\{}\\{}.jpg".format(jpeg_folder, file_name)
		except: pass

	def PrintReport(self, results, ids, images, walls, levels, comments):
		log_username = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')
		system_username = log_username[len(log_username) - 1]
		now = datetime.datetime.now()
		time = "{}/{}/{} {}:{}:{}".format(now.year, now.month, now.day, now.hour, now.minute, now.second)
		self.output.print_html('<font size="4"><b>Отчет подготовил: </b><font color=#0066ff size="4"><ruby>{}<rt>Windows</rt></ruby> (<ruby>{}<rt>Revit</rt></ruby>)</font><br><b>Дата запуска:</b> {}<hr>'.format(system_username, doc.Application.Username, time))
		self.output.print_html('<b>Локальный путь:</b> {}'.format(doc.PathName))
		if doc.IsWorkshared:
			try:
				self.output.print_html('<b>Хранилище:</b> {}'.format(DB.BasicFileInfo.Extract(doc.PathName).CentralPath))
			except: pass
		self.output.print_html('<hr><br>')
		part = ""
		for i in range(0, len(results)):
			try:
				ids1 = ""
				part += '<tr>'
				for element in ids[i]:
					try:
						ids1 += self.output.linkify(element.Id)
						ids1 += '<br>'
					except:
						ids1 += ""
				if results[i] == "Не утверждено":
					part += '<td><center><font color=#ff6666><h3>{}</h1></font></center></td>'.format(results[i])
				elif results[i] == "Задание утверждено":
					part += '<td><center><font color=#009900><h3>{}</h1></font></center></td>'.format(results[i])
				else:
					part += '<td><center><font color=#b4b4b4><h3>{}</h1></font></center></td>'.format(results[i])
				part += '<td><center>{}</center></td>'.format(ids1)
				part += '<td><center><img src="{}" alt=""></center></td>'.format(str(images[i]))
				part += '<td><center>{}</center></td>'.format(self.output.linkify(walls[i].Wall.Id))
				part += '<td><center><q>{}</q></center></td>'.format(levels[i].Name)
				part += '<td><center>{}</center></td>'.format(comments[i])
				part += '</tr>'
			except: pass
		self.output.print_html('<table align="center" cols="6"><caption></caption>'\
				  '<tr>'\
				  '<th>Результат</th>'\
				  '<th>Id отверстий (задания)</th>'\
				  '<th>Эскиз</th>'\
				  '<th>Id стены</th>'\
				  '<th>Уровень</th>'\
				  '<th>Комментарий</th>' + part + '</table>')

commands = [CommandLink('Утвердить', return_value="1"), CommandLink('Отменить и удалить все', return_value="2"), CommandLink('Сохранить с пометкой', return_value="3")]
dialog = TaskDialog('Выберите действие для последнего расчета',
		title = "Группировка",
		title_prefix=False,
		content="",
		commands=commands,
		footer='',
		show_close=False)

family = False
family_round = False
family_symbol = None
family_symbol_round = None

try:
	if uidoc.ActiveView.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() != "SYS_OPENINGS":
		collector_viewFamily = DB.FilteredElementCollector(doc).OfClass(DB.View3D).WhereElementIsNotElementType()
		bool = False
		for view in collector_viewFamily:
			if view.get_Parameter(DB.BuiltInParameter.VIEW_NAME).AsString() == "SYS_OPENINGS":
				zview = view
				bool = True
				uidoc.ActiveView = zview
		if not bool:
			collector_viewFamilyType = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).WhereElementIsElementType()
			for view in collector_viewFamilyType:
				if view.ViewFamily == DB.ViewFamily.ThreeDimensional:
					viewFamilyType = view
					break
			with db.Transaction(name = "CreateView"):
				zview = DB.View3D.CreateIsometric(doc, viewFamilyType.Id)
				zview.get_Parameter(DB.BuiltInParameter.VIEW_NAME).Set("SYS_OPENINGS")
except: pass

with db.Transaction(name = "Activation"):
	for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment):
		if symbol.FamilyName == "199_Отверстие в стене прямоугольное":
			family = symbol.Family
			symbol.Activate()
			family_box = clr.StrongBox[DB.Family](family)
		if symbol.FamilyName == "199_Отверстие в стене круглое":
			family_round = symbol.Family
			symbol.Activate()
			family_box_round = clr.StrongBox[DB.Family](family)

with db.Transaction(name = "Load families"):
	if family == False:
		try:
			doc.LoadFamily("Z:\\pyRevit\\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\\CUT_Openings\\199_Отверстие в стене прямоугольное.rfa")
		except: pass
	else:
		try:
			doc.LoadFamily("Z:\\pyRevit\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\CUT_Openings\\199_Отверстие в стене прямоугольное.rfa", load_options, family_box)
		except: pass
	if family_round == False:
		try:
			doc.LoadFamily("Z:\\pyRevit\\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\\CUT_Openings\\199_Отверстие в стене круглое.rfa")
		except: pass
	else:
		try:
			doc.LoadFamily("Z:\\pyRevit\pyKPLN_KR (alpha)\\pyKPLN_LoadUtils\\Families\CUT_Openings\\199_Отверстие в стене круглое.rfa", load_options, family_box_round)
		except: pass

for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment):
	if symbol.FamilyName == "199_Отверстие в стене прямоугольное":
		if symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == "КР":
			family_symbol = symbol
	elif symbol.FamilyName == "199_Отверстие в стене круглое":
		if symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == "КР":
			family_symbol_round = symbol
if family_symbol != None:
	form = MainForm()
	Application.Run(form)
