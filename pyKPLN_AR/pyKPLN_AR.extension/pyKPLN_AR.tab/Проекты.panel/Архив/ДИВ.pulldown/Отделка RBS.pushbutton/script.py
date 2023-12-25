# -*- coding: utf-8 -*-
"""
RBS_InnerFinishingWithFinishing

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "RBS - Отделка"
__doc__ = 'Присвоение значения RBS типоразмерам элементов отделки по связанным помещениям' \


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
import System
from System import Guid
from System.Windows.Forms import *
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert
import datetime

class ValueWeights():
	def __init__(self, element_type, value, room, op, parent):
		self.Id = element_type.Id.IntegerValue
		self.Type = element_type
		self.Values = [value]
		self.Parent = parent
		self.Rooms = [room]
		self.output = op

	def Append(self, value, room):
		self.Values.append(value)
		self.Rooms.append(room)

	def IsErrored(self):
		if len(set(self.Values)) > 1:
			return True
		return False
	def ToDepartments(self, list, room):
		if len(list) == 0:
			list.append([room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString(), [room]])
			return
		else:
			dep_value = room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString()
			for i in list:
				if i[0] == dep_value:
					i[1].append(room)
					return
		list.append([room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString(), [room]])

	def GetParentRBS(self, value):
		for obj in self.Parent.objects:
			if obj.Group.Text == value:
				for dep_type in self.Parent.value_list:
					if dep_type.Name == obj.CB.Text:
						if type(self.Type) == DB.WallType:
							return dep_type.Walls
						elif type(self.Type) == DB.FloorType:
							return dep_type.Floors
						elif type(self.Type) == DB.CeilingType:
							return dep_type.Ceilings


	def SetValue(self):
		departments = []
		for room in self.Rooms:
			self.ToDepartments(departments ,room)
		value = max(set(self.Values), key=self.Values.count)
		self.Type.LookupParameter("ДИВ_RBS_Код по классификатору").Set(value)
		if len(set(self.Values)) > 1:
			if value in ["НК303.04.01.10.02", "НК303.04.01.10.01", "НК303.04.01.10.03"]:
				self.output.print_html("<h3>{} <лоджии> : «{}»\n</h3>".format(self.Type.Category.Name, self.Type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()))
			else:
				self.output.print_html("<h3>{} : «{}»\n</h3>".format(self.Type.Category.Name, self.Type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()))
			self.output.print_html("\t<font color=#ff6666><b>Предупреждение: Подходит несколько вариантов значения RBS! По умолчанию выбрано «{}»</b></font>".format(value))
			for dep in departments:
				string = "\t«{}» : RBS = «{}»:\n".format(dep[0], self.GetParentRBS(dep[0]))
				i = 0
				for room in dep[1]:
					i+=1
					if i == 10:
						string += "\n"
						i = 1
					if i == 1:
						string += "\t\t"
					string += self.output.linkify(room.Id) + "\t"
				print(string)
			print("\n\n")
		else:
			if value in ["НК303.04.01.10.02", "НК303.04.01.10.01", "НК303.04.01.10.03"]:
				self.output.print_html("<h3><font color=#009900>{} <лоджии> : «{}» - {}</font></h3>".format(self.Type.Category.Name, self.Type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), value))
			else:
				self.output.print_html("<h3><font color=#009900>{} : «{}» - {}</font></h3>".format(self.Type.Category.Name, self.Type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), value))

class DepartmentGroup():
	def __init__(self, parent, department, y):
		self.Parent = parent
		self.Group = GroupBox()
		self.Group.Text = department
		self.Group.Size = Size(360, 45)
		self.Group.Location = Point(5, y)
		self.Group.Parent = self.Parent
		self.CB = ComboBox()
		self.CB.Parent = self.Group
		self.CB.Location = Point(5,15)
		self.CB.Size = Size(330, 10)
		self.CB.DropDownStyle = ComboBoxStyle.DropDownList
		for i in self.Parent.value_list:
			self.CB.Items.Add(i.Name)
		self.CB.Text = "<Не назначено>"
		for i in self.Parent.value_list:
			if i.Default.lower() == self.Group.Text.lower():
				self.CB.Text = i.Name

class ValueStack():
	def __init__(self, name, rbs_floors, rbs_walls, rbs_ceilings, def_dep):
		self.Name = name
		self.Floors = rbs_floors
		self.Walls = rbs_walls
		self.Ceilings = rbs_ceilings
		self.Default = def_dep

class CreateWindow(Form):
	def __init__(self):
		self.output = script.get_output()
		self.output.set_title("KPLN : Отчет")
		#INIT
		self.Name = "RBS Отделка"
		self.Text = "RBS Отделка"
		self.Size = Size(400, 400)
		self.MinimumSize = Size(400, 200)
		self.MaximumSize = Size(400, 800)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = True
		self.FormBorderStyle = FormBorderStyle.SizableToolWindow
		self.TopMost = True
		self.AutoScroll = True
		self.Errors = []
		self.Types = []
		self.department_list = []
		self.value_list = [ValueStack("<Не назначено>", "", "", "", ""),
					  ValueStack("Технические помещения", "D2.01", "D2.02", "D2.03", "Технические помещения"),
					  ValueStack("Служебные помещения", "D3.01", "D3.02", "D3.03", "Пом. службы эксплуатации"),
					  ValueStack("Места общественного пользования", "D4.01", "D4.02", "D4.03", "МОП"),
					  ValueStack("Коммерческие помещения", "D5.01", "D5.02", "D5.03", "Коммерческие помещения"),
					  ValueStack("Квартиры с отделкой", "D6.01", "D6.02", "D6.03", "Квартира"),
					  ValueStack("Квартиры без отделки", "D10.01", "D10.02", "D10.03", "Квартира")]
		self.objects = []

		for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
			try:
				if not room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString() in self.department_list:
					self.department_list.append(room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString())
			except:
				continue
		self.v = 10
		for i in self.department_list:
			self.objects.append(DepartmentGroup(self, i, self.v))
			self.v += 50
		self.Btn = Button()
		self.Btn.Text = "Ok"
		self.Btn.Parent = self
		self.Btn.Location = Point(5, self.v)
		self.Btn.Click += self.Button_Click
		self.Size = Size(400, self.v+62)

	def GetElementType(self, element):
		if type(element) == DB.Wall:
			return element.WallType
		elif type(element) == DB.Floor:
			return element.FloorType
		elif type(element) == DB.Ceiling:
			return doc.GetElement(element.GetTypeId())
		return None

	def AppendType(self, lement_type, value, room):
		for t in self.Types:
			if t.Id == lement_type.Id.IntegerValue:
				t.Append(value, room)
				return
		self.Types.append(ValueWeights(lement_type, value, room, self.output, self))

	def Button_Click(self, sender, e):
		self.Hide()
		try:
			with db.Transaction(name = "pyKPLN"):
				for collector in [DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements(),
						 DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements(),
						 DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()]:
					for element in collector:
						try:
							elem_type = self.GetElementType(element)
							if elem_type != None:
								if elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Отделка":
									room = self.GetRoomDepartment(element)
									if room != None:
										name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
										department = room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString()
										if name.lower() != "лоджия":
											for obj in self.objects:
												if obj.Group.Text == department:
													for dep_type in self.value_list:
														if dep_type.Name == obj.CB.Text:
															param = self.GetElementType(element).LookupParameter("ДИВ_RBS_Код по классификатору")
															if param != None:
																if type(element) == DB.Wall:
																	value = dep_type.Walls
																elif type(element) == DB.Floor:
																	value = dep_type.Floors
																elif type(element) == DB.Ceiling:
																	value = dep_type.Ceilings
																if value != None:
																	self.AppendType(self.GetElementType(element), value, room)
															else:
																error = "Отсутствует параметр типоразмера:\n«ДИВ_RBS_Код по классификатору»"
																if not error in self.Errors:
																	forms.alert(error)
																	self.Errors.append(error)
										else:
											param = self.GetElementType(element).LookupParameter("ДИВ_RBS_Код по классификатору")
											if param != None:
												if type(element) == DB.Wall:
													value = "D8.02"
												elif type(element) == DB.Floor:
													value = "D8.01"
												elif type(element) == DB.Ceiling:
													value = "D8.03"
												if value != None:
													self.AppendType(self.GetElementType(element), value, room)
											else:
												error = "Отсутствует параметр типоразмера:\n«ДИВ_RBS_Код по классификатору»"
												if not error in self.Errors:
													forms.alert(error)
													self.Errors.append(error)
						except Exception as e: 
							print(str(e))
							pass
				for lement_type in self.Types:
					try:
						if (lement_type.IsErrored()):
							lement_type.SetValue()
					except Exception as e: print(str(e))
				for lement_type in self.Types:
					try:
						if (not lement_type.IsErrored()):
							lement_type.SetValue()
					except Exception as e: print(str(e))
		except Exception as e: print(str(e))
		self.Close()

	def GetRoomDepartment(self, element):
		try:
			room = doc.GetElement(DB.ElementId(int(element.get_Parameter(Guid("49b81516-cf69-4bc4-931e-2856eae66c3f")).AsString())))
			if type(room) == DB.Architecture.Room:
				return room
		except: pass
		return None

form = CreateWindow()
Application.Run(form)