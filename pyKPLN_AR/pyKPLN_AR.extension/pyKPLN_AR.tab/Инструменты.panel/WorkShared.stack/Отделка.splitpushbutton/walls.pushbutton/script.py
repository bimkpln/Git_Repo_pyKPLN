# -*- coding: utf-8 -*-
"""
FW_Walls

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Расчитать стены и плинтусы"
__doc__ = 'Определение принадлежности элементов к помещениям.\n\n' \
          'Список параметров элементов:\n' \
          '«О_Имя помещения» - имя помещения которому принадлежит элемент\n ' \
          '«О_Назначение помещения» - назначение помещения которому принадлежит элемент\n ' \
          '«О_Id помещения» - id помещения которому принадлежит элемент\n ' \
		  '«О_Номер помещения» - номер помещения которому принадлежит элемент\n ' \
		  '«О_Номер секции» - номер секции из «ПОМ_Секция» связанного помещения\n ' \
		  '«О_Тип» - метод определения принадлежности (см.«?»)\n ' \
		  '«О_Плинтус_Высота» - параметр типоразмера, определяющий высоту плинтуса\n ' \
		  '«О_Плинтус» - параметр типоразмера (определяет является ли стена плинтусом или нет)\n [длина плинтуса] = [площадь стены] / [«О_Плинтус_Высота»]\n\n' \
          'Предупреждение: Расчет может занять продолжительное кол-во времени (в среднем: 15-30мин)' \


"""
Архитекурное бюро KPLN

"""
import webbrowser
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
from System.Windows.Forms import *
from System.Drawing import *
import datetime
from rpw.ui.forms import Alert, CommandLink, TaskDialog

class CreateWindow(Form):
	def __init__(self): 
		self.Name = "KPLN_AR_Отделка_Зависимости"
		self.Text = "KPLN Отделка - Зависимости"
		self.H = 300
		self.Size = Size(600, self.H)
		self.MinimumSize = Size(600, self.H)
		self.MaximumSize = Size(600, self.H)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = True
		self.TopMost = True
		self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")	
		#VARIABLES
		self.collector_rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
		self.collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
		self.type_description = "WALLS"
		self.rooms = []
		self.elements = []
		for room in self.collector_rooms:
			self.rooms.append(room)
		for element in self.collector_elements:
			self.elements.append(element)
		self.detected = 0
		self.maximun = 0
		self.unresult_elements = []
		self.def_params = ["О_Имя помещения", "О_Назначение помещения", "О_Id помещения", "О_Номер помещения", "О_Номер секции"]
		self.state_param = "О_Тип"
		#UI
		self.gb = []
		self.cb = []
		self.eq = ["Равно", "Не равно", "Больше", "Меньше", "Больше или равно", "Меньше или равно"]
		self.room_parameters = []
		self.element_parameters = []
		self.elementtype_parameters = []
		if len(self.rooms) != 0:
			for room in self.rooms:
				for j in room.Parameters:
					el = "{}".format(j.Definition.Name)
					if str(j.StorageType) == "String":
						if self.check_parameter_elements(el, self.rooms):
							self.room_parameters.append(el)
					elif str(j.StorageType) == "Double":
						if self.check_parameter_elements(el, self.rooms):
							self.room_parameters.append(el)
				break
		if len(self.elements) != 0:
			for element in self.elements:
				elementtype = self.getType(element)
				if elementtype != False:
					for j in element.Parameters:
						el = "Экземпляр: {}".format(j.Definition.Name)
						if str(j.StorageType) == "String":
							if self.check_parameter_elements(el, self.elements):
								self.element_parameters.append(el)
						elif str(j.StorageType) == "Double":
							if self.check_parameter_elements(el, self.elements):
								self.element_parameters.append(el)
					for j in elementtype.Parameters:
						ty = "Типоразмер: {}".format(j.Definition.Name)
						if str(j.StorageType) == "String":
							if self.check_parameter_elements(ty, self.elements):
								self.elementtype_parameters.append(ty)
						elif str(j.StorageType) == "Double":
							if self.check_parameter_elements(ty, self.elements):
								self.elementtype_parameters.append(ty)
					break
		self.room_parameters.sort()
		self.element_parameters.sort()
		self.elementtype_parameters.sort()

		self.gb.append(GroupBox())
		self.gb[0].Text = "Фильтр для помещений"
		self.gb[0].Size = Size(573, 100)
		self.gb[0].Location = Point(5, 10)
		self.gb[0].Parent = self

		self.cb.append(ComboBox())
		self.cb[0].Parent = self.gb[0]
		self.cb[0].Size = Size(180, 10)
		self.cb[0].Location = Point(15, 42)
		self.cb[0].DropDownStyle = ComboBoxStyle.DropDownList
		self.cb[0].SelectedIndexChanged += self.CblOnChanged
		self.cb[0].FlatStyle = FlatStyle.Flat

		self.cb.append(ComboBox())
		self.cb[1].Parent = self.gb[0]
		self.cb[1].Size = Size(180, 10)
		self.cb[1].Location = Point(380, 42)
		self.cb[1].DropDownStyle = ComboBoxStyle.DropDownList
		self.cb[1].FlatStyle = FlatStyle.Flat

		self.cb.append(ComboBox())
		self.cb[2].Parent = self.gb[0]
		self.cb[2].Size = Size(135, 10)
		self.cb[2].Location = Point(220, 42)
		self.cb[2].DropDownStyle = ComboBoxStyle.DropDownList

		self.gb.append(GroupBox())
		self.gb[1].Text = "Фильтр для стен"
		self.gb[1].Size = Size(573, 100)
		self.gb[1].Location = Point(5, 120)
		self.gb[1].Parent = self

		self.cb.append(ComboBox())
		self.cb[3].Parent = self.gb[1]
		self.cb[3].Size = Size(180, 10)
		self.cb[3].Location = Point(15, 42)
		self.cb[3].DropDownStyle = ComboBoxStyle.DropDownList
		self.cb[3].SelectedIndexChanged += self.CblOnChanged
		self.cb[3].FlatStyle = FlatStyle.Flat


		self.cb.append(ComboBox())
		self.cb[4].Parent = self.gb[1]
		self.cb[4].Size = Size(180, 10)
		self.cb[4].Location = Point(380, 42)
		self.cb[4].DropDownStyle = ComboBoxStyle.DropDownList
		self.cb[4].FlatStyle = FlatStyle.Flat

		self.cb.append(ComboBox())
		self.cb[5].Parent = self.gb[1]
		self.cb[5].Size = Size(135, 10)
		self.cb[5].Location = Point(220, 42)
		self.cb[5].DropDownStyle = ComboBoxStyle.DropDownList

		self.bt = Button(Text = "Запустить")
		self.bt.Parent = self
		self.bt.Location = Point(5, 230)
		self.bt.Click += self.run

		self.bt_web = Button(Text = "?")
		self.bt_web.Parent = self
		self.bt_web.Location = Point(85, 230)
		self.bt_web.Click += self.go_to_help

		for i in self.eq:
			self.cb[5].Items.Add(i)
			self.cb[5].Text = self.eq[0]
			self.cb[2].Items.Add(i)
			self.cb[2].Text = self.eq[0]

		self.cb[3].Items.Add("(Нет)")
		if len(self.elements) != 0:
			for i in self.element_parameters:
				self.cb[3].Items.Add(i)
			for i in self.elementtype_parameters:
				self.cb[3].Items.Add(i)
		self.cb[3].Text = "(Нет)"


		self.cb[0].Items.Add("(Нет)")
		if len(self.rooms) != 0:
			for i in self.room_parameters:
				self.cb[0].Items.Add(i)
		self.cb[0].Text = "(Нет)"

		self.CenterToScreen()

	def getType(self, element):
		try:
			id = element.GetTypeId()
			type = doc.GetElement(id)
			return type
		except:
			try:
				type = element.WallType
				return type
			except:
				try:
					type = element.FloorType
					return type
				except:
					return False

	def logger(self, result):
		try:
			now = datetime.datetime.now()
			filename = "{}-{}-{}_{}-{}-{}_FW_{}_{}.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), self.type_description, revit.username)
			file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
			text = "kpln log\nfile:{}\nversion:{}\nuser:{}\n\nlog:\n{}".format(doc.PathName, revit.version, revit.username, result)
			file.write(text.encode('utf-8'))
			file.close()
		except: pass

	def CblOnChanged(self, sender, event):
		values = []
		v = ""
		if sender == self.cb[0]:
			self.cb[1].Enabled = False
			self.cb[2].Enabled = False
			self.cb[2].Text = "Равно"
			if self.cb[0].Text != "(Нет)":
				self.cb[1].Items.Clear()
				for room in self.rooms:
					v = self.getparameter(self.cb[0].Text, room)
					if not self.inlist(v, values):
						values.append(v)
			else:
				self.cb[1].Items.Clear()
			values.sort()
			if len(values) != 0:
				for i in values:
					self.cb[1].Items.Add(i)
				self.cb[1].Text = values[0]
				self.cb[1].Enabled = True
				self.cb[2].Enabled = True
			else:
				self.cb[1].Items.Clear()
				self.cb[1].Enabled = False
				self.cb[2].Enabled = False
				self.cb[2].Text = "Равно"
		if sender == self.cb[3]:
			self.cb[4].Enabled = False
			self.cb[5].Enabled = False
			self.cb[5].Text = "Равно"
			if self.cb[3].Text != "(Нет)":
				self.cb[4].Items.Clear()
				for element in self.elements:
					v = self.getparameter(self.cb[3].Text, element)
					if str(v) != "" and str(v) != "Null":
						if not self.inlist(v, values):
							values.append(v)
			else:
				self.cb[4].Items.Clear()
			values.sort()
			if len(values) != 0:
				for i in values:
					self.cb[4].Items.Add(i)
				self.cb[4].Text = values[0]
				self.cb[4].Enabled = True
				self.cb[5].Enabled = True
			else:
				self.cb[4].Items.Clear()
				self.cb[4].Enabled = False
				self.cb[5].Enabled = False
				self.cb[5].Text = "Равно"
	
	def in_list(self, W, list):
		for i in list:
			if str(W.Id) == str(i.Id):
				return True
		return False

	def inlist(self, e, list):
		for i in list:
			if e == i:
				return True
		return False

	def neib(self, element):
		lc = element.Location
		neib = []
		for i in range(0,2):
			try:
				neighbours = lc.get_ElementsAtJoin(i)
				for nb in neighbours:
					if str(nb.Id) != str(element.Id):
						if self.element_filter(nb):
							neib.append(nb)
			except: pass
		if len(neib) != 0:
			return neib
		return False

	def element_filter(self, element):
		try:
			picked_par = str(self.cb[3].Text)
			if picked_par.startswith("Типоразмер: "):
				param = picked_par[12:]
				type = self.getType(element)
				if type == False:
					return False
				par = type.LookupParameter(param)
			elif picked_par.startswith("Экземпляр: "):
				param = picked_par[11:]
				par = element.LookupParameter(param)
			elif picked_par == "(Нет)":
				return True
			else:
				return False
			if str(par.StorageType) == "String":
				value = str(self.getparameter(str(self.cb[3].Text), element))
				checker = str(self.cb[4].Text)
			elif str(par.StorageType) == "Double":
				if float(self.cb[1].Text):
					try:
						value = float(self.getparameter(self.cb[0].Text, element))
						checker = float(self.cb[1].Text)
					except:
						value = str(self.getparameter(self.cb[0].Text, element))
						checker = str(self.cb[1].Text)
				else:
					value = str(self.getparameter(self.cb[0].Text, element))
					checker = str(self.cb[1].Text)
			if str(self.cb[5].Text) == "Равно":
				if str(value) == checker:
					return True
				else:
					return False
			elif str(self.cb[5].Text) == "Не равно":
				if value != checker:
					return True
				else:
					return False
			elif str(self.cb[5].Text) == "Больше":
				if value > checker:
					return True
				else:
					return False
			elif str(self.cb[5].Text) == "Меньше":
				if value < checker:
					return True
				else:
					return False
			elif str(self.cb[5].Text) == "Больше или равно":
				if value >= checker:
					return True
				else:
					return False
			elif str(self.cb[5].Text) == "Меньше или равно":
				if value <= checker:
					return True
				else:
					return False
			return False
		except: return False

	def room_filter(self, room):
		try:
			if self.cb[0].Text == "(Нет)":
				return True
			picked_par = str(self.cb[0].Text)
			par = room.LookupParameter(picked_par)
			if str(par.StorageType) == "String":
					value = str(self.getparameter(self.cb[0].Text, room))
					checker = str(self.cb[1].Text)
			elif str(par.StorageType) == "Double":
				if self.cb[0].Text != "(Нет)":
					if float(self.cb[1].Text):
						try:
							value = float(self.getparameter(self.cb[0].Text, room))
							checker = float(self.cb[1].Text)
						except:
							value = str(self.getparameter(self.cb[0].Text, room))
							checker = str(self.cb[1].Text)
					else:
						value = str(self.getparameter(self.cb[0].Text, room))
						checker = str(self.cb[1].Text)
				else:
					return True
			if self.cb[2].Text == "Равно":
				if value == checker:
					return True
				else:
					return False
			elif self.cb[2].Text == "Не равно":
				if value != checker:
					return True
				else:
					return False
			elif self.cb[2].Text == "Больше":
				if value > checker:
					return True
				else:
					return False
			elif self.cb[2].Text == "Меньше":
				if value < checker:
					return True
				else:
					return False
			elif self.cb[2].Text == "Больше или равно":
				if value >= checker:
					return True
				else:
					return False
			elif self.cb[2].Text == "Меньше или равно":
				if value <= checker:
					return True
				else:
					return False
				return True
		except: return False

	def go_to_help(self, sender, args):
		webbrowser.open('https://kpln.kdb24.ru/article/60339/')

	def get_all_neighbourhood(self, element):
		try:
			nlist = []
			for i in range (0, 5):
				coup = self.neib(element)
				if coup:
					for e in coup:
						if not self.in_list(e, nlist): nlist.append(e)
			if len(nlist) != 0:
				return nlist
			return False
		except:
			return False

	def boolean_intersection(self, room, element):
		try:
			SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
			calculator = DB.SpatialElementGeometryCalculator(doc, DB.SpatialElementBoundaryOptions())
			results = calculator.CalculateSpatialElementGeometry(room)
			room_solid = results.GetGeometry()
			w_solid = element.get_Geometry(DB.Options())
			for z in w_solid:
				element_solid = z
				break
			s_intersect = DB.BooleanOperationsUtils.ExecuteBooleanOperation(room_solid, element_solid, DB.BooleanOperationsType.Intersect)
			if abs(s_intersect.Volume) > 0.0000001:
				return True
			else:
				s_union = DB.BooleanOperationsUtils.ExecuteBooleanOperation(room_solid, element_solid, DB.BooleanOperationsType.Union)
				area = abs(room_solid.SurfaceArea + element_solid.SurfaceArea - s_union.SurfaceArea)
				if (area < 0.0000001) or (room_solid.Edges.Size + element_solid.Edges.Size == s_union.Edges.Size):
					return False
				else:
					s_difference = DB.BooleanOperationsUtils.ExecuteBooleanOperation(room_solid, element_solid, DB.BooleanOperationsType.Difference)
					area = abs(room_solid.SurfaceArea + element_solid.SurfaceArea - s_difference.SurfaceArea)
					if (area < 0.0000001) and (room_solid.Edges.Size + element_solid.Edges.Size == s_difference.Edges.Size):
						return False
					else:
						return True
		except: return False

	def get_bounds(self, room):
			calculator = DB.SpatialElementGeometryCalculator(doc)
			options = DB.SpatialElementBoundaryOptions()
			boundloc = DB.AreaVolumeSettings.GetAreaVolumeSettings(doc).GetSpatialElementBoundaryLocation(DB.SpatialElementType.Room)
			options.SpatialElementBoundaryLocation = boundloc
			elist = []
			try:
				results = calculator.CalculateSpatialElementGeometry(room)
				for face in results.GetGeometry().Faces:
					for bface in results.GetBoundaryFaceInfo(face):
						elist.append(doc.GetElement(bface.SpatialBoundaryElement.HostElementId))
				return elist
			except:
				pass	
			return False

	def bound(self, room, element):
		try:
			list = self.get_bounds(room)
			for bound in list:
				if str(bound.Id) == str(element.Id):
					return True
			return False
		except:
			return False

	def bb_intersection(self, element, room):
		try:
			e_solid = element.get_Geometry(DB.Options())
			element_bb = e_solid.GetBoundingBox()
			room_bb = room.ClosedShell.GetBoundingBox()
			element_max = element_bb.Max
			element_min = element_bb.Min
			room_max = room_bb.Max
			room_min = room_bb.Min
			if room_max.Z < element_min.Z or room_min.Z > element_max.Z:
				return False
			elif room_max.Y < element_min.Y or room_min.Y > element_max.Y:
				return False
			elif room_max.X < element_min.X or room_min.X > element_max.X:
				return False
			return True
		except: return False

	def most_frequent(self, List): 
		dict = []
		dict_f = []
		for i in List: 
			if i != "":
				if self.inlist(i, dict):
					for z in range(0, len(dict)):
						if i == dict[z]:
							dict_f[z] += 1
				else:
					dict.append(i)
					dict_f.append(1)
		if len(dict) == 0:
			return False
		max = -1
		max_id = 0
		for i in range(0, len(dict_f)):
			if dict_f[i] > max:
				max = dict_f[i]
				max_id = i
		return List[max_id]

	def collect_values(self, list_of_elements):
		v = [[],[],[],[],[]]
		for element in list_of_elements:
			for i in range(0, len(self.def_params)):
				if element.LookupParameter(self.def_params[i]).AsString() != "":
					v[i].append(element.LookupParameter(self.def_params[i]).AsString())
		return [self.most_frequent(v[0]), self.most_frequent(v[1]), self.most_frequent(v[2]), self.most_frequent(v[3]), self.most_frequent(v[4])]

	def getparameter(self, parameter_name, element):
		try:
			if parameter_name.startswith("Типоразмер: "):
				param =  parameter_name[12:]
				el = self.getType(element)
			elif parameter_name.startswith("Экземпляр: "):
				param =  parameter_name[11:]
				el = element
			else:
				param =  parameter_name
				el = element
			try:
				parameter = el.LookupParameter(param)
				if str(parameter.StorageType) == "String":
					if str(parameter.AsString()) != "None" and str(parameter.AsString()) != "":
						return str(parameter.AsString())
					else: return "(Не существует)"
				elif str(parameter.StorageType) == "Double":
					if "SQUARE" in str(parameter.DisplayUnitType):
						return str(round(parameter.AsDouble() * 0.09290304, 2))
					elif "CUBIC" in str(parameter.DisplayUnitType):
						return str(round(parameter.AsDouble() * 0.0283168, 2))
					else:
						return str(round(parameter.AsDouble() * 304.8, 2))
				else: return "(Не существует)"
			except:
				if str(parameter.StorageType) == "Double":
					return 0.00
				else: return "(Не существует)"
		except: return "(Не существует)"

	def check_parameter_elements(self, parameter_name, elements):
		values = []
		for element in elements:
			parameter = self.getparameter(parameter_name, element)
			if not self.inlist(parameter, values):
				values.append(parameter)
		if len(values) > 1:
			return True
		return False

	def insert_parameters(self, r, e, value):
		try:
			name = str(r.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString())
			parameter_1 = e.LookupParameter(self.def_params[0])
			parameter_1.Set(name)
		except: pass
		try:
			function = str(r.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString())
			parameter_2 = e.LookupParameter(self.def_params[1])
			parameter_2.Set(function)
		except: pass
		try:
			id = str(r.Id)
			parameter_3 = e.LookupParameter(self.def_params[2])
			parameter_3.Set(id)
		except: pass
		try:
			number = str(r.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
			parameter_4 = e.LookupParameter(self.def_params[3])
			parameter_4.Set(number)
		except: pass
		try:
			parameter_6 = e.LookupParameter(self.state_param)
			parameter_6.Set(str(value))
		except: pass
		try:
			section = str(r.LookupParameter("ПОМ_Секция").AsString())
			parameter_5 = e.LookupParameter(self.def_params[4])
			parameter_5.Set(section)
		except: pass

	def check_def_params(self, room, element):
		errors = []
		try:
			test = room.LookupParameter("ПОМ_Секция").AsString()
		except:
			pass
		try:
			test = element.LookupParameter(self.def_params[0]).AsString()
		except:
			errors.append("Ошибка параметра : «{}» (элементы)".format(self.def_params[0]))
		try:
			test = element.LookupParameter(self.def_params[1]).AsString()
		except:
			errors.append("Ошибка параметра : «{}» (элементы)".format(self.def_params[1]))
		try:
			test = element.LookupParameter(self.def_params[2]).AsString()
		except:
			errors.append("Ошибка параметра : «{}» (элементы)".format(self.def_params[2]))
		try:
			test = element.LookupParameter(self.def_params[3]).AsString()
		except:
			errors.append("Ошибка параметра : «{}» (элементы)".format(self.def_params[3]))
		try:
			test = element.LookupParameter(self.def_params[4]).AsString()
		except:
			errors.append("Ошибка параметра : «{}» (элементы)".format(self.def_params[4]))
		try:
			test = element.LookupParameter(self.state_param).AsString()
		except:
			errors.append("Ошибка параметра : «{}» (элементы)".format(self.state_param))
		return errors

	def sort_by_distance(self, point, list_of_rooms, list_of_room_points):
		distance = []
		distance_sorted = []
		sorted_rooms = []
		for p in list_of_room_points:
			distance.append(math.sqrt((point.X - p.X)**2 + (point.Y - p.Y)**2 + (point.Z - p.Z)**2))
			distance_sorted.append(math.sqrt((point.X - p.X)**2 + (point.Y - p.Y)**2 + (point.Z - p.Z)**2))
		distance_sorted.sort()
		for ds in distance_sorted:
			for i in range(0, len(distance)):
				if ds == distance[i]:
					sorted_rooms.append(list_of_rooms[i])
					break
		return sorted_rooms

	def no_recalculation(self, element):
		try:
			string = element.LookupParameter(self.def_params[2]).AsString()
			if string == None or string == "": return False
			return True
		except:
			return False

	def run(self, sender, args):
		self.Hide()
		commands = [CommandLink('Да', return_value = True), CommandLink('Нет', return_value = False)]
		dialog = TaskDialog('Пересчитывать элементы с заполненными значениями?',
							title = "Расчет отделки",
							title_prefix = False,
							commands = commands,
							footer ='См. правила работы с отделкой',
							show_close = False)
		s_result = dialog.show()
		self.Show()
		if len(self.elements) != 0 and len(self.rooms) != 0:
			self.elements_passed = []
			self.elements_passed_point = []
			self.rooms_passed = []
			self.rooms_passed_point = []
			self.maximun = 0
			self.cb[0].Enabled = False
			self.cb[1].Enabled = False
			self.cb[2].Enabled = False
			self.cb[3].Enabled = False
			self.cb[4].Enabled = False
			self.cb[5].Enabled = False
			self.bt_web.Enabled = False
			self.bt.Enabled = False
			log = ""
			result = self.check_def_params(self.rooms[0], self.elements[0])
			if len(result) == 0:
				with forms.ProgressBar(title='Initialization... | Detected {value} of {max_value}', step = 100) as pb:
					pb.title='Collecting room data... | Detected {value} of {max_value}'
					for room in self.rooms:
						if room.Area > 0.00:
							if self.room_filter(room):
								room_bbox = room.ClosedShell.GetBoundingBox()
								point = DB.XYZ((room_bbox.Max.X + room_bbox.Min.X)/2, (room_bbox.Max.Y + room_bbox.Min.Y)/2, (room_bbox.Max.Z + room_bbox.Min.Z)/2)
								self.rooms_passed.append(room)
								self.rooms_passed_point.append(point)
					pb.step = 10
					pb.title='Collecting element data... | Detected {value} of {max_value}'
					with db.Transaction(name = "reset"):
						for element in self.elements:
							if s_result == False:
								if not self.no_recalculation(element):
									if self.element_filter(element):
										try:
											e_solid = element.get_Geometry(DB.Options())
											element_bb = e_solid.GetBoundingBox()
											point = DB.XYZ((element_bb.Max.X + element_bb.Min.X)/2, (element_bb.Max.Y + element_bb.Min.Y)/2, (element_bb.Max.Z + element_bb.Min.Z)/2)
											self.elements_passed.append(element)
											self.elements_passed_point.append(point)
											self.maximun += 1
											pb.update_progress(self.detected, max_value = self.maximun)
										except: pass
								else:
									pass
							else:
								if self.element_filter(element):
									try:
										e_solid = element.get_Geometry(DB.Options())
										element_bb = e_solid.GetBoundingBox()
										point = DB.XYZ((element_bb.Max.X + element_bb.Min.X)/2, (element_bb.Max.Y + element_bb.Min.Y)/2, (element_bb.Max.Z + element_bb.Min.Z)/2)
										self.elements_passed.append(element)
										self.elements_passed_point.append(point)
										self.maximun += 1
										pb.update_progress(self.detected, max_value = self.maximun)
										for x in range(0, len(self.def_params)):
											try:
												parameter_1 = element.LookupParameter(self.def_params[x])
												parameter_1.Set("")
											except:
												pass
									except: pass

					with db.Transaction(name = "step_01"):
						if len(self.rooms_passed) != 0 and len(self.elements_passed) != 0:
							pb.title='Intersection & Bound Method | Detected {value} of {max_value}'
							for i in range(0, len(self.elements_passed)):
								element = self.elements_passed[i]
								e_point = self.elements_passed_point[i]
								rooms_passed_sorted = self.sort_by_distance(e_point, self.rooms_passed, self.rooms_passed_point)
								for room in rooms_passed_sorted:
									if self.bb_intersection(element, room):
										if self.boolean_intersection(room, element):
											self.insert_parameters(r = room, e = element, value = "intersection")
											self.detected += 1
											pb.update_progress(self.detected, max_value = self.maximun)
											break
										else:
											if self.bound(room, element):
												self.insert_parameters(r = room, e = element, value = "bound")
												self.detected += 1
												pb.update_progress(self.detected, max_value = self.maximun)
												break
											else:
												pass
									else: pass
					if self.type_description == "WALLS":
						with db.Transaction(name = "step_02"):
							if len(self.rooms_passed) != 0 and len(self.elements_passed) != 0:
								pb.title='Chain Method#1 | Detected {value} of {max_value}'
								for i in range(0, 2):
									self.unresult_elements = []
									for element in self.elements_passed:
										if str(element.LookupParameter(self.def_params[0]).AsString()) == "":
											self.unresult_elements.append(element)
									if len(self.unresult_elements) != 0:
										for element in self.unresult_elements:
											list = self.get_all_neighbourhood(element)
											if list:
												values = self.collect_values(list)
												if values:
													for i in range(0, len(values)):
														if values[i] != "":
															try:
																parameter = element.LookupParameter(self.def_params[i])
																parameter.Set(values[i])
															except: pass
													try:
														parameter_5 = element.LookupParameter(self.state_param)
														parameter_5.Set("chain")
													except: pass
					with db.Transaction(name = "step_03"):
						self.unresult_elements = []
						pb.title='Bounding Box Method | Detected {value} of {max_value}'
						for element in self.elements_passed:
							try:
								if element.LookupParameter(self.def_params[0]).AsString() == "":
									self.unresult_elements.append(element)
							except: pass
						if len(self.unresult_elements) != 0: 
							for element in self.unresult_elements:
								try:
									e_solid = element.get_Geometry(DB.Options())
									element_bb = e_solid.GetBoundingBox()
									e_point = DB.XYZ((element_bb.Max.X + element_bb.Min.X)/2, (element_bb.Max.Y + element_bb.Min.Y)/2, (element_bb.Max.Z + element_bb.Min.Z)/2)
									rooms_passed_sorted = self.sort_by_distance(e_point, self.rooms_passed, self.rooms_passed_point)
									for room in rooms_passed_sorted:
										if self.bb_intersection(element, room):
											self.insert_parameters(r = room, e = element, value = "bbox")
											self.detected += 1
											pb.update_progress(self.detected, max_value = self.maximun)
											break
								except: pass
						self.logger("Detected {}/{}".format(str(self.detected), str(self.maximun)))
						self.TopMost = False
						if self.type_description == "WALLS":
							if len(self.rooms_passed) != 0 and len(self.elements_passed) != 0:
								pb.title='Chain Method#2 | Detected {value} of {max_value}'
								for i in range(0, 2):
									self.unresult_elements = []
									for element in self.elements_passed:
										if str(element.LookupParameter(self.def_params[0]).AsString()) == "":
											self.unresult_elements.append(element)
									if len(self.unresult_elements) != 0:
										for element in self.unresult_elements:
											list = self.get_all_neighbourhood(element)
											if list:
												values = self.collect_values(list)
												if values:
													for i in range(0, len(values)):
														if values[i] != "":
															try:
																parameter = element.LookupParameter(self.def_params[i])
																parameter.Set(values[i])
															except: pass
													try:
														parameter_5 = element.LookupParameter(self.state_param)
														parameter_5.Set("chain")
													except: pass
					self.TopMost = False
					self.logger("Detected {}/{}".format(str(self.detected), str(self.maximun)))
					Alert("Detected {}/{}".format(str(self.detected), str(self.maximun)), title="Отделка", header = "Завершено!")
					self.Close()
			else:
				for i in range(0, len(result)):
					if i == 0:
						log += "{}: {}".format(str(i+1), result[i])
					else:
						log += "\n\n{}: {}".format(str(i+1), result[i])
				self.TopMost = False
				Alert(log, title="Отделка", header = "Необходимо исправить следующие ошибки:")
				self.Close()
		else:
			self.TopMost = False
			Alert("Отсутствуют элементы для расчета!", title="Отделка", header = "Необходимо исправить следующие ошибки:")
			self.Close()
try:
	form = CreateWindow()
	Application.Run(form)
except:
	Alert("Ошибка инициализации!", title="Отделка", header = "Необходимо исправить следующие ошибки:")