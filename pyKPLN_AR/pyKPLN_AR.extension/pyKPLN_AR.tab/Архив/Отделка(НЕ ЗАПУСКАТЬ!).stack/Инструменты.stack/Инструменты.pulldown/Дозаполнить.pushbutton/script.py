# -*- coding: utf-8 -*-
"""
Link_By_Params

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Привязать по id/номеру"
__doc__ = 'Определение принадлежности элементов к помещениям.\n\n' \
          'Для работы необходимо заполнить один из параметров:\n' \
		  '«О_Номер помещения» - для корректности необходимо избавиться от дубликатов в номерах помещений\n ' \
		  '«О_Id помещения» - самый надежный способ\n ' \


"""
Архитекурное бюро KPLN

"""
from pyrevit.framework import clr
import math
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
import datetime
from rpw.ui.forms import Alert

def logger():
	try:
		now = datetime.datetime.now()
		filename = "{}-{}-{}_{}-{}-{}_FWMANUALADD_{}_{}.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), self.type_description, revit.username)
		file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
		text = "kpln log\nfile:{}\nversion:{}\nuser:{}".format(doc.PathName, revit.version, revit.username)
		file.write(text.encode('utf-8'))
		file.close()
	except: pass

def insert_parameters(room, element, value):
	name = str(room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString())
	parameter_1 = element.LookupParameter(params[0])
	parameter_1.Set(name)
	function = str(room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString())
	parameter_2 = element.LookupParameter(params[1])
	parameter_2.Set(function)
	id = str(room.Id)
	parameter_3 = element.LookupParameter(params[2])
	parameter_3.Set(id)
	number = str(room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
	parameter_4 = element.LookupParameter(params[3])
	parameter_4.Set(number)
	parameter_6 = element.LookupParameter("О_Тип")
	parameter_6.Set(value)
	try:
		section = str(room.LookupParameter("ПОМ_Секция").AsString())
		parameter_5 = element.LookupParameter(self.def_params[4])
		parameter_5.Set(section)
	except: pass

def get_method(element):
	name = "none"
	calculate = False
	try:
		name = str(element.LookupParameter(params[0]).AsString())
		if name == "" or name.lower() == "none": calculate = True
	except: calculate = True
	department = "none"
	try:
		department = str(element.LookupParameter(params[1]).AsString())
		if department == "" or department.lower() == "none": calculate = True
	except: calculate = True
	id = "none"
	try:
		id = str(element.LookupParameter(params[2]).AsString())
		if ide == "" or id.lower() == "none": calculate = True
	except: calculate = True
	number = "none"
	try:
		number = str(element.LookupParameter(params[3]).AsString())
		if number == "" or number.lower() == "none": calculate = True
	except: calculate = True
	section = "none"
	if calculate:
		if id.lower() != "none" and id != "":
			return "Id_{}".format(id)
		elif number.lower() != "none" and number != "":
			return "number_{}".format(number)
		else: return "none"

def sort_by_distance(point, list_of_rooms, list_of_room_points):
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

rooms_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
walls_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
floors_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
ceilings_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()

params = ["О_Имя помещения", "О_Назначение помещения", "О_Id помещения", "О_Номер помещения", "О_Номер секции"]

elements = []
rooms = []
rooms_points = []
num = 0
succeed = 0

for e in walls_collector:
	elements.append(e)
for e in floors_collector:
	elements.append(e)
for e in ceilings_collector:
	elements.append(e)

for room in rooms_collector:
	room_bbox = room.ClosedShell.GetBoundingBox()
	rooms.append(room)
	rooms_points.append(DB.XYZ((room_bbox.Max.X + room_bbox.Min.X)/2, (room_bbox.Max.Y + room_bbox.Min.Y)/2, (room_bbox.Max.Z + room_bbox.Min.Z)/2))

with db.Transaction(name = "run"):
	with forms.ProgressBar(title='Анализ',indeterminate=True ,cancellable=False, step=10) as pb:
		for element in elements:
			method = get_method(element)
			if method.startswith("Id_"):
				num += 1
				pb.update_progress(succeed, max_value = num)
				id = method.split("_")[1]
				try:
					room = doc.GetElement(DB.ElementId(int(id)))
					insert_parameters(room, element, "manual_id")
					succeed += 1
					pb.update_progress(succeed, max_value = num)
				except Exception as e: print(id + "\n" + str(e) + "\n\n")
			if method.startswith("number_"):
				e_solid = element.get_Geometry(DB.Options())
				element_bb = e_solid.GetBoundingBox()
				point = DB.XYZ((element_bb.Max.X + element_bb.Min.X)/2, (element_bb.Max.Y + element_bb.Min.Y)/2, (element_bb.Max.Z + element_bb.Min.Z)/2)
				sorted_rooms = sort_by_distance(point, rooms, rooms_points)
				num += 1
				pb.update_progress(succeed, max_value = num)
				number = method.split("_")[1]
				for room in sorted_rooms:
					try:
						room_number = str(room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
						if number == room_number:
							insert_parameters(room, element, "manual_number")
							succeed += 1
							pb.update_progress(succeed, max_value = num)
					except: pass
if num != 0:
	if num == succeed:
		Alert("Просчитано : {}".format(str(succeed)), title="Отделка", header = "Привязка по частично заполненным параметрам")
	else:
		Alert("Просчитано : {} / {}".format(str(succeed), str(num)), title="Отделка", header = "Привязка по частично заполненным параметрам")
logger()