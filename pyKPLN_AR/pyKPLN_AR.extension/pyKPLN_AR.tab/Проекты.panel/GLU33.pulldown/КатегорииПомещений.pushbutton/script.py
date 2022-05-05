# -*- coding: utf-8 -*-
"""
GLU33_ROOM_CAT

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Квартирография [Экспорт]"
__doc__ = 'Заполнение параметров квартирографии (с автоподгрузкой параметров) для подготовки спецификации в отдел продаж.\n\nВажно:\n - Необходимо запускать после проверки/запуска квартирографии;\n - В санузлах должны быть расставлена сантехника.' \

"""
Архитекурное бюро KPLN

"""
import math
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance,\
                              RevitLinkInstance
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit import revit
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert
import datetime


out = script.get_output()
out.set_title("Рассчет квартирографии (export)")


def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False

fop_path = "Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\lib\\GLU33_Export.txt"
parameters_to_load =[["G33Ex_Кв_Площадь_Доп.с/у2", "Area", True],
    ["G33Ex_Кв_Площадь_Комн2", "Area", True],
    ["G33Ex_Кв_Площадь_Балконы_Кол-во", "Text", True],
    ["G33Ex_Кв_Площадь_Лоджия_Кол-во", "Text", True],
    ["G33Ex_Кв_Площадь_Комн4", "Area", True],
    ["G33Ex_Выбранный диапазон", "Text", True],
    ["G33Ex_Кв_Площадь_Совмещенная ванная", "Area", True],
    ["G33Ex_Кв_Площадь_Лоджия1_Факт", "Area", True],
    ["G33Ex_Кв_Площадь_Лоджии_К", "Area", True],
    ["G33Ex_Кв_Площадь_Комн1", "Area", True],
    ["G33Ex_Кв_Площадь_Прихожая", "Area", True],
    ["G33Ex_Кв_Площадь_Ванная", "Area", True],
    ["G33Ex_Кв_Площадь_ГардеробнаяКладовая", "Area", True],
    ["G33Ex_Кв_Площадь_Совмещенная ванная2", "Area", True],
    ["G33Ex_Кв_Площадь_Балкон1_Факт", "Area", True],
    ["G33Ex_Кв_Площадь_Душевая", "Area", True],
    ["G33Ex_Кв_Площадь_Комн3", "Area", True],
    ["G33Ex_Кв_Площадь_Лоджия2_Факт", "Area", True],
    ["G33Ex_Кв_Площадь_Балконы_К", "Area", True],
    ["G33Ex_Кв_Площадь_Балкон3_Факт", "Area", True],
    ["G33Ex_Кв_Площадь_Совмещенная душевая", "Area", True],
    ["G33Ex_Кв_Площадь_Балкон2_Факт", "Area", True],
    ["G33Ex_Кв_Площадь_Лоджия3_Факт", "Area", True],
    ["G33Ex_Кв_Площадь_Доп.с/у", "Area", True],
    ["G33Ex_Типовые этажи", "Text", True],
    ["G33Ex_Кв_Площадь_Коридор", "Area", True],
    ["G33Ex_Кв_Площадь_Кухня", "Area", True],
    ["G33Ex_Кв_Площадь_ГардеробнаяКладовая2", "Area", True],
    ["G33Ex_Кв_Площадь_ГардеробнаяКладовая3", "Area", True],
    ["G33Ex_Категория помещения", "Text", True],
    ["G33Ex_Кол-во жилых комнат", "Text", True]]

params_found = []
group = "Квартирография"
common_parameters_file = fop_path
app = doc.Application
originalFile = app.SharedParametersFilename
category_set_rooms = app.Create.NewCategorySet()
insert_cat_rooms = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Rooms)
category_set_rooms.Insert(insert_cat_rooms)
app.SharedParametersFilename = common_parameters_file
SharedParametersFile = app.OpenSharedParameterFile()
#CHECK CURRENT PARAMETERS
map = doc.ParameterBindings
it = map.ForwardIterator()
it.Reset()
while it.MoveNext():
    d_Definition = it.Key
    d_Name = it.Key.Name
    d_Binding = it.Current
    d_catSet = d_Binding.Categories	
    for i in range(0, len(parameters_to_load)):
        if d_Name == parameters_to_load[i][0]:
            params_found.append(d_Name)
any_parameters_loaded = False
if len(parameters_to_load) != len(params_found):
    with db.Transaction(name="Подгрузить недостающие параметры"):
        for dg in SharedParametersFile.Groups:
            if dg.Name == group:
                for ps in parameters_to_load:
                    if not in_list(ps[0], params_found):
                        externalDefinition = dg.Definitions.get_Item(ps[0])
                        newIB = app.Create.NewInstanceBinding(category_set_rooms)
                        doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
                        doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
                        any_parameters_loaded = True

map = doc.ParameterBindings
it = map.ForwardIterator()
it.Reset()
if any_parameters_loaded:
    with db.Transaction(name="Настройка параметров"):
        while it.MoveNext():
            for i in range(0, len(parameters_to_load)):
                if not in_list(parameters_to_load[i][0], params_found):
                    d_Definition = it.Key
                    d_Name = it.Key.Name
                    d_Binding = it.Current
                    if d_Name == parameters_to_load[i][0]:
                        d_Definition.SetAllowVaryBetweenGroups(doc, parameters_to_load[i][2])

def RoomContainsPoint(room, points):
    for pt in points:
        if room.IsPointInRoom(pt):
            return True
    return False

def RoomIsBath(room):
    try:
        dict = ["с/у", "ванная", "душевая", "санузел", "туалет", "уборная"]
        department = room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString().lower()
        name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower()
        for d in dict:
            if d in name and department == "квартира":
                return True
        return False
    except :
        return False


sets = []
for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance).WhereElementIsNotElementType().ToElements():
    transform = link.GetTransform()
    try:
        if link.GetLinkDocument():
            for element in FilteredElementCollector(link.GetLinkDocument()).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements():
                if "ванна" in element.Symbol.FamilyName.lower() or "душ" in element.Symbol.FamilyName.lower() or "унитаз" in element.Symbol.FamilyName.lower():
                    transformed_point = transform.OfPoint(element.Location.Point)
                    points = []
                    points.append(DB.XYZ(transformed_point.X + 0.5, transformed_point.Y + 0.5, transformed_point.Z + 1))
                    points.append(DB.XYZ(transformed_point.X + 0.5, transformed_point.Y - 0.5, transformed_point.Z + 1))
                    points.append(DB.XYZ(transformed_point.X - 0.5, transformed_point.Y + 0.5, transformed_point.Z + 1))
                    points.append(DB.XYZ(transformed_point.X - 0.5, transformed_point.Y - 0.5, transformed_point.Z + 1))
                    sets.append([element, points])
    except Exception as e:
        pass
        print(str(e))

try:
    for element in FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements():
        if "ванна" in element.Symbol.FamilyName.lower() or "душ" in element.Symbol.FamilyName.lower() or "унитаз" in element.Symbol.FamilyName.lower():
            transformed_point = element.Location.Point
            points = []
            points.append(DB.XYZ(transformed_point.X + 0.5, transformed_point.Y + 0.5, transformed_point.Z + 1))
            points.append(DB.XYZ(transformed_point.X + 0.5, transformed_point.Y - 0.5, transformed_point.Z + 1))
            points.append(DB.XYZ(transformed_point.X - 0.5, transformed_point.Y + 0.5, transformed_point.Z + 1))
            points.append(DB.XYZ(transformed_point.X - 0.5, transformed_point.Y - 0.5, transformed_point.Z + 1))
            sets.append([element, points])
except Exception as e:
    print(str(e))


got_errors = False
flat_numbers = []


with db.Transaction(name = "Поиск сантехники"):
    for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
        try:
            flat_number = room.LookupParameter("КВ_Номер").AsString()
            if flat_number != None and flat_number != "0":
                if not flat_number in flat_numbers:
                    flat_numbers.append(flat_number)
        except:
            pass
        if RoomIsBath(room):
            got_toilet = False
            got_bath = False
            got_shower = False
            type = ""
            for set in sets:
                if RoomContainsPoint(room, set[1]):
                    if "ванна" in set[0].Symbol.FamilyName.lower():
                        got_bath = True
                    if "душ" in set[0].Symbol.FamilyName.lower():
                        got_shower = True
                    if "унитаз" in set[0].Symbol.FamilyName.lower():
                        got_toilet = True
            if got_toilet and got_bath and not got_shower:
                type = "Совмещенная ванная"
            if got_toilet and not got_bath and got_shower:
                type = "Совмещенная душевая"
            if got_toilet and not got_bath and not got_shower:
                type = "Доп. С/У"
            if not got_toilet and got_bath and not got_shower:
                type = "Ванная"
            if not got_toilet and not got_bath and got_shower:
                type = "Душевая"
            room.LookupParameter("G33Ex_Категория помещения").Set(type)
            if not got_toilet and not got_bath and not got_shower:
                got_errors = True
                print("Ошибка: В «{1}» ({2}) отсутствует сантехника: {0} - Расчет будет остановлен!".format(out.linkify(room.Id), room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString(), room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()))

flat_numbers.sort()

def GetLivingRooms(rooms):
	result = []
	dict = ["жилая комната", "спальная", "жилая", "гостиная"]
	for room in rooms:
		name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower()
		if name in dict:
			result.append(room)
	return result

def GetBalconies(rooms):
	result = []
	dict = ["балкон"]
	for room in rooms:
		name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower()
		if name in dict:
			result.append(room)
	return result

def GetLodgias(rooms):
	result = []
	dict = ["лоджия"]
	for room in rooms:
		name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower()
		if name in dict:
			result.append(room)
	return result

def GetLivingRoomArea(source_rooms, number):
	rooms = GetLivingRooms(source_rooms)
	try:
		room = rooms[number]
		room.LookupParameter("G33Ex_Категория помещения").Set("Комн." + str(number + 1))
		return room.LookupParameter("ПОМ_Площадь").AsDouble()
	except: pass
	return 0.0

def GetBathRoomArea(source_rooms):
	area = 0.0
	for room in source_rooms:
		if room.LookupParameter("G33Ex_Категория помещения").AsString() == "Ванная":
			area += room.LookupParameter("ПОМ_Площадь").AsDouble()
			room.LookupParameter("G33Ex_Категория помещения").Set("Ванная")
	return area

def GetShowerRoomArea(source_rooms):
	area = 0.0
	for room in source_rooms:
		if room.LookupParameter("G33Ex_Категория помещения").AsString() == "Душевая":
			area += room.LookupParameter("ПОМ_Площадь").AsDouble()
			room.LookupParameter("G33Ex_Категория помещения").Set("Душевая")
	return area

def GetCompBathRoomArea(source_rooms, number):
    rooms = []
    for room in source_rooms:
        roomCatStr = room.LookupParameter("G33Ex_Категория помещения").AsString()
        if roomCatStr:
            if "совмещенная ванная" in roomCatStr.lower():
                rooms.append(room)
    try:
        rooms[number].LookupParameter("G33Ex_Категория помещения").Set("Совмещенная ванная " + str(number + 1))
        return rooms[number].LookupParameter("ПОМ_Площадь").AsDouble()
    except :
        pass
    return 0.0

def GetCompShowerRoomArea(source_rooms):
	area = 0.0
	for room in source_rooms:
		if room.LookupParameter("G33Ex_Категория помещения").AsString() == "Совмещенная душевая":
			try:
			    area += room.LookupParameter("ПОМ_Площадь").AsDouble()
			except :
			    pass
			room.LookupParameter("G33Ex_Категория помещения").Set("Совмещенная душевая")	
	return area

def GetToiletRoomArea(source_rooms, number):
    rooms = []
    for room in source_rooms:
        roomCatStr = room.LookupParameter("G33Ex_Категория помещения").AsString()
        if roomCatStr:
            if "доп. с/у" in roomCatStr.lower():
                rooms.append(room)
    try:
        rooms[number].LookupParameter("G33Ex_Категория помещения").Set("Доп. с/у " + str(number + 1))
        return rooms[number].LookupParameter("ПОМ_Площадь").AsDouble()
    except :
        pass
    return 0.0

def GetWardrobeRoomArea(source_rooms, number):
    rooms = []
    for room in source_rooms:
        roomCatStr = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
        if roomCatStr:
            if "гардероб" in roomCatStr.lower():
                rooms.append(room)
    try:
        rooms[number].LookupParameter("G33Ex_Категория помещения").Set("Гардеробная/кладовая " + str(number + 1))
        return rooms[number].LookupParameter("ПОМ_Площадь").AsDouble()
    except :
        pass
    return 0.0

def GetKitchenRoomArea(source_rooms):
	area = 0.0
	for room in source_rooms:
		if "кухня-ниша" == room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower():
			room.LookupParameter("G33Ex_Категория помещения").Set("Кухня")
			area += room.LookupParameter("ПОМ_Площадь").AsDouble()
		elif "кухня-гостиная" == room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower():
			room.LookupParameter("G33Ex_Категория помещения").Set("Кухня")
			area += room.LookupParameter("ПОМ_Площадь").AsDouble()
		elif "кухня-столовая" == room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower():
			room.LookupParameter("G33Ex_Категория помещения").Set("Кухня")
			area += room.LookupParameter("ПОМ_Площадь").AsDouble()
	return area

def GetCoridorRoomArea(source_rooms):
	area = 0.0
	for room in source_rooms:
		if "коридор" in room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower():
			room.LookupParameter("G33Ex_Категория помещения").Set("Коридор")
			area += room.LookupParameter("ПОМ_Площадь").AsDouble()
	return area

def GetEnterRoomArea(source_rooms):
	area = 0.0
	for room in source_rooms:
		if "прихожая" == room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower():
			room.LookupParameter("G33Ex_Категория помещения").Set("Прихожая")
			area += room.LookupParameter("ПОМ_Площадь").AsDouble()
	return area

def GetLodgiaRoomArea(source_rooms, number, fact):
	rooms = GetLodgias(source_rooms)
	try:
		rooms[number].LookupParameter("G33Ex_Категория помещения").Set("Лоджия " + str(number + 1))
		if fact:
			return rooms[number].LookupParameter("ПОМ_Площадь").AsDouble()
		else:
			return rooms[number].LookupParameter("ПОМ_Площадь_К").AsDouble()
	except :
	    pass	
	return 0.0

def GetBalconyRoomArea(source_rooms, number, fact):
	rooms = GetBalconies(source_rooms)
	try:
		rooms[number].LookupParameter("G33Ex_Категория помещения").Set("Балкон " + str(number + 1))
		if fact:
			return rooms[number].LookupParameter("ПОМ_Площадь").AsDouble()
		else:
			return rooms[number].LookupParameter("ПОМ_Площадь_К").AsDouble()
	except :
	    pass	
	return 0.0

def ApplyValue(rooms, value, name):
	for room in rooms:
		try:
		    room.LookupParameter(name).Set(value)
		except Exception as e:
		    print(str(e))


if not got_errors:
    with db.Transaction(name = "Запись значений"):
        for fn in flat_numbers:
            try:
                rooms = []
                for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
                    a = room.Area
                    try:
                        flat_number = room.LookupParameter("КВ_Номер").AsString()
                        if flat_number == fn and a > 0 and a != None:
                            rooms.append(room)
                    except:
                        continue
                if len(rooms) != 0:
                    v = GetToiletRoomArea(rooms, 1)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Доп.с/у2")
                    v = GetLivingRoomArea(rooms, 1)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Комн2")
                    v = str(len(GetBalconies(rooms)))
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Балконы_Кол-во")
                    v = str(len(GetLodgias(rooms)))
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Лоджия_Кол-во")
                    v = GetLivingRoomArea(rooms, 3)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Комн4")
                    v = GetCompBathRoomArea(rooms, 0)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Совмещенная ванная")
                    v = GetLodgiaRoomArea(rooms, 0, True)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Лоджия1_Факт")
                    v = GetLodgiaRoomArea(rooms, 0, False) + GetLodgiaRoomArea(rooms, 1, False) + GetLodgiaRoomArea(rooms, 2, False) + GetLodgiaRoomArea(rooms, 3, False)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Лоджии_К")
                    v = GetLivingRoomArea(rooms, 0)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Комн1")
                    v = GetEnterRoomArea(rooms)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Прихожая")
                    v = GetBathRoomArea(rooms)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Ванная")
                    v = GetCompBathRoomArea(rooms, 1)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Совмещенная ванная2")
                    v = GetBalconyRoomArea(rooms, 0, True)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Балкон1_Факт")
                    v = GetShowerRoomArea(rooms)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Душевая")
                    v = GetLivingRoomArea(rooms, 2)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Комн3")
                    v = GetLodgiaRoomArea(rooms, 1, True)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Лоджия2_Факт")
                    v = GetBalconyRoomArea(rooms, 0, False) + GetBalconyRoomArea(rooms, 1, False) + GetBalconyRoomArea(rooms, 2, False)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Балконы_К")
                    v = GetBalconyRoomArea(rooms, 2, True)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Балкон3_Факт")
                    v = GetCompShowerRoomArea(rooms)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Совмещенная душевая")
                    v = GetBalconyRoomArea(rooms, 1, True)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Балкон2_Факт")
                    v = GetLodgiaRoomArea(rooms, 2, True)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Лоджия3_Факт")
                    v = GetToiletRoomArea(rooms, 0)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Доп.с/у")
                    v = GetCoridorRoomArea(rooms)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Коридор")
                    v = GetKitchenRoomArea(rooms)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_Кухня")
                    v = GetWardrobeRoomArea(rooms, 0)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_ГардеробнаяКладовая")
                    v = GetWardrobeRoomArea(rooms, 1)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_ГардеробнаяКладовая2")
                    v = GetWardrobeRoomArea(rooms, 2)
                    ApplyValue(rooms, v, "G33Ex_Кв_Площадь_ГардеробнаяКладовая3")
                    v = str(len(GetLivingRooms(rooms)))
                    ApplyValue(rooms, v, "G33Ex_Кол-во жилых комнат")
            except Exception as e:
                print(str(e))