# -*- coding: utf-8 -*-
"""
DIV_WINDOWS_FINISHIG

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Отливы/Откосы - Код"
__doc__ = 'Заполнение параметра "СИТИ_Тип помещения", "СИТИ_Номер помещения", или "СИТИ_Номер квартиры" в окнных элементах (отливы и откосы) по связанному помещению' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr
from rpw import doc, DB, UI, db, revit
from pyrevit import script
from pyrevit import DB, UI
from pyrevit import revit
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *


out = script.get_output()


class WindowFinish():
    def __init__(self, element):
        self.Element = element
        self.Rooms = []
        self.Mark = None

    def ApplyValue(self):
        for id in self.Element.GetSubComponentIds():
            sub_element = doc.GetElement(id)
            if sub_element.Name.startswith("175_")\
                    or sub_element.Symbol.FamilyName.startswith("175_"):
                sub_element.LookupParameter("СИТИ_Тип помещения").Set(self.Mark.split(".")[0])
                if self.Mark.split(".")[1] != "0" and\
                        self.Mark.split(".")[1] != "None" and\
                        len(self.Mark.split(".")[1]) > 0:
                    sub_element.LookupParameter("СИТИ_Номер квартиры").Set(self.Mark.split(".")[1])
                    sub_element.LookupParameter("СИТИ_Номер помещения").Set(" ")
                else:
                    sub_element.LookupParameter("СИТИ_Номер квартиры").Set(" ")
                    sub_element.LookupParameter("СИТИ_Номер помещения").Set(self.Mark.split(".")[2]+"."+self.Mark.split(".")[3])
                """ sub_element.LookupParameter("ДИВ_Код помещения").Set(self.Mark.split(".")[0])
                sub_element.LookupParameter("ДИВ_Номер помещения").Set(self.Mark.split(".")[1]) """

    def AppendRoom(self, r, bool):
        self.Rooms.append(RoomFinish(r, bool))
        priority = 3
        pickedRoom = None
        for rf in self.Rooms:
            if pickedRoom is None or rf.Priority < priority:
                pickedRoom = rf
                priority = rf.Priority
                self.Mark = rf.Mark


class RoomFinish():
    def __init__(self, room, isOuter):
        self.Mark = None
        self.Room = None
        self.Priority = 0
        try:
            roomADSK_Type = room.LookupParameter("ADSK_Тип помещения").AsValueString()
        except AttributeError:
            roomADSK_Type = room.LookupParameter("СИТИ_Тип помещения").AsString()
        if room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower() == "прихожая":
            self.Priority = 0
        elif roomADSK_Type == "2" or roomADSK_Type == "3":
            self.Priority = 0
        elif isOuter:
            self.Priority = 1
        else:
            self.Priority = 2
        self.Mark = roomADSK_Type
        self.Mark += "."
        try:
            self.Mark += str(room.LookupParameter("КВ_Номер").AsString())
        except AttributeError:
            self.Mark += str(room.LookupParameter("СИТИ_Номер квартиры").AsString())
        self.Mark += "."
        try:
            self.Mark += str(room.LookupParameter("ПОМ_Номер_Дополнительный").AsString())
        except AttributeError:
            self.Mark += str(room.LookupParameter("СИТИ_Номер помещения").AsString())
        """ self.Mark += "."
        if room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString() == "Квартира":
            self.Mark += str(room.LookupParameter("КВ_Номер").AsString())
            self.Room = room
        else:
            self.Mark += str(room.LookupParameter("ПОМ_Номер помещения").AsString())
            self.Room = room
 """


def normilise_angle(orientaion):
    angle = math.atan2(orientaion.X, orientaion.Y) / 0.01745329251994
    if angle < 0:
        angle += 360
    return angle


def vector_turn(point, clockwise=True):
    x = point.X
    y = point.Y
    z = point.Z
    if clockwise:
        return DB.XYZ(y,-x,z)
    else:
        return DB.XYZ(-y,x,z)


def intersect(room, curve):
    calculator = DB.SpatialElementGeometryCalculator(doc, DB.SpatialElementBoundaryOptions())
    results = calculator.CalculateSpatialElementGeometry(room)
    room_solid = results.GetGeometry()
    options = DB.SolidCurveIntersectionOptions()
    options.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsInside
    result = room_solid.IntersectWithCurve(curve, options)
    return result.SegmentCount


def get_points(vector, distance, location, height):
    points = []
    x = vector.X
    y = vector.Y
    d = math.sqrt(x*x + y*y)
    x_move = distance * x / d
    y_move = distance * y / d
    points.append(DB.XYZ(location.X + x_move, location.Y + y_move, location.Z + height))
    points.append(DB.XYZ(location.X + x_move, location.Y + y_move, location.Z + height/2))
    points.append(DB.XYZ(location.X + x_move, location.Y + y_move, location.Z))
    return points


def rooms_sorted_by_distance(point, rooms, rooms_points):
    rooms_sorted = []
    distance = []
    distance_sorted = []
    for p in rooms_points:
        distance.append(math.sqrt((point.X - p.X)**2 + (point.Y - p.Y)**2 + (point.Z - p.Z)**2))
        distance_sorted.append(math.sqrt((point.X - p.X)**2 + (point.Y - p.Y)**2 + (point.Z - p.Z)**2))
    distance_sorted.sort()
    for ds in distance_sorted:
        for i in range(0, len(distance)):
            if ds == distance[i]:
                rooms_sorted.append(rooms[i])
                break
    return rooms_sorted


rooms_all = []
rooms_points_all = []
elements_all = []

collector_rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).ToElements()
collector_windows = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_Windows)

for window in collector_windows:
    try:
        try:
            window_name = window.Symbol.FamilyName
            if window_name.startswith("110_") or window_name.startswith("120_"):
                type = doc.GetElement(window.GetTypeId())
                height = type.get_Parameter(DB.BuiltInParameter.WINDOW_HEIGHT).AsDouble()
                if height == 0:
                    height = type.get_Parameter(DB.BuiltInParameter.GENERIC_HEIGHT).AsDouble()
                elements_all.append([window, window.Location.Point, window.FacingOrientation, height])
        except: pass
    except: pass

for room in collector_rooms:
    rooms_all.append(room)
    rooms_points_all.append(room.Location.Point)

with db.Transaction(name="Write value"):
    for element_set in elements_all:
        element = element_set[0]
        DF = WindowFinish(element)
        location = element_set[1]
        orient = element_set[2]
        height = element_set[3]
        Inner_points = get_points(orient.Negate(), 1, location, height)
        Outer_points = get_points(orient, 1, location, height)
        for room in rooms_sorted_by_distance(location, rooms_all, rooms_points_all):
            if room.IsPointInRoom(Inner_points[0]):
                DF.AppendRoom(room, False)
                break
            elif room.IsPointInRoom(Inner_points[1]):
                DF.AppendRoom(room, False)
                break
            elif room.IsPointInRoom(Inner_points[2]):
                DF.AppendRoom(room, False)
                break
        for room in rooms_sorted_by_distance(location, rooms_all, rooms_points_all):
            if room.IsPointInRoom(Outer_points[0]):
                DF.AppendRoom(room, True)
                break
            elif room.IsPointInRoom(Outer_points[1]):
                DF.AppendRoom(room, True)
                break
            elif room.IsPointInRoom(Outer_points[2]):
                DF.AppendRoom(room, True)
                break
        try:
            DF.ApplyValue()
        except:
            type = doc.GetElement(element.GetTypeId())
            el_info = "{}: ({})".format(element.Symbol.FamilyName, revit.query.get_name(element.Symbol))
            print("[WARNING] Не расчитано для элемента: {}\n{}\n...".format(out.linkify(element.Id), el_info))
print("Готово!")