# -*- coding: utf-8 -*-


__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Двери - Код"
__doc__ = 'Заполнение параметра "СИТИ_Тип помещения", "СИТИ_Номер помещения", или "СИТИ_Номер квартиры" в дверях по связанному помещению' \


import math
import clr
clr.AddReference('RevitAPI')
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import DB, UI
from pyrevit import revit
from System.Collections.Generic import *
from System import Guid


out = script.get_output()
typeCodeParam = Guid('0c7c200b-7670-479a-9e8c-2ca10acf8d96')
apartNumbParam = Guid('c36396ea-858a-4ec8-8487-0d3e99fbacbc')
roomNumbParam = Guid('93f37fc2-cb70-40ef-9310-594981561474')
userPriorityCodeList = ["6", "3", "5"]


class DoorFinish():
    def __init__(self, element):
        self.Element = element
        self.Rooms = []
        self.Mark = None

    def ApplyValue(self):
        self.Element.LookupParameter("СИТИ_Тип помещения").Set(self.Mark.split("~")[0])
        if self.Mark.split("~")[1] != "0" and self.Mark.split("~")[1] != "None"\
                and self.Mark.split("~")[1]:
            self.Element.LookupParameter("СИТИ_Номер квартиры").Set(self.Mark.split("~")[1])
            self.Element.LookupParameter("СИТИ_Номер помещения").Set(" ")
        else:
            self.Element.LookupParameter("СИТИ_Номер квартиры").Set(" ")
            self.Element.LookupParameter("СИТИ_Номер помещения").Set(self.Mark.split("~")[2])

    def AppendRoom(self, r, bool):
        self.Rooms.append(RoomFinish(r, bool))
        priority = 0
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
        if room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower() == "прихожая"\
            or room.get_Parameter(typeCodeParam).AsString() in userPriorityCodeList:
            self.Priority = 0
        elif isOuter:
            self.Priority = 1
        else:
            self.Priority = 2
        # try:
        #     self.Mark = str(room.LookupParameter("ADSK_Тип помещения").AsValueString())
        # except AttributeError:
        self.Mark = str(room.LookupParameter("СИТИ_Тип помещения").AsString())
        self.Mark += "~"
        # try:
        #     self.Mark += str(room.LookupParameter("КВ_Номер").AsString())
        # except AttributeError:
        self.Mark += str(room.LookupParameter("СИТИ_Номер квартиры").AsString())
        self.Mark += "~"
        # try:
        #     self.Mark += str(room.LookupParameter("ПОМ_Номер_Дополнительный").AsString())
        # except AttributeError:
        self.Mark += str(room.LookupParameter("СИТИ_Номер помещения").AsString())
        """ if room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString() == "Квартира":
            self.Mark += str(room.LookupParameter("КВ_Номер").AsString())
            self.Room = room
        else:
            self.Mark += str(room.LookupParameter("ПОМ_Номер_Дополнительный").AsString())
            self.Room = room """


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


def get_points(vector, distance, location, height, isDoor):
    if isDoor:
        points = []
        x = vector.X
        y = vector.Y
        d = math.sqrt(x*x + y*y)
        x_move = 0
        y_move = 0
        try:
            x_move = distance * x / d
        except :
            pass
        try:
            y_move = distance * y / d
        except :
            pass
        points.append(DB.XYZ(location.X + x_move, location.Y + y_move, location.Z + height))
        points.append(DB.XYZ(location.X + x_move, location.Y + y_move, location.Z + height/2))
        points.append(DB.XYZ(location.X + x_move, location.Y + y_move, location.Z))
        return points
    else:
        points = []
        points.append(DB.XYZ(location.X + 1, location.Y + 1, location.Z))
        points.append(DB.XYZ(location.X + 1, location.Y - 1, location.Z))
        points.append(DB.XYZ(location.X - 1, location.Y + 1, location.Z))
        points.append(DB.XYZ(location.X - 1, location.Y - 1, location.Z))
        return points


def intersect(room, curve):
    SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
    calculator = DB.SpatialElementGeometryCalculator(doc, DB.SpatialElementBoundaryOptions())
    results = calculator.CalculateSpatialElementGeometry(room)
    room_solid = results.GetGeometry()
    options = DB.SolidCurveIntersectionOptions()
    options.ResultType = DB.SolidCurveIntersectionMode.CurveSegmentsInside
    result = room_solid.IntersectWithCurve(curve, options)
    return result.SegmentCount


def rooms_sorted_by_distance(point, rooms, rooms_points):
    rooms_sorted = []
    rooms_points_sorted = []
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
collector_doors = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_Doors)
collector_openings = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_Furniture)
collector_windows = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).OfCategory(DB.BuiltInCategory.OST_Windows)
docName = doc.Title

for door in collector_doors:
    try:
        door_name = door.Symbol.FamilyName
        if door_name.startswith("100_Дверь"):
            type = doc.GetElement(door.GetTypeId())
            height = type.get_Parameter(DB.BuiltInParameter.DOOR_HEIGHT).AsDouble()
            if height == 0:
                height = type.get_Parameter(DB.BuiltInParameter.GENERIC_HEIGHT).AsDouble()
            elements_all.append([door, door.Location.Point, door.FacingOrientation, height, True])
        elif door_name.startswith("100_Дверка")\
                or door_name.startswith("175_Дверка")\
                or door_name.startswith("100_Гермодверь"):
            type = doc.GetElement(door.GetTypeId())
            height = type.get_Parameter(DB.BuiltInParameter.DOOR_HEIGHT).AsDouble()
            if height == 0:
                height = type.get_Parameter(DB.BuiltInParameter.GENERIC_HEIGHT).AsDouble()
            elements_all.append([door, door.Location.Point, door.FacingOrientation, height, False])
    except: pass

for window in collector_windows:
    try:
        window_name = window.Symbol.FamilyName
        if window_name.startswith("110_Блок"):
            type = doc.GetElement(window.GetTypeId())
            height = type.get_Parameter(DB.BuiltInParameter.WINDOW_HEIGHT).AsDouble()
            if height == 0:
                height = type.get_Parameter(DB.BuiltInParameter.GENERIC_HEIGHT).AsDouble()
            elements_all.append([window, window.Location.Point, window.FacingOrientation, height, True])
    except: pass
for mep_opening in collector_openings:
    try:
        mep_opening_name = mep_opening.Symbol.FamilyName
        if mep_opening_name == "175_Дверка шкафов инженерных (сборный)":
            height = mep_opening.LookupParameter("Высота").AsDouble()
            elements_all.append([mep_opening, mep_opening.Location.Point, mep_opening.FacingOrientation, height, False])
    except: pass

for room in collector_rooms:
    rooms_all.append(room)
    try:
        rooms_points_all.append(room.Location.Point)
    except AttributeError:
        out.print_md("Сначала удали неразмещенные помещения!")
        script.exit()

with db.Transaction(name="Write value"):
    for element_set in elements_all:
        element = element_set[0]
        DF = DoorFinish(element)
        location = element_set[1]
        orient = element_set[2]
        height = element_set[3]
        value = 1
        if not element_set[4]:
            value = 0.2
        Inner_points = get_points(orient.Negate(), value, location, height, element_set[4])
        Outer_points = get_points(orient, value, location, height, element_set[4])
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
            try:
                if room.IsPointInRoom(Inner_points[3]):
                    DF.AppendRoom(room, False)
                    break
            except:
                pass
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
                if room.IsPointInRoom(Outer_points[3]):
                    DF.AppendRoom(room, True)
                break
            except:
                pass
        try:
            DF.ApplyValue()
        except Exception as exc:
            el_info = "{}: ({})".format(
                element.Symbol.FamilyName, revit.query.get_name(element.Symbol)
            )
            print(
                "[WARNING] Не расчитано для элемента: {}\n{}\n{}".
                format(out.linkify(element.Id), el_info, exc.ToString())
            )
ui.forms.Alert("Готово!", title="СИТИ_Параметризация")
