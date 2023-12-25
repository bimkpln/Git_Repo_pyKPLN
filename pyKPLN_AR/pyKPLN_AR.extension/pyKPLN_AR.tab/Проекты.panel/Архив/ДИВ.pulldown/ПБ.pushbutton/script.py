# -*- coding: utf-8 -*-
"""
KPLN:DIV:PAZOGREB

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Пазогребневые блоки"
__doc__ = 'Заполнение параметров "СИТИ_Тип помещения", "СИТИ_Номер помещения", или "СИТИ_Номер квартиры" по связанным помещениям' \

"""
Архитекурное бюро KPLN

"""
from rpw import doc, DB, db
from pyrevit import script
from pyrevit import forms

class Container:
    def __init__(self, box):
        self.BoundingBox = box
        self.Containers = []
        self.Elements = []
        self.Width = box.Max.X - box.Min.X
        x = self.BoundingBox.Min.X;
        y = self.BoundingBox.Min.Y;
        z = self.BoundingBox.Min.Z;
        if self.Width >= 32:
            for multiple_x in [0, 0.5]:
                for multiple_y in [0, 0.5]:
                    for multiple_z in [0, 0.5]:
                        _min = DB.XYZ(x + self.Width * multiple_x, y + self.Width * multiple_y, z + self.Width * multiple_z)
                        _max = DB.XYZ(x + self.Width * 0.5 + self.Width * multiple_x, y + self.Width * 0.5 + self.Width * multiple_y, z + self.Width * 0.5 + self.Width * multiple_z)
                        b_box = DB.BoundingBoxXYZ()
                        b_box.Min = _min
                        b_box.Max = _max
                        self.Containers.append(Container(b_box))

    def Optimize(self):
        optimized_collection = []
        if len(self.Containers) != 0:
            for c in self.Containers:
                if not c.IsEmpty():
                    c.Optimize()
                    optimized_collection.append(c)
        self.Containers = optimized_collection

    def IsEmpty(self):
        if len(self.Elements) != 0:
            return False
        if len(self.Containers) != 0:
            for c in self.Containers:
                if c.IsEmpty(): return True
        return False

    def InsertItem(self, item, box):
        if len(self.Containers) != 0:
            for c in self.Containers:
                c.InsertItem(item, box)
        else:
            if self.boxPassesCondition(box):
                self.Elements.append(item)

    def GetContext(self, box):
        context = []
        if self.boxPassesCondition(box):
            if len(self.Containers) != 0:
                for c in self.Containers:
                    for i in c.GetContext(box):
                        context.append(i)
            else:
                if len(self.Elements) != 0:
                    for i in self.Elements:
                        context.append(i)
        return context

    def boxPassesCondition(self, box):
        if box.Min.X > self.BoundingBox.Max.X + 3: return False
        if box.Min.Y > self.BoundingBox.Max.Y + 3: return False
        if box.Min.Z > self.BoundingBox.Max.Z + 3: return False
        if box.Max.X < self.BoundingBox.Min.X - 3: return False
        if box.Max.Y < self.BoundingBox.Min.Y - 3: return False
        if box.Max.Z < self.BoundingBox.Min.Z - 3: return False
        return True

def get_stack_by_id(id):
    for stack in defined_stacks:
        if id == stack[0].Id.IntegerValue:
            return stack
    return None

def get_neighbourhood(wall):
    location = wall.Location
    elements = []
    for i in [0, 1]:
        try:
            elements_at_join = location.get_ElementsAtJoin(i)
            for wall in elements_at_join:
                if wall.Id.IntegerValue != wall.Id.IntegerValue and "03_ВН(Паз" in wall.Name:
                    elements.append(wall)
        except: pass
    return elements

matrix = None
with forms.ProgressBar(title='Инициализация . . .', step = 10) as pb:
    global_box = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement().get_BoundingBox(None)
    step = 0
    _max = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements().Count
    for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
        step+=1
        pb.update_progress(step, max_value = _max)
        box = room.get_BoundingBox(None)
        if global_box.Min.X > box.Min.X: global_box.Min = DB.XYZ(box.Min.X, global_box.Min.Y, global_box.Min.Z)
        if global_box.Min.Y > box.Min.Y: global_box.Min = DB.XYZ(global_box.Min.X, box.Min.Y, global_box.Min.Z)
        if global_box.Min.Z > box.Min.Z: global_box.Min = DB.XYZ(global_box.Min.X, global_box.Min.Y, box.Min.Z)
        if global_box.Max.X < box.Max.X: global_box.Max = DB.XYZ(box.Max.X, global_box.Max.Y, global_box.Max.Z)
        if global_box.Max.Y < box.Max.Y: global_box.Max = DB.XYZ(global_box.Max.X, box.Max.Y, global_box.Max.Z)
        if global_box.Max.Z < box.Max.Z: global_box.Max = DB.XYZ(global_box.Max.X, global_box.Max.Y, box.Max.Z)
    max_length = max([global_box.Max.X - global_box.Min.X, global_box.Max.Y - global_box.Min.Y, global_box.Max.Z - global_box.Min.Z])
    rounded_length = (round(max_length / 32) + 1) * 32
    global_center = DB.XYZ((global_box.Max.X + global_box.Min.X)/2, (global_box.Max.Y + global_box.Min.Y)/2, (global_box.Max.Z + global_box.Min.Z)/2)
    global_box.Min = DB.XYZ(global_center.X - rounded_length / 2, global_center.Y - rounded_length / 2, global_center.Z - rounded_length / 2)
    global_box.Max = DB.XYZ(global_center.X + rounded_length / 2, global_center.Y + rounded_length / 2, global_center.Z + rounded_length / 2)
    pb.title = "Подготовка помещений . . ."
    matrix = Container(global_box)
    step = 0
    for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
        step+=1
        pb.update_progress(step, max_value = _max)
        matrix.InsertItem(room, room.get_BoundingBox(None))
    matrix.Optimize()

defined_stacks = []
_max = 0
for wall in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements():
    if "03_ВН(Паз" in wall.Name:
        _max += 1
step = 0
out = script.get_output()
got_undefined = False
with forms.ProgressBar(title='Привязка элементов . . .', step = 10) as pb:
    for wall in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements():
        if "03_ВН(Паз" in wall.Name:
            step += 1
            pb.update_progress(step, max_value = _max)
            #pb.title = "Прямая привязка: {} <{}>".format(wall.Name, wall.Id.ToString())
            joined_rooms = []
            box = wall.get_BoundingBox(None)
            curve = wall.Location.Curve
            line = DB.Line.CreateBound(curve.GetEndPoint(0), curve.GetEndPoint(1))
            direction = line.Direction
            clock_wise_direction = DB.XYZ(direction.Y, -direction.X, direction.Z)
            counter_clock_wise_direction = DB.XYZ(-direction.Y, direction.X, direction.Z)
            center_points = []
            start_points = []
            end_points = []
            center_points.append(DB.XYZ((curve.GetEndPoint(0).X + curve.GetEndPoint(1).X)/2,(curve.GetEndPoint(0).Y + curve.GetEndPoint(1).Y)/2, (box.Min.Z + box.Max.Z)/2))
            center_points.append(DB.XYZ((curve.GetEndPoint(0).X + curve.GetEndPoint(1).X)/2,(curve.GetEndPoint(0).Y + curve.GetEndPoint(1).Y)/2, box.Min.Z))
            center_points.append(DB.XYZ((curve.GetEndPoint(0).X + curve.GetEndPoint(1).X)/2,(curve.GetEndPoint(0).Y + curve.GetEndPoint(1).Y)/2, box.Max.Z))
            start_points.append(DB.XYZ(curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y, (box.Min.Z + box.Max.Z)/2))
            start_points.append(DB.XYZ(curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y, box.Min.Z))
            start_points.append(DB.XYZ(curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y, box.Max.Z))
            end_points.append(DB.XYZ(curve.GetEndPoint(1).X, curve.GetEndPoint(1).Y, (box.Min.Z + box.Max.Z)/2))
            end_points.append(DB.XYZ(curve.GetEndPoint(1).X, curve.GetEndPoint(1).Y, box.Min.Z))
            end_points.append(DB.XYZ(curve.GetEndPoint(1).X, curve.GetEndPoint(1).Y, box.Max.Z))
            room_ids = []
            context = matrix.GetContext(box)
            for center_point in center_points:
                for i in [0, 0.164042, 0.82020, 1.64042]:
                    pt_clw = center_point + clock_wise_direction * 1.64042
                    pt_ctr_clw = center_point + counter_clock_wise_direction * 1.64042
                    for room in context:
                        if room.LevelId.IntegerValue == wall.LevelId.IntegerValue:
                            if room.IsPointInRoom(pt_clw) or room.IsPointInRoom(pt_ctr_clw):
                                if not room.Id.IntegerValue in room_ids:
                                    joined_rooms.append(room)
                                    room_ids.append(room.Id.IntegerValue)
                    if len(joined_rooms) == 2:
                        break
                if len(joined_rooms) == 2:
                    break
            for end_point in end_points:
                for i in [0, 0.164042, 0.82020, 1.64042]:
                    pt = end_point + direction * 1.64042
                    for room in context:
                        if room.LevelId.IntegerValue == wall.LevelId.IntegerValue:
                            if room.IsPointInRoom(pt):
                                if not room.Id.IntegerValue in room_ids:
                                    joined_rooms.append(room)
                                    room_ids.append(room.Id.IntegerValue)
                    if len(joined_rooms) != 0:
                        break
                if len(joined_rooms) != 0:
                    break
            for start_point in start_points:
                for i in [0, 0.164042, 0.82020, 1.64042]:
                    pt = start_point + direction.Negate() * 1.64042
                    for room in context:
                        if room.LevelId.IntegerValue == wall.LevelId.IntegerValue:
                            if room.IsPointInRoom(pt):
                                if not room.Id.IntegerValue in room_ids:
                                    joined_rooms.append(room)
                                    room_ids.append(room.Id.IntegerValue)
                    if len(joined_rooms) != 0:
                        break
                if len(joined_rooms) != 0:
                    break
            if len(joined_rooms) != 0:
                if len(joined_rooms) != 1:
                    defined = False
                    for room in joined_rooms:
                        if room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString().lower() == "квартира":
                            defined_stacks.append([wall, room])
                            defined = True
                            break
                    if not defined:
                        for room in joined_rooms:
                            if room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower() == "пуи":
                                defined_stacks.append([wall, room])
                                defined = True
                                break
                    if not defined:
                        for room in joined_rooms:
                            if room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower() == "колясочная":
                                defined_stacks.append([wall, room])
                                defined = True
                                break
                    if not defined:
                        for room in joined_rooms:
                            if room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower() == "лестничная клетка" or room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower() == "лк" or "лк-" in room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString().lower():
                                defined_stacks.append([wall, room])
                                defined = True
                                break
                    if not defined:
                        defined_stacks.append([wall, joined_rooms[0]])
                else:
                    defined_stacks.append([wall, joined_rooms[0]])
            else:
                got_undefined = True
                defined_stacks.append([wall, None])
step = 0
if got_undefined:
    with forms.ProgressBar(title='Привязка по соседним . . .', step = 30) as pb:
        for loop in range(0, 3):
            for stack in defined_stacks:
                wall = stack[0]
                room = stack[1]
                step += 1
                pb.update_progress(step, max_value = len(defined_stacks) * 3)
                if room == None:
                    elements = get_neighbourhood(wall)
                    if len(elements) != 0:
                        for i in elements:
                            st = get_stack_by_id(i.Id.IntegerValue)
                            if st != None:
                                if st[1] != None:
                                    room = st[1]
                                    break
step = 0
with db.Transaction(name = "Запись значений"):
    with forms.ProgressBar(title='Запись значений . . .', step = 10) as pb:
        for stack in defined_stacks:
            step += 1
            pb.update_progress(step, max_value = len(defined_stacks))
            wall = stack[0]
            room = stack[1]
            if room != None:
                try:
                    wall.LookupParameter("СИТИ_Тип помещения").Set(room.LookupParameter("ADSK_Тип помещения").AsValueString())
                except AttributeError:
                    wall.LookupParameter("СИТИ_Тип помещения").Set(room.LookupParameter("СИТИ_Тип помещения").AsString())
                if room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString().lower() == "квартира":
                        try:
                            wall.LookupParameter("СИТИ_Номер квартиры").Set(room.LookupParameter("КВ_Номер").AsString())
                        except AttributeError:
                            wall.LookupParameter("СИТИ_Номер квартиры").Set(room.LookupParameter("СИТИ_Номер квартиры").AsString())
                        wall.LookupParameter("СИТИ_Номер помещения").Set(" ")
                else:
                    wall.LookupParameter("СИТИ_Номер квартиры").Set(" ")
                    try:
                        wall.LookupParameter("СИТИ_Номер помещения").Set(room.LookupParameter("ПОМ_Номер_Дополнительный").AsString())
                    except AttributeError:
                        wall.LookupParameter("СИТИ_Номер помещения").Set(room.LookupParameter("СИТИ_Номер помещения").AsString())
            else:
                print("{} ({}) не найдено граничащих помещений!".format(out.linkify(wall.Id), wall.Name))
