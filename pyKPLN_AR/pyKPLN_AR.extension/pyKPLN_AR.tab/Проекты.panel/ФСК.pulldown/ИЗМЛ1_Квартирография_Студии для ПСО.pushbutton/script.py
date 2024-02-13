# -*- coding: utf-8 -*-

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

commands= [CommandLink('Запустить', return_value='go'),
            CommandLink('Отмена', return_value='stop')]
dialog = TaskDialog('Внимание!',
               title="Запустить квартирографию для студий для ПСО",
               title_prefix=False,
               content = 'Плагин обновит параметры:\n -КВ_Площадь_Жилая (для всех помещений студий)\n -КВ_Площадь_Нежилая (для всех помещений студий)\n -ПОМ_Площадь_К (для Кухонь-нишь и Комнат №1 в студиях)\n -ПОМ_Площадь (для Кухонь-нишь и Комнат №1 в студиях)',
               commands=commands,
               show_close=True)
dialog_out = dialog.show()

#подгрузка параметров 
def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False


if dialog_out == 'go':

    #Собираем номера ПОН и квартир
    got_errors = False
    flat_numbers = []
    with db.Transaction(name = "Поиск номеров студий"):
        for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
            try:
                flat_code = room.LookupParameter("КВ_Код").AsString()
                if flat_code == "C" or flat_code == "С":
                    flat_number = room.LookupParameter("КВ_Номер").AsString()
                    if flat_number != None and flat_number != "0":
                        if not flat_number in flat_numbers:
                            flat_numbers.append(flat_number)                                    
            except:
                pass

    def GetAndTransformRoomArea(source_rooms, param_name, name):
        rooms = []
        for room in source_rooms:
            if room.LookupParameter(param_name).AsString() == name:
                rooms.append(room)
                area = round(room.Area * 0.09290304, 1)
                room.LookupParameter("ПОМ_Площадь").Set(area / 0.09290304)
                room.LookupParameter("ПОМ_Площадь_К").Set(area / 0.09290304)
        try:
            return (area/ 0.09290304)
        except :
            pass
        return 0.0

    def ApplyValue(rooms, value, name):
        for room in rooms:
            try:
                room.LookupParameter(name).Set(value)
            except Exception as e:
                print(str(e))

    with db.Transaction(name = "Запись значений"):
        #Итерация по помещениям каждой квартиры
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
                    print("Расчитываю помещения студии №" + fn)
                    kitchen_area = GetAndTransformRoomArea(rooms, "ФСК_Имя помещения", "Кухня-ниша")
                    living_room_area = GetAndTransformRoomArea(rooms, "ФСК_Имя помещения", "Комната №1")
                    flat_area = rooms[0].LookupParameter("КВ_Площадь_Отапливаемые").AsDouble()
                    for room in rooms:
                        ApplyValue(rooms, living_room_area, "КВ_Площадь_Жилая")   
                        ApplyValue(rooms, (flat_area-living_room_area), "КВ_Площадь_Нежилая")     
            except Exception as e:
                print(str(e))

    ui.forms.Alert("Квартирография для студий обновлена.", title = "Готово!")