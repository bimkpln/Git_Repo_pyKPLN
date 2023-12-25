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
               title="Подгрузить и заполнить параметры для реестра ПСО",
               title_prefix=False,
               content = 'Время работы плагина 5 минут',
               commands=commands,
               show_close=True)
dialog_out = dialog.show()

#подгрузка параметров 
def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False

fop_path = "Y:\\Жилые здания\\ФСК_Измайловский\\6.BIM\\5.Файл_общих_параметров_ФОП\\ФСК_ФОП_v2.txt"
parameters_to_load =[["ФСК_Реестр_КВ_Площадь_Балкон", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Ванная комната", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Гардеробная", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Гардеробная №2", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Комната №1", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Комната №2", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Комната №3", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Комната №4", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Коридор", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Кухня", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Кухня-ниша", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Кухня-гостинная", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Лоджия", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Постирочная", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Санузел", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Совмещенный санузел №1", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Совмещенный санузел №2", "Area", True],
    ["ФСК_Реестр_КВ_Площадь_Прихожая", "Area", True],
    ["ФСК_Реестр_КВ_Высота потолка", "Text", True],
    ["ФСК_Реестр_КВ_Кол-во комнат", "Text", True],
    ["ФСК_Реестр_ПОН_Площадь_ПУИ", "Area", True],
    ["ФСК_Реестр_ПОН_Площадь_С/у МГН", "Area", True],
    ["ФСК_Реестр_ПОН_Площадь_Помещение ПОН", "Area", True],
    ["ФСК_ПОН (БКТ)_Площадь_Общая", "Area", True],
    ["ФСК_Группировка_ПОН (БКТ)", "Text", False]]

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
map = doc.ParameterBindings
it = map.ForwardIterator()
it.Reset()

if dialog_out == 'go':
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
        with db.Transaction(name = "Настройка параметров"):
            while it.MoveNext():
                for i in range(0, len(parameters_to_load)):
                    if not in_list(parameters_to_load[i][0], params_found):
                        d_Definition = it.Key
                        d_Name = it.Key.Name
                        d_Binding = it.Current
                        if d_Name == parameters_to_load[i][0]:
                            d_Definition.SetAllowVaryBetweenGroups(doc, parameters_to_load[i][2])

    #Собираем номера ПОН и квартир
    got_errors = False
    flat_numbers = []
    pon_numbers = []
    with db.Transaction(name = "Поиск номеров ПОН и квартир"):
        for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
            try:
                pon_number = room.LookupParameter("ФСК_Группировка_ПОН (БКТ)").AsString()
                flat_number = room.LookupParameter("КВ_Номер").AsString()

                if pon_number != None and pon_numbers != "0":
                    if not pon_number in pon_numbers:
                        pon_numbers.append(pon_number)

                if flat_number != None and flat_number != "0":
                    if not flat_number in flat_numbers:
                        flat_numbers.append(flat_number)                                    
            except:
                pass

    def GetRoomArea(source_rooms, param_name, name):
        rooms = []
        for room in source_rooms:
            if room.LookupParameter(param_name).AsString() == name:
                rooms.append(room)
        try:
            return rooms[0].LookupParameter("ПОМ_Площадь_К").AsDouble()
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
        #Итерация по помещениям каждого ПОН(БКТ)
        for pn in pon_numbers:
            try:
                pon_rooms = []
                for pon_room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
                    a = pon_room.Area
                    try:
                        pon_number = pon_room.LookupParameter("ФСК_Группировка_ПОН (БКТ)").AsString()
                        if pon_number == pn and a > 0 and a != None:
                            pon_rooms.append(pon_room)
                    except:
                        continue
                if len(pon_rooms) != 0:
                    print("Расчитываю помещения " + pn)
                    for room in pon_rooms:
                        sum_area = 0.0
                        area = GetRoomArea(pon_rooms, "Имя", "ПУИ")
                        sum_area += area
                        ApplyValue(pon_rooms, area, "ФСК_Реестр_ПОН_Площадь_ПУИ")
                        area = GetRoomArea(pon_rooms, "Имя", "С/у МГН")
                        sum_area += area
                        ApplyValue(pon_rooms, area, "ФСК_Реестр_ПОН_Площадь_С/у МГН")
                        area = GetRoomArea(pon_rooms, "Имя", "Помещение ПОН")
                        sum_area += area
                        ApplyValue(pon_rooms, area, "ФСК_Реестр_ПОН_Площадь_Помещение ПОН")

                        ApplyValue(pon_rooms, sum_area, "ФСК_ПОН (БКТ)_Площадь_Общая")
            except Exception as e:
                print(str(e))

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
                    print("Расчитываю помещения квартиры №" + fn)
                    for room in rooms:
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Санузел")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Санузел")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Балкон")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Балкон")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Ванная комната")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Ванная комната")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Гардеробная №1")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Гардеробная")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Гардеробная №2")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Гардеробная №2")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Лоджия")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Лоджия")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Кухня")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Кухня")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Кухня-ниша")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Кухня-ниша")                        
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Коридор")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Коридор")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Прихожая")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Прихожая")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Совмещенный санузел №1")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Совмещенный санузел №1")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Совмещенный санузел №2")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Совмещенный санузел №2")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Постирочная")
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Постирочная")

                        living_rooms_amount = 0
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Кухня-гостиная")
                        if (area != 0.0):
                            living_rooms_amount+=1
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Кухня-гостинная")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Комната №1")
                        if (area != 0.0):
                            living_rooms_amount+=1
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Комната №1")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Комната №2")
                        if (area != 0.0):
                            living_rooms_amount+=1
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Комната №2")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Комната №3")
                        if (area != 0.0):
                            living_rooms_amount+=1
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Комната №3")
                        area = GetRoomArea(rooms, "ФСК_Имя помещения", "Комната №4")
                        if (area != 0.0):
                            living_rooms_amount+=1
                        ApplyValue(rooms, area, "ФСК_Реестр_КВ_Площадь_Комната №4")

                        if (living_rooms_amount > 0):
                            ApplyValue(rooms, str(living_rooms_amount), "ФСК_Реестр_КВ_Кол-во комнат")
            except Exception as e:
                print(str(e))

    ui.forms.Alert("Параметры для реестра ПСО заполнены.", title = "Готово!")