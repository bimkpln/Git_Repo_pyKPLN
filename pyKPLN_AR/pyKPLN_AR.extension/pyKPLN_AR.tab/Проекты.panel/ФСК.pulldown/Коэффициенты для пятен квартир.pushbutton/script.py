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
from rpw.ui.forms import Alert, FlexForm, Label, TextBox, Button, Separator
import datetime


# FlexForm для ввода кооэфициентов
coeff_form = FlexForm('Добавить коээфициенты к площадям', [
    Label('Плагин обновит в помещении параметры:'),
    Label('"ПОМ_Площадь_К", "КВ_ПлощадьОбщая_К"'),
    Separator(),
    Label('Укажите коэффициент для помещения "квартира":'),
    TextBox('coeff_1', Text='0.95'),
    Label('Укажите коэффициент для помещения "Лоджия":'),
    TextBox('coeff_2', Text='0.5'),
    Separator(),
    Button('Пуск'),
])

coeff_form.show()

if coeff_form.values:

    coeff_1 = float(coeff_form.values['coeff_1']) if coeff_form.values['coeff_1'] else 0.95
    coeff_2 = float(coeff_form.values['coeff_2']) if coeff_form.values['coeff_2'] else 0.5

    #Собираем номера квартир
    got_errors = False
    flat_numbers = []
    with db.Transaction(name = "Поиск номеров квартир"):
        for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
            try:
                flat_number = room.LookupParameter("КВ_Номер").AsString()
                if flat_number != None and flat_number != "0":
                    if not flat_number in flat_numbers:
                        flat_numbers.append(flat_number)                                    
            except:
                pass

    #Функция добавления кооэффициентов к помещениям и возвращения суммарной площади квартиры
    def GetAndTransformRoomArea(source_rooms):
        sum_area = 0.0
        for room in source_rooms:
            if room.LookupParameter("Имя").AsString() == "квартира":
                area = round(room.Area * 0.09290304 * coeff_1, 1) #поменять коэф
                room.LookupParameter("ПОМ_Площадь_К").Set(area / 0.09290304)
                sum_area += area
            elif room.LookupParameter("Имя").AsString() == "Лоджия":
                area = round(room.Area * 0.09290304 * coeff_2, 1) #поменять коэф
                room.LookupParameter("ПОМ_Площадь_К").Set(area / 0.09290304)
                sum_area += area
        try:
            return (sum_area/ 0.09290304)
        except :
            pass
        return 0.0

    #Итерация по помещениям каждой квартиры
    with db.Transaction(name = "Запись значений площадей"):
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
                    flat_area = GetAndTransformRoomArea(rooms)
                    for room in rooms:
                        room.LookupParameter("КВ_Площадь_Общая_К").Set(flat_area)
            except Exception as e:
                print(str(e))

    ui.forms.Alert("Площади умноженны на коэффициенты.", title = "Готово!")
