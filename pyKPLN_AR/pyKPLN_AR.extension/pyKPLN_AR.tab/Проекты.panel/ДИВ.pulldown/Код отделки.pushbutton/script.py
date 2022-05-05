# -*- coding: utf-8 -*-
"""
DIV:FW:ROOMCODE

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Отделка - Код"
__doc__ = 'Для отделки стен, полов и потолков заполнение параметра "СИТИ_Тип помещения", "СИТИ_Номер помещения", или "СИТИ_Номер квартиры"*\n\n*- из связанных квартир'

"""
KPLN

"""
from rpw import doc, DB, UI, db
from pyrevit import script
from pyrevit import revit as REVIT

categories = [DB.BuiltInCategory.OST_Walls, DB.BuiltInCategory.OST_Floors, DB.BuiltInCategory.OST_Ceilings]
with db.Transaction(name="DIV:FW:ROOMCODE"):
    for cat in categories:
        for element in DB.FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements():
            try:
                type = None
                if element.GetType() == DB.Wall:
                    type = element.WallType
                if element.GetType() == DB.Floor:
                    type = element.FloorType
                if element.GetType() != DB.Wall and element.GetType() != DB.Floor:
                    type = doc.GetElement(element.GetTypeId())
                modelGroup = type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MODEL).AsString()
                if modelGroup and "отделка" in modelGroup.lower():
                    value = element.LookupParameter("О_Id помещения").AsString()
                    try:
                        intValue = int(value)
                    except:
                        out = script.get_output()
                        category_name = DB.Category.GetCategory(doc, cat).Name
                        print("\n\nПРЕДУПРЕЖДЕНИЕ: {0} <{1}> [{2}] - Не назначено помещение в параметре «О_Id помещения»\n\tПодсказка: Необходимо привязать элемент к помещению и перезапустить рассчет или назначить код вручную!".
                              format(out.linkify(element.Id),
                                     category_name,
                                     REVIT.query.get_name(element)))
                        continue
                    room = doc.GetElement(DB.ElementId(intValue))
                    try:
                        element.LookupParameter("СИТИ_Тип помещения").Set(room.LookupParameter("ADSK_Тип помещения").AsValueString())
                    except AttributeError:
                        element.LookupParameter("СИТИ_Тип помещения").Set(room.LookupParameter("СИТИ_Тип помещения").AsString())
                    if room.get_Parameter(DB.BuiltInParameter.ROOM_DEPARTMENT).AsString().lower() == "квартира":
                        try:
                            element.LookupParameter("СИТИ_Номер квартиры").Set(room.LookupParameter("КВ_Номер").AsString())
                        except AttributeError:
                            element.LookupParameter("СИТИ_Номер квартиры").Set(room.LookupParameter("СИТИ_Номер квартиры").AsString())
                        element.LookupParameter("СИТИ_Номер помещения").Set(" ")
                    else:
                        element.LookupParameter("СИТИ_Номер квартиры").Set(" ")
                        try:
                            element.LookupParameter("СИТИ_Номер помещения").Set(room.LookupParameter("ПОМ_Номер_Дополнительный").AsString())
                        except AttributeError:
                            element.LookupParameter("СИТИ_Номер помещения").Set(room.LookupParameter("СИТИ_Номер помещения").AsString())
            except Exception as e:
                category_name = DB.Category.GetCategory(doc, cat).Name
                if "'NoneType' object has no attribute 'LookupParameter'" in e.ToString():
                    print("\nERROR: <{0}> [{1}]\n«{2}»".format(category_name, element.Id.ToString(), "Привязанного к элементу помещения с таким id не существует"))
                elif "AsValueString" not in e.ToString():
                    print("\nERROR: <{0}> [{1}]\n«{2}»".format(category_name, element.Id.ToString(), str(e.ToString())))