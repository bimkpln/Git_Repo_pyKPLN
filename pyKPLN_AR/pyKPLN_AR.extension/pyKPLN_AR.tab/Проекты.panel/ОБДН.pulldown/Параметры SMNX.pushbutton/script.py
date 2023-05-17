# -*- coding: utf-8 -*-

import clr
clr.AddReference('RevitAPI')
from rpw import doc, DB, db

rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

with db.Transaction(name = "ta"):
    for room in rooms:
        r_type_com = room.LookupParameter("SMNX_Тип_ком.помещений").AsString()
        r_level = room.LookupParameter("SMNX_Этаж").AsString()

        s_room_koef = room.LookupParameter("ПОМ_Площадь_К").AsValueString()
        room.LookupParameter("SMNX_Площадь с коэффициентом").Set(s_room_koef)

        r_number = room.LookupParameter("ПОМ_Номер помещения").AsString()
        room.LookupParameter("SMNX_Номер_помещения").Set(r_number)

        r_name = room.LookupParameter("Имя").AsString()
        room.LookupParameter("SMNX_Имя_помещения").Set(r_name)

        r_func = room.LookupParameter("Назначение").AsString()
        room.LookupParameter("SMNX_Назначение_помещения").Set(r_func)

        if room.LookupParameter("Назначение").AsString() == "Квартира":
            s_flat_fact = room.LookupParameter("КВ_Площадь_Общая").AsValueString()
            room.LookupParameter("SMNX_Площадь квартиры_Общая").Set(s_flat_fact)
            s_flat_living = room.LookupParameter("КВ_Площадь_Жилая").AsDouble()
            room.LookupParameter("SMNX_Площадь квартиры_Жилая").Set(s_flat_living)
            s_flat_unliving = room.LookupParameter("КВ_Площадь_Нежилая").AsDouble()
            room.LookupParameter("SMNX_Площадь квартиры_Нежил").Set(s_flat_unliving)
            s_flat_heat = room.LookupParameter("КВ_Площадь_Отапливаемые").AsDouble()
            room.LookupParameter("SMNX_Площадь квартиры_Отапливаемая").Set(s_flat_heat)
            flat_type = room.LookupParameter("Вместимость").AsString()
            room.LookupParameter("SMNX_Типоразмер").Set(flat_type)
            r_flat_number = room.LookupParameter("ПОМ_Номер квартиры").AsString()
            # flat_number = room.LookupParameter("КВ_Номер").AsString()
            room.LookupParameter("SMNX_Номер_квартиры").Set(r_flat_number)
            flat_code = room.LookupParameter("КВ_Код").AsString()
            room.LookupParameter("SMNX_Наименование_типа_квартиры").Set(flat_code)
            s_flat_koef = room.LookupParameter("КВ_Площадь_Общая_К").AsValueString()
            room.LookupParameter("SMNX_Площадь квартиры_с коэф.").Set(s_flat_koef)
            if r_name == "Балкон" or r_name == "Терраса":
                room.LookupParameter("SMNX_Площадь_балкона").Set(room.LookupParameter("ПОМ_Площадь").AsDouble())

        if r_type_com != None and r_level != None and r_flat_number != None:
            if len(r_type_com + "." + r_level + "." + r_flat_number) > 4:
                room.LookupParameter("SMNX_Номер_квартиры_стандарт_SMNX").Set(r_type_com + "." + r_level + "." + r_flat_number)

        if r_level == '1':
            r_level_name = 'Первый'
        elif r_level == '2':
            r_level_name = 'Второй'
        elif r_level == '3':
            r_level_name = 'Третий'
        elif r_level == '4':
            r_level_name = 'Четвертый'
        elif r_level == '5':
            r_level_name = 'Пятый'
        elif r_level == '6':
            r_level_name = 'Шестой'
        elif r_level == 'м1':
            r_level_name = 'Минус первый'
        elif r_level == 'м2':
            r_level_name = 'Минус второй'
        elif r_level == 'м3':
            r_level_name = 'Минус третий'
        else:
            r_level_name = ''
        room.LookupParameter("SMNX_Этаж_Имя").Set(r_level_name)

        r_wet = False
        if r_name.lower() in 'санузел с постирочной постирочная пуи кухня-ниша душевая техническое оборудование спа санузел персонала':
            r_wet = True
        room.LookupParameter("SMNX_Мокрая зона").Set(r_wet)