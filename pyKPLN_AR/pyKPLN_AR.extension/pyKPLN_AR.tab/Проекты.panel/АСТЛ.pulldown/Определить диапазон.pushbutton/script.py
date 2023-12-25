# -*- coding: utf-8 -*-

import clr
clr.AddReference('RevitAPI')
from rpw import doc, DB, db, ui
from rpw.ui.forms import CommandLink, TaskDialog

rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

commands= [CommandLink('Запустить', return_value='go'),
            CommandLink('Отмена', return_value='stop')]
dialog = TaskDialog('Внимание!',
               title="Определение диапазона",
               title_prefix=False,
               content = 'Для успешного определения диапазона, должен быть заполнен параметр:\n\n- КВ_Код (1Е, 2К и т.д.)',
               commands=commands,
               show_close=True)
dialog_out = dialog.show()

if dialog_out == 'go':
    with db.Transaction(name = "ta"):
        for room in rooms:
            name = room.LookupParameter("КВ_Код").AsString()
            area = room.LookupParameter("КВ_Площадь_Общая_К").AsDouble() * 0.09290304

            # Обнуляем значения
            room.LookupParameter("КВ_Диапазон").Set("ВНЕ ДИАПАЗОНА")

            if name == "C" or name == "С":
                print("Студия"+ str(area))
                if area >= 20 and area <= 25.99:
                    room.LookupParameter("КВ_Диапазон").Set("S-S")
                if area >= 26 and area <= 29.99:
                    room.LookupParameter("КВ_Диапазон").Set("S-M")
                if area >= 30 and area <= 34:
                    room.LookupParameter("КВ_Диапазон").Set("S-L")
            if name == "2Е":
                if area >= 36 and area <= 40.99:
                    room.LookupParameter("КВ_Диапазон").Set("2Е-S")
                if area >= 41 and area <= 44.99:
                    room.LookupParameter("КВ_Диапазон").Set("2Е-M")
                if area >= 45 and area <= 49:
                    room.LookupParameter("КВ_Диапазон").Set("2Е-L")
            if name == "3Е":
                if area >= 52 and area <= 56.99:
                    room.LookupParameter("КВ_Диапазон").Set("3Е-S")
                if area >= 57 and area <= 61.99:
                    room.LookupParameter("КВ_Диапазон").Set("3Е-M")
                if area >= 62 and area <= 66.99:
                    room.LookupParameter("КВ_Диапазон").Set("3Е-L")
                if area >= 67 and area <= 73.99:
                    room.LookupParameter("КВ_Диапазон").Set("3Е-XL")                    
            if name == "4Е":
                if area >= 77 and area <= 82.99:
                    room.LookupParameter("КВ_Диапазон").Set("4Е-S")
                if area >= 83 and area <= 88.99:
                    room.LookupParameter("КВ_Диапазон").Set("4Е-M")
                if area >= 101 and area <= 110:
                    room.LookupParameter("КВ_Диапазон").Set("4Е-XXL")
            if name == "4Е":
                if area >= 110 and area <= 130.99:
                    room.LookupParameter("КВ_Диапазон").Set("4К-M")
                if area >= 131 and area <= 150:
                    room.LookupParameter("КВ_Диапазон").Set("5Е-L")

    ui.forms.Alert("Диапазоны определены.", title = "Готово!")