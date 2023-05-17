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
               content = 'Для успешного определения диапазона, должны быть заполнены следущие параметры:\n\n- Комментарии (38 или 45)\n- КВ_Код (1Е, 2К и т.д.)',
               commands=commands,
               show_close=True)
dialog_out = dialog.show()

if dialog_out == 'go':
    with db.Transaction(name = "ta"):
        for room in rooms:
            name = room.LookupParameter("КВ_Код").AsString()
            area = room.LookupParameter("КВ_Площадь_Общая_К").AsDouble() * 0.09290304
            # area = (room.Area * 0.09290304 - 3.5) * 0.91 + 3.5 * 0.5
            s_type = room.LookupParameter("Комментарии").AsString()
            # Обнуляем значения
            room.LookupParameter("КВ_Диапазон").Set("ВНЕ ДИАПАЗОНА")
            # Средняя площадь 38м2
            if s_type == "38":
                if name == "C" or name == "С":
                    if area > 17 and area < 21:
                        room.LookupParameter("КВ_Диапазон").Set("18-20")
                    if area >= 20 and area < 27:
                        room.LookupParameter("КВ_Диапазон").Set("20-26")
                    if area >= 27 and area < 30:
                        room.LookupParameter("КВ_Диапазон").Set("27-29")
                if name == "1К":
                    if area > 29 and area < 34:
                        room.LookupParameter("КВ_Диапазон").Set("30-33")
                    if area >= 34 and area < 38:
                        room.LookupParameter("КВ_Диапазон").Set("34-37")
                    if area >= 38 and area < 42:
                        room.LookupParameter("КВ_Диапазон").Set("38-41")
                if name == "2Е":
                    if area > 41 and area < 46:
                        room.LookupParameter("КВ_Диапазон").Set("42-45")
                    if area >= 46 and area < 50:
                        room.LookupParameter("КВ_Диапазон").Set("46-49")
                    if area >= 50 and area < 55:
                        room.LookupParameter("КВ_Диапазон").Set("50-54")
                if name == "2К":
                    if area > 44 and area < 51:
                        room.LookupParameter("КВ_Диапазон").Set("45-50")
                    if area >= 50 and area < 55:
                        room.LookupParameter("КВ_Диапазон").Set("50-54")
                    if area >= 55 and area < 60:
                        room.LookupParameter("КВ_Диапазон").Set("55-59")
                if name == "3Е":
                    if area > 59 and area < 65:
                        room.LookupParameter("КВ_Диапазон").Set("60-64")
                    if area >= 65 and area < 70:
                        room.LookupParameter("КВ_Диапазон").Set("65-69")
                if name == "3К":
                    if area > 69 and area < 75:
                        room.LookupParameter("КВ_Диапазон").Set("70-74")
                    if area >= 75 and area < 80:
                        room.LookupParameter("КВ_Диапазон").Set("75-79")
                if name == "4Е":
                    if area > 77 and area < 86:
                        room.LookupParameter("КВ_Диапазон").Set("78-85")
                    if area >= 86 and area < 91:
                        room.LookupParameter("КВ_Диапазон").Set("86-90")

            # Средняя площадь 45м2
            if s_type == "45":
                if name == "C" or name == "С":
                    if area > 19 and area < 24:
                        room.LookupParameter("КВ_Диапазон").Set("20-23")
                    if area >= 24 and area < 27:
                        room.LookupParameter("КВ_Диапазон").Set("24-26")
                    if area >= 27 and area < 30:
                        room.LookupParameter("КВ_Диапазон").Set("27-29")
                if name == "1К":
                    if area > 29 and area < 34:
                        room.LookupParameter("КВ_Диапазон").Set("30-33")
                    if area >= 34 and area < 38:
                        room.LookupParameter("КВ_Диапазон").Set("34-37")
                    if area >= 38 and area < 42:
                        room.LookupParameter("КВ_Диапазон").Set("38-41")
                if name == "2Е":
                    if area > 41 and area < 46:
                        room.LookupParameter("КВ_Диапазон").Set("42-45")
                    if area >= 46 and area < 50:
                        room.LookupParameter("КВ_Диапазон").Set("46-49")
                if name == "2К":
                    if area > 49 and area < 55:
                        room.LookupParameter("КВ_Диапазон").Set("50-54")
                    if area >= 55 and area < 60:
                        room.LookupParameter("КВ_Диапазон").Set("55-59")
                if name == "3Е":
                    if area > 59 and area < 66:
                        room.LookupParameter("КВ_Диапазон").Set("60-65")
                    if area >= 65 and area < 70:
                        room.LookupParameter("КВ_Диапазон").Set("65-69")
                if name == "3К":
                    if area > 69 and area < 75:
                        room.LookupParameter("КВ_Диапазон").Set("70-74")
                    if area >= 75 and area < 80:
                        room.LookupParameter("КВ_Диапазон").Set("75-79")
                    if area >= 80 and area < 85:
                        room.LookupParameter("КВ_Диапазон").Set("80-84")
                if name == "4Е":
                    if area > 84 and area < 90:
                        room.LookupParameter("КВ_Диапазон").Set("85-89")
                    if area >= 90 and area < 95:
                        room.LookupParameter("КВ_Диапазон").Set("90-94")
                if name == "4К":
                    if area > 79 and area < 90:
                        room.LookupParameter("КВ_Диапазон").Set("80-89")

    ui.forms.Alert("Диапазоны определены.", title = "Готово!")