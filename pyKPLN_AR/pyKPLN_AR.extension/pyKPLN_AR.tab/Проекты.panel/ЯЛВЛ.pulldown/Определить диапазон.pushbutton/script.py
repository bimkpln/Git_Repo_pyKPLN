# -*- coding: utf-8 -*-

import clr
clr.AddReference('RevitAPI')
from rpw import doc, DB, db, ui
from rpw.ui.forms import CommandLink, TaskDialog

def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False

rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

commands= [CommandLink('Запустить', return_value='go'),
            CommandLink('Отмена', return_value='stop')]
dialog = TaskDialog('Внимание!',
               title="Определение диапазона",
               title_prefix=False,
               content = 'Значение будет вписано в системный параметр "КВ_Диапазон"',
               commands=commands,
               show_close=True)
dialog_out = dialog.show()

if dialog_out == 'go':
    fop_path = "X:\\BIM\\4_ФОП\\КП_Файл общих парамеров.txt"
    parameters_to_load =[["КВ_Диапазон", "Text", True], ["КВ_Число", "Currency", True]]
    params_found = []
    group = "11 Временные АРХИТЕКТУРА"
    common_parameters_file = fop_path
    app = doc.Application
    originalFile = app.SharedParametersFilename
    category_set_rooms = app.Create.NewCategorySet()
    insert_cat_rooms = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Rooms)
    category_set_rooms.Insert(insert_cat_rooms)
    app.SharedParametersFilename = common_parameters_file
    SharedParametersFile = app.OpenSharedParameterFile()
    #CHECK CURRENT PARAMETERS
    map = doc.ParameterBindings
    it = map.ForwardIterator()
    it.Reset()
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
        with db.Transaction(name="Настройка параметров"):
            while it.MoveNext():
                for i in range(0, len(parameters_to_load)):
                    if not in_list(parameters_to_load[i][0], params_found):
                        d_Definition = it.Key
                        d_Name = it.Key.Name
                        d_Binding = it.Current
                        if d_Name == parameters_to_load[i][0]:
                            d_Definition.SetAllowVaryBetweenGroups(doc, parameters_to_load[i][2])
    with db.Transaction(name = "Определить диапазоны"):
        if len(rooms) > 0:
            for room in rooms:
                name = room.LookupParameter("КВ_Наименование").AsString()
                area = room.Area * 0.09290304
                # Обнуляем значения
                room.LookupParameter("КВ_Диапазон").Set("ВНЕ ДИАПАЗОНА")
                room.LookupParameter("КВ_Число").Set(1)
                # Определяем диапазон
                if name == "Студия":
                    if area >= 23 and area < 26:
                        room.LookupParameter("КВ_Диапазон").Set("23-25")
                    if area >= 26 and area < 32:
                        room.LookupParameter("КВ_Диапазон").Set("26-31")
                    if area >= 32 and area < 60:
                        room.LookupParameter("КВ_Диапазон").Set("32-59")
                if name == "Однокомнатная квартира":
                    if area >= 36 and area < 50:
                        room.LookupParameter("КВ_Диапазон").Set("36-49")
                    if area >= 50 and area < 61:
                        room.LookupParameter("КВ_Диапазон").Set("50-60")
                    if area >= 61 and area < 71:
                        room.LookupParameter("КВ_Диапазон").Set("61-70")
                if name == "Двухкомнатная квартира":
                    if area >= 42 and area < 81:
                        room.LookupParameter("КВ_Диапазон").Set("42-80")
                    if area >= 81 and area < 121:
                        room.LookupParameter("КВ_Диапазон").Set("81-120")
                    if area >= 121 and area < 164:
                        room.LookupParameter("КВ_Диапазон").Set("121-163")
                if name == "Трехкомнатная квартира":
                    if area >= 76 and area < 101:
                        room.LookupParameter("КВ_Диапазон").Set("76-100")
                    if area >= 101 and area < 161:
                        room.LookupParameter("КВ_Диапазон").Set("101-160")
                    if area >= 161 and area < 205:
                        room.LookupParameter("КВ_Диапазон").Set("161-204")
            ui.forms.Alert("Диапазоны определены.", title = "Готово!")
        else:
            ui.forms.Alert("Помещения отсутствуют.", title = "Ошибка!")
