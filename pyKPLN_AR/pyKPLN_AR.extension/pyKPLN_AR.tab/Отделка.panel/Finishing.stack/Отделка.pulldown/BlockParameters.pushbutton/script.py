# -*- coding: utf-8 -*-
"""
FINISHING_BlockParameters

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Заблокировать параметры"
__doc__ = 'Плагин записывает параметры привязки элементов\n' \
          'отделки (стены, перекрытия, потолки) к помещениям\n' \
          'в нередактируемые параметры с префиксом "SYS_"\n' \

"""
Архитекурное бюро KPLN

"""
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
from Autodesk.Revit.DB import *
from libKPLN import kpln_logs

#Параметры для логирования в Extensible Storage. Не менять по ходу изменений кода
extStorage_guid = "720080C5-DA99-40D7-9445-E53F288AA160"
extStorage_name = "kpln_ar_finishing"

if __shiftclick__:
   try:
       obj = kpln_logs.create_obj(extStorage_guid, extStorage_name)
       kpln_logs.read_log(obj)
   except:
       print("Записи отсутствуют")
   script.exit()

out = script.get_output()
out.set_title("Запись значений")

categories = [BuiltInCategory.OST_Walls, BuiltInCategory.OST_Ceilings, BuiltInCategory.OST_Floors]

def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False

commands= [CommandLink('Да', return_value=True),
            CommandLink('Нет', return_value=False)]
dialog = TaskDialog('Перезаписать параметры привязки отделки?',
               title="Блокировка параметров",
               title_prefix=False,
               content = "ВНИМАНИЕ! Привязка элементов отделки к помещениям будет перезаписана! Продолжить?",
               commands=commands,
               show_close=True)
dialog_out = dialog.show()
if dialog_out:
    fop_path = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt"
    parameters_to_load =[["SYS_О_Назначение помещения", "Text", True],
        ["SYS_О_Номер секции", "Text", True],
        ["SYS_О_Имя помещения", "Text", True],
        ["SYS_О_Группа", "Text", True],
        ["SYS_О_Id помещения", "Text", True],
        ["SYS_О_Номер помещения", "Text", True],
        ["SYS_О_Тип", "Text", True]]
    parameters_name = ["О_Назначение помещения",
        "О_Номер секции",
        "О_Имя помещения",
        "О_Группа",
        "О_Id помещения",
        "О_Номер помещения",
        "О_Тип"]
    params_found = []
    group = "АР_Отделка"
    common_parameters_file = fop_path
    app = doc.Application
    originalFile = app.SharedParametersFilename
    category_set_elements = app.Create.NewCategorySet()
    category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Walls))
    category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Ceilings))
    category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Floors))
    # category_set_rooms.Insert(insert_cat_rooms)
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
                            newIB = app.Create.NewInstanceBinding(category_set_elements)
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
    with db.Transaction(name = "КПЛН_Запись результатов"):
        for cat in categories:
            elements = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for i in range(0, len(parameters_name)):
                for element in elements:
                    try:
                        param = element.LookupParameter(parameters_name[i]).AsString()
                        if param:
                            element.LookupParameter(parameters_to_load[i][0]).Set(param)
                    except:
                        pass
        # Запись логов
        try:
            obj = kpln_logs.create_obj(extStorage_guid, extStorage_name)
            kpln_logs.write_log(obj, "Значения были зафиксированы")
        except:
            print("Ошибка записи. Обратитесь в BIM - отдел!")
    ui.forms.Alert("Значения параметров перезаписаны и зафиксированы.", title = "Готово!")