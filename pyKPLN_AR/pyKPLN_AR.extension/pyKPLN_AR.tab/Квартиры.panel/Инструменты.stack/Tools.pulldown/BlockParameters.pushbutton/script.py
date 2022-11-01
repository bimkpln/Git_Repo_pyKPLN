# -*- coding: utf-8 -*-
"""
ROOMS_BlockParameters

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Заблокировать площади стадии П"
__doc__ = 'Запись утвержденных значений площадей в заблокированные параметры' \

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

out = script.get_output()
out.set_title("Запись значений")


def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False

rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

def is_already_blocked():
    for room in rooms:
        try:
            if room.LookupParameter('П_Площадь').IsReadOnly == False:
                return True
            else:
                return False
        except:
            return False


commands= [CommandLink('Да', return_value=True),
            CommandLink('Нет', return_value=False)]
dialog = TaskDialog('Зафиксировать значения параметров?',
               title="Блокировка параметров",
               title_prefix=False,
               content = "ВНИМАНИЕ! Значения параметров площадей помещений будут перезаписаны и заблокированы! Продолжить?",
               commands=commands,
               show_close=True)
dialog_out = dialog.show()
if dialog_out:
    if not is_already_blocked():
        fop_path = "X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\lib\\ФОП_Scripts.txt"
        parameters_to_load =[["П_КВ_Площадь_Жилая", "Area", True],
            ["П_КВ_Площадь_Летние", "Area", True],
            ["П_КВ_Площадь_Нежилая", "Area", True],
            ["П_КВ_Площадь_Общая", "Area", True],
            ["П_КВ_Площадь_Общая_К", "Area", True],
            ["П_КВ_Площадь_Отапливаемые", "Area", True],
            ["П_Имя", "Text", True],
            ["П_Номер", "Text", True],
            ["П_Площадь", "Area", True],
            ["П_ПОМ_Площадь", "Area", True],
            ["П_ПОМ_Площадь_К", "Area", True]]
        params_found = []
        group = "АРХИТЕКТУРА - Квартирография"
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
        with db.Transaction(name = "Запись результатов"):
            for room in rooms:
                    room.LookupParameter('П_КВ_Площадь_Жилая').Set(room.LookupParameter('КВ_Площадь_Жилая').AsDouble())
                    room.LookupParameter('П_КВ_Площадь_Летние').Set(room.LookupParameter('КВ_Площадь_Летние').AsDouble())
                    room.LookupParameter('П_КВ_Площадь_Нежилая').Set(room.LookupParameter('КВ_Площадь_Нежилая').AsDouble())
                    room.LookupParameter('П_КВ_Площадь_Общая').Set(room.LookupParameter('КВ_Площадь_Общая').AsDouble())
                    room.LookupParameter('П_КВ_Площадь_Общая_К').Set(room.LookupParameter('КВ_Площадь_Общая_К').AsDouble())
                    room.LookupParameter('П_КВ_Площадь_Отапливаемые').Set(room.LookupParameter('КВ_Площадь_Отапливаемые').AsDouble())
                    room.LookupParameter('П_Имя').Set(room.LookupParameter('Имя').AsString())
                    room.LookupParameter('П_Номер').Set(room.LookupParameter('Номер').AsString())
                    room.LookupParameter('П_Площадь').Set(room.LookupParameter('Площадь').AsDouble())
                    room.LookupParameter('П_ПОМ_Площадь').Set(room.LookupParameter('ПОМ_Площадь').AsDouble())
                    room.LookupParameter('П_ПОМ_Площадь_К').Set(room.LookupParameter('ПОМ_Площадь_К').AsDouble())
        ui.forms.Alert("Значения параметров зафиксированы.", title = "Готово!")
    else:
        ui.forms.Alert("Значения уже были заблокированы ранее!", title = "Операция не выполнена")