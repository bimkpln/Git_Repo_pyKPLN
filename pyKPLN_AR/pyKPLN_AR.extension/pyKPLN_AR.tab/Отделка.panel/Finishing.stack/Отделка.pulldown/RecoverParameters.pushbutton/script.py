# -*- coding: utf-8 -*-
"""
FINISHING_BlockParameters

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Восстановить параметры"
__doc__ = 'Плагин восстанавливает знаения параметров привязки элементов\n' \
          'отделки (стены, перекрытия, потолки) к помещениям\n' \


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

categories = [BuiltInCategory.OST_Walls, BuiltInCategory.OST_Ceilings, BuiltInCategory.OST_Floors]

def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False

commands= [CommandLink('Да', return_value=True),
            CommandLink('Нет', return_value=False)]
dialog = TaskDialog('Восстановить параметры привязки отделки?',
               title="Восстановление параметров",
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

    with db.Transaction(name = "КПЛН_Запись результатов"):
        for cat in categories:
            elements = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for i in range(0, len(parameters_name)):
                for element in elements:
                    try:
                        param = element.LookupParameter(parameters_to_load[i][0]).AsString()
                        if param:
                            element.LookupParameter(parameters_name[i]).Set(param)
                    except:
                        pass
    ui.forms.Alert("Значения параметров восстановлены и перезаписаны.", title = "Готово!")