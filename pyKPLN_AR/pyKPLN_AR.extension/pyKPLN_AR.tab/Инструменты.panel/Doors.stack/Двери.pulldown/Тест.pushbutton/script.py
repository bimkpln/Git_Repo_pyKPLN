# -*- coding: utf-8 -*-
"""
KPLN:DIV:AMMOUNT:COUNTER

"""
__author__ = ''
__title__ = "Тест"
__doc__ = 'Подсчет кол-ва элементов на этаж с записью значений в выбранные параметры' \

"""
Архитекурное бюро KPLN

"""
import math
import re
import os
from Autodesk.Revit.DB import *
from pyrevit.framework import clr
from System.Collections.Generic import *
from System.Windows import Application, Window
from System.Collections.ObjectModel import ObservableCollection
from rpw import doc, uidoc, DB, UI, db, ui, revit
from rpw.ui.forms import CommandLink, TaskDialog, Alert
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
import datetime
import webbrowser
import wpf

#################################
#################################

# doors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements()

# levels_param_list = ['КП_А_03_Этаж', 'КП_А_04_Этаж', 'КП_А_05_Этаж']


# with db.Transaction(name = "Запись результатов"):
#     for door in doors:
#         for i in range(0, len(levels_param_list)):
#             door.LookupParameter(levels_param_list[i]).Set(door.LookupParameter('КП_А_02_Этаж').AsDouble())


element = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().FirstElement()
params = []

param_list = element.Parameters
for param in param_list:
    # if param.StorageType == DB.StorageType.Double:
    if param.Definition.Name.startswith("КП_А_"):
        params.append(param.Definition.Name)

parameters = forms.SelectFromList.show(sorted(params), title="Выберите параметры типового этажа", multiselect=True,)

for i in parameters:
    print(i)








