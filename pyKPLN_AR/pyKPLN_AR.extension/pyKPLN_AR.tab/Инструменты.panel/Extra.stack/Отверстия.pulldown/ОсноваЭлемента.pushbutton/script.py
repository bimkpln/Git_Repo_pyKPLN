	# -*- coding: utf-8 -*-
"""
OpeningsHost

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Основа проема"
__doc__ = 'Запись имени типа стены, в котором находится отверстие\n' \
		  '      «00_Основа_Элемента» - имя основы элемента\n' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import TextInput, Alert
from rpw.ui.forms import CommandLink, TaskDialog
import datetime

#ПРОВЕРКА НАЛИЧИЯ ПАРАМЕТРОВ
try:
    param_found = False
    app = doc.Application
    category_set_elements = app.Create.NewCategorySet()
    category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Windows))
    category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Doors))
    category_set_elements.Insert(doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_MechanicalEquipment))
    originalFile = app.SharedParametersFilename
    app.SharedParametersFilename = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt"
    SharedParametersFile = app.OpenSharedParameterFile()
    map = doc.ParameterBindings
    it = map.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        d_Definition = it.Key
        d_Name = it.Key.Name
        d_Binding = it.Current
        d_catSet = d_Binding.Categories	
        if d_Name == "00_Основа_Элемента":
            if d_Binding.GetType() == DB.InstanceBinding:
                if str(d_Definition.ParameterType) == "Text":
                    if d_Definition.VariesAcrossGroups:
                        if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Windows)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Doors)) and d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_MechanicalEquipment)):
                            param_found = True
    with db.Transaction(name = "AddSharedParameter"):
        for dg in SharedParametersFile.Groups:
            if dg.Name == "АР_Отверстия":
                if not param_found:
                    externalDefinition = dg.Definitions.get_Item("00_Основа_Элемента")
                    newIB = app.Create.NewInstanceBinding(category_set_elements)
                    doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
                    doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)

    map = doc.ParameterBindings
    it = map.ForwardIterator()
    it.Reset()
    with db.Transaction(name = "SetAllowVaryBetweenGroups"):
        while it.MoveNext():
            d_Definition = it.Key
            d_Name = it.Key.Name
            d_Binding = it.Current
            if d_Name == "00_Основа_Элемента":
                d_Definition.SetAllowVaryBetweenGroups(doc, True)
except: pass
#ОСНОВА
with db.Transaction(name='Запись параметров'):
    for element in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType():
        fam_name = element.Symbol.FamilyName
        if fam_name.startswith("199_Отверстие"):
            try:
                host = element.Host.Name
                parameter = element.LookupParameter("00_Основа_Элемента").Set(host)
            except Exception as e:
                print(str(e))
                pass