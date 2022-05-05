# -*- coding: utf-8 -*-
"""
KPLN:DIV:DIV_DESCRIPTION SETTER

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Описание по WBS"
__doc__ = 'Заполнение описания по коду классификатора (Y:\\Жилые здания\\ДИВНОЕ СитиXXI\\7. Стадия РД\\Исходные данные\\Классификатор City-ХХI век\\CITY_Классификатор_csv.txt)' \

"""
KPLN

"""
import math
import webbrowser
import re
import os
import io
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import framework
from pyrevit import script
from pyrevit import forms
from pyrevit import revit, DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
import System
from System.Windows.Forms import *
from System.Drawing import *
import re
from itertools import chain
import datetime
from rpw.ui.forms import TextInput, Alert, select_folder
max = 0
for category in doc.Settings.Categories:
    if not category.IsTagCategory:
        max += DB.FilteredElementCollector(doc).OfCategoryId(category.Id).ToElements().Count
current = 0
items = []
categories = []
parameter_to_set = str(TextInput('Наименование параметра для записи описания', default="", exit_on_close = False))

def GetCategory(value, list):
    for i in list:
        if value == i[0]:
            return i[1]
    return None

def GetName(element):
    value = "<{}> ".format(element.Category.Name)
    try:
        if type(element) == DB.FamilyInstance:
            value += "ЭКЗ. - "
            value += "{}({})".format(element.Symbol.FamilyName, revit.query.get_name(element)) 
        elif type(element) == DB.FamilySymbol:
            value += "ТИП - "
            value += "{}({})".format(element.FamilyName, revit.query.get_name(element))
        else: 
            value += revit.query.get_name(element)
    except Exception as e:
       value += str(e)
    return value

if parameter_to_set != "":
    with io.open('Y:\\Жилые здания\\ДИВНОЕ СитиXXI\\7. Стадия РД\\Исходные данные\\Классификатор City-ХХI век\\CITY_Классификатор_csv.txt', "r") as f:
        lines = f.readlines()
        for i in lines:
            parts = i.split(';')
            if("Item" in parts[2]):
                items.append([parts[0], parts[1]])
            else:
                categories.append([parts[0], parts[1]])

    out = script.get_output()
    num = 0;
    with db.Transaction(name = "Копировать описание"):
        with forms.ProgressBar(title='Анализ {value} из {max_value}', step = 100) as pb:
            for category in doc.Settings.Categories:
                if not category.IsTagCategory:
                    for element in DB.FilteredElementCollector(doc).OfCategoryId(category.Id).ToElements():
                        current += 1
                        pb.update_progress(current, max_value = max)
                        for j in element.Parameters:
                            if j.IsShared and j.Definition.Name == "ДИВ_WBS_Код по классификатору":
                                value = j.AsString()
                                description = GetCategory(value, items)
                                if description == None:
                                    description = GetCategory(value, categories)
                                    if description != None:
                                        num+=1
                                        print("\n[#{}]{}: {}\nНедопустимое значение кода WBS - {}".format(str(num), out.linkify(element.Id), GetName(element), description))
                                    else:
                                        num+=1
                                        print("\n[#{}]{}: {}\nЗначение WBS отсутствует в таблице - {}".format(str(num), out.linkify(element.Id), GetName(element), description))
                                else:
                                    try:
                                        for p in element.Parameters:
                                            if p.IsShared and p.Definition.Name == parameter_to_set:
                                                p.Set(description)
                                    except Exception as e:
                                        print(str(e))
                                