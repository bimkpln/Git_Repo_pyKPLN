# -*- coding: utf-8 -*-
"""
KPLN:DIV:CYRILLICSCOPER

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "DIV CYRILLIC SCOPER"
__doc__ = 'Поиск кириллических символов в параметрах WBS, RBS' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

import wpf
import re
import System
from rpw import doc, uidoc, DB, UI, db, ui
from pyrevit import framework
from pyrevit import script
from pyrevit import forms
from pyrevit import revit, DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert
import datetime
import webbrowser

max = 0
for category in doc.Settings.Categories:
    if not category.IsTagCategory:
        max += DB.FilteredElementCollector(doc).OfCategoryId(category.Id).ToElements().Count
current = 0

def InvalidRBS(element):
    value = "None"
    try:
        for j in element.Parameters:
            if j.IsShared and j.Definition.Name == "ДИВ_RBS_Код по классификатору":
                value = j.AsString()
                if value != None:
                    if value.upper().startswith("НK") or value.upper().startswith("HК") or value.upper().startswith("НК"):
                        return [True, value]
    except Exception as e:
        print(str(e))
    return [False, value]

def InvalidWBS(element):
    value = "None"
    try:
        for j in element.Parameters:
            if j.IsShared and j.Definition.Name == "ДИВ_WBS_Код по классификатору":
                value = j.AsString()
                if value != None:
                    if value.upper().startswith("В") or value.upper().startswith("С"):
                        return [True, value]
    except Exception as e:
        print(str(e))
    return [False, value]

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


out = script.get_output()
num = 0;
with forms.ProgressBar(title='Анализ {value} из {max_value}', step = 100) as pb:
    for category in doc.Settings.Categories:
        if not category.IsTagCategory:
            for element in DB.FilteredElementCollector(doc).OfCategoryId(category.Id).ToElements():
                current += 1
                pb.update_progress(current, max_value = max)
                rslt = InvalidRBS(element)
                if rslt[0]:
                    num+=1
                    print("\n[#{}]{}: {}\nКириллические символы в параметре «ДИВ_RBS_Код по классификатору» - {}".format(str(num), out.linkify(element.Id), GetName(element), rslt[1]))
                rslt = InvalidWBS(element)
                if rslt[0]:
                    num+=1
                    print("\n[#{}]{}: {}\nКириллические символы в параметре «ДИВ_WBS_Код по классификатору» - {}".format(str(num), out.linkify(element.Id), GetName(element), rslt[1]))