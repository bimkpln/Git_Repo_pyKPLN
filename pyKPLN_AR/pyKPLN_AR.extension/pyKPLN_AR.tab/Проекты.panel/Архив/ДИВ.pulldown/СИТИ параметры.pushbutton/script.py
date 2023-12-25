# -*- coding: utf-8 -*-
"""
KPLN:DIV:CITYPARAMETERS

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Параметры СИТИ для АР"
__doc__ = '...' \

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

whitelist = []

exceptions = [-2000700, -2000576, -2000924, -2009609, -2001342,
              -2001272, -2001272, -2000500, -2000575, -2003101,
              -2003100, -2000279, -2000045, -2002000]

out = script.get_output()

categories = [
    DB.BuiltInCategory.OST_Rooms,#0
    DB.BuiltInCategory.OST_Walls,#1
    DB.BuiltInCategory.OST_Floors,#2
    DB.BuiltInCategory.OST_Ceilings,#3
    DB.BuiltInCategory.OST_Doors,#4
    DB.BuiltInCategory.OST_Windows,#5
    DB.BuiltInCategory.OST_CurtainWallMullions,#6 Импосты витража
    DB.BuiltInCategory.OST_StairsRailing,#7 Ограждения
    DB.BuiltInCategory.OST_CurtainWallPanels,#8 Панели витража
    DB.BuiltInCategory.OST_MechanicalEquipment,#9 Оборудование
    DB.BuiltInCategory.OST_GenericModel#10 Обобщенные модели
    ]

sections = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]

levels = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "м1"]

presets = []
#Исходный параметр - Целевой параметр - Список категорий (-1 - если на все категории) - Проверять на допустимое значение - Список допустимых знач. - Проверять на латинницу - Описание допустимого значения
presets.append(["ДИВ_RBS_Код по классификатору", "СИТИ_Классификатор", [-1], False, [], True, "", True])
presets.append(["ДИВ_Секция_Текст", "СИТИ_Секция", [-1], True, sections, False, "от 1-9", False])
presets.append(["ДИВ_Этаж_Текст", "СИТИ_Этаж", [-1], True, levels, False, "м1 либо 1-9", False])
presets.append(["ДИВ_Наименование по классификатору", "СИТИ_Тип", [-1], False, [], False, "", False])
presets.append(["ДИВ_Имя_классификатора", "СИТИ_Имя_классификатора", [-1], False, [], False, "", False])
presets.append(["ДИВ_Комплект_документации", "СИТИ_Комплект_документации", [-1], False, [], False, "", False])
presets.append(["КП_Марка_Номер", "СИТИ_Марка", [10, 9, 7, 6], False, [], False, "", False])
presets.append(["Маркировка типоразмера", "СИТИ_Марка", [1, 2, 3], False, [], False, "", False])
presets.append(["ADSK_Марка", "СИТИ_Марка", [4, 5], False, [], False, "", False])
presets.append(["КП_А_Маркировка типоразмера", "СИТИ_Марка", [4, 5], False, [], False, "", False])
presets.append(["Позиция", "СИТИ_Марка", [8], False, [], False, "", False])
presets.append(["ДИВ_Плоскость здания", "СИТИ_Плоскость здания", [1, 8], False, [], False, "", False])
presets.append(["ДИВ_Код помещения", "СИТИ_Тип помещения", [1, 2, 3], False, [], False, "", False])
presets.append(["ДИВ_Номер помещения", "СИТИ_Номер помещения", [1, 2, 3], False, [], False, "", False])
presets.append(["ADSK_Тип помещения", "СИТИ_Тип помещения", [0], False, [], False, "", False])
presets.append(["ПОМ_Номер квартиры", "СИТИ_Номер квартиры", [0], False, [], False, "", False])
presets.append(["ПОМ_Номер помещения", "СИТИ_Номер помещения", [0], False, [], False, "", False])
presets.append(["ДИВ_Описание по классификатору", "СИТИ_Описание", [-1], False, [], False, "", False])

def ParameterExist(element, parameter):
    for j in element.Parameters:
        if j.Definition.Name == parameter:
            return True
    return False

def ValueExist(element, parameter, check9999):
    value = element.LookupParameter(parameter).AsString()
    if value == None or value == "":
        return False
    if check9999:
        if value == "9999":
            if not element.Id.IntegerValue in whitelist:
                whitelist.append(element.Id.IntegerValue)
    return True

def ValueInList(element, parameter, values):
    value = element.LookupParameter(parameter).AsString()
    for v in values:
        if v == value:
            return True
    return False

def HasCyrillicChars(element, parameter):
    dict = "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
    for char in element.LookupParameter(parameter).AsString().lower():
        if char in dict:
            return True
    return False

def CheckResult(element, preset):
    if element == None: return [-1, None]
    for parameter in [preset[0], preset[1]]:
        valueMustBeInList = preset[3]
        values = preset[4]
        onlyLatinChars = preset[5]
        check999 = preset[6]
        if not ParameterExist(element, parameter):
            return [1, parameter]
        if not ValueExist(element, parameter, check999):
            return [2, parameter]
        if valueMustBeInList:
            if not ValueInList(element, parameter, values):
                return [3, parameter]
        if onlyLatinChars:
            if HasCyrillicChars(element, parameter):
                return [4, parameter]
    return [0, None]

errors = []

def AddError(element, message):
    if element == None: return
    error = None
    for e in errors:
        if e[0] == message:
            error = e
            break
    if error == None:
        errors.append([message, []])
    error = errors[-1]
    type = doc.GetElement(element.GetTypeId())
    if type == None:
        for part in error[1]:
            if part[0].IntegerValue == element.Id.IntegerValue:
                return
        error[1].append([element.Id, []])
        return
    else:
        for part in error[1]:
            if part[0].IntegerValue == type.Id.IntegerValue:
                part[1].append(element.Id)
                return
        error[1].append([type.Id, [element.Id]])
        return

def PrepareError(code, element, preset):
    global errors
    parameter = code[1]
    error = ""
    if code[0] == 1:
        error += "Отсутствует параметр «{0}»".format(parameter)
    elif code[0] == 2:
        error += "Отсутствует значение параметра «{0}»".format(parameter)
    elif code[0] == 3:
        condition = preset[6]
        error += "Значение параметра «{0}» не соответствует формату: [{1}]".format(parameter, condition)
    elif code[0] == 4:
        error += "Параметр «{0}» не должен содержать кирилических символов".format(parameter)
    else:
        return
    AddError(element, error)

def CheckElement(element, preset):
    if element == None: return False
    type = doc.GetElement(element.GetTypeId())
    if type != None:
        resultType = CheckResult(type, preset)
        resultElement = CheckResult(element, preset)
        #
        if resultElement[0] == 0 or resultType[0] == 0:
            return True
        else:
            if resultType[0] >= resultElement[0]:
                PrepareError(resultType, type, preset)
            else:
                PrepareError(resultElement, element, preset)
            return False
    else:
        resultElement = CheckResult(element, preset)
        #
        if resultElement[0] != 0:
            PrepareError(resultElement, element, preset)
            return False
        else: return True


def GetMax(current, max):
    result = current * 100 / max
    return result

def ListContainsType(t, list):
    for i in list:
        if i == t.Id.IntegerValue:
            return True
    return False

def Run():
    i = 0
    max = 0
    ammount = 0
    types = []
    for element in DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
        try:
            if element.Category.Id.IntegerValue in exceptions: continue
            type = doc.GetElement(element.GetTypeId())
            if type != None:
                elName = revit.query.get_name(type)
                if ".dwg" in elName or ".rvt" in elName: continue
                if type != None:
                    if not ListContainsType(type, types):
                        types.append(type.Id.IntegerValue)
        except:
            pass

    for preset in presets:
        for catId in preset[2]:
            if catId != -1:
                builtInCategory = categories[catId]
                max += DB.FilteredElementCollector(doc).OfCategory(builtInCategory).WhereElementIsNotElementType().ToElements().Count
            else:
                max += DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements().Count
    with forms.ProgressBar(title='{value}/{max_value}', step = 100, cancellable  = True) as pb:
        for preset in presets:
            for catId in preset[2]:
                if catId != -1:
                    builtInCategory = categories[catId]
                    for element in DB.FilteredElementCollector(doc).OfCategory(builtInCategory).WhereElementIsNotElementType().ToElements():
                        ammount += 1
                        if element.Id.IntegerValue in whitelist: continue
                        pb.update_progress(ammount, max_value = max)
                        if pb.cancelled:
                            return
                        CheckElement(element, preset)
                else:
                    for element in DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
                        type = doc.GetElement(element.GetTypeId())
                        if type != None:
                            if type.Id.IntegerValue in types:
                                ammount += 1
                                pb.update_progress(ammount, max_value = max)
                                if element.Category == None: continue
                                if element.Category.CategoryType != DB.CategoryType.Model: continue
                                if element.Category.IsTagCategory: continue
                                if element.Category.Id.IntegerValue in exceptions: continue
                                if pb.cancelled:
                                    return
                                CheckElement(element, preset)
Run()
for errorGroup in errors:
    msg = "☢ ERROR: {0}".format(errorGroup[0])
    for el in errorGroup[1]:
        element = doc.GetElement(el[0])
        s = ""
        if element.Category != None:
            s = "\n\t↳ [{0}] {1}: {2}".format(element.Category.Name, revit.query.get_name(element), out.linkify(element.Id))
        else:
            s = "\n\t↳ {0}: {1}".format(revit.query.get_name(element), out.linkify(element.Id))
        for id in el[1]:
            element = doc.GetElement(id)
            s += "\n\t\t↳ {0}: {1}".format(revit.query.get_name(element), out.linkify(element.Id))
    msg += s
    msg += "\n\n"
    print(msg)