# -*- coding: utf-8 -*-
"""
OpeningsHeight

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Высотн. отм.\nОТВЕРСТИЙ в стенах_v.1"
__doc__ = 'Запись значений высоты задания на отверстия в параметр «Комментарии»\n (для семейств из менеджера отверстий)' \

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

def getStringHeight(length_feet):
    comma = "."
    if length_feet >= 0:
        sign = "+"
    else:
        sign = "-"
    length_feet_abs = math.fabs(length_feet)
    length_meters = int(round(length_feet_abs * 304.8 / 5, 0) * 5)
    length_string = str(length_meters)
    if len(length_string) == 7:
        value = length_string[:4] + comma + length_string[4:]
    elif len(length_string) == 6:
        value = length_string[:3] + comma + length_string[3:]
    elif len(length_string) == 5:
        value = length_string[:2] + comma + length_string[2:]
    elif len(length_string) == 4:
        value = length_string[:1] + comma + length_string[1:]
    elif len(length_string) == 3:
        value = "0{}".format(comma) + length_string
    elif len(length_string) == 2:
        value = "0{}0".format(comma) + length_string
    elif len(length_string) == 1:
        value = "0{}00".format(comma) + length_string
    value = sign + value
    return value

collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()

with db.Transaction(name='Global height'):
    for element in collector_elements:
        try:
            fam_name = element.Symbol.FamilyName 
            if fam_name.startswith("501_Задание на отверстие в стене прямоугольное") or fam_name.startswith("501_MEP_TSW"):
                offset = element.get_Parameter(DB.BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble()
                height = element.LookupParameter("Высота").AsDouble()
                extend = element.LookupParameter("Расширение границ").AsDouble()
                elevation = doc.GetElement(element.LevelId).Elevation
                value = elevation + offset - extend
                textValue = "Отм. низа: " + getStringHeight(value) + " мм от нуля здания"
                element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(textValue)

            if fam_name.startswith("501_Задание на отверстие в стене круглое") or fam_name.startswith("501_MEP_TRW"):
                offset = element.get_Parameter(DB.BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble()
                height = element.LookupParameter("Высота").AsDouble()
                elevation = doc.GetElement(element.LevelId).Elevation
                value = elevation + offset + height / 2
                textValue = "Отм. центра: " + getStringHeight(value) + " мм от нуля здания"
                element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(textValue)
        except Exception as e:
            print(e)