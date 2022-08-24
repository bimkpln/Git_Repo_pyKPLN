# -*- coding: utf-8 -*-
"""
OpeningsHeight

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Высотная отметка"
__doc__ = 'Запись значений высоты проема (относительная и абсолютная)\n' \
    '      «00_Отметка_Относительная» - высота проема относительно уровня ч.п. 1-го этажа\n' \
    '      «00_Отметка_Абсолютная» - высота проема относительно ч.п. связанного уровня\n' \

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, \
    BuiltInParameter, Category, StorageType, InstanceBinding, Transaction,\
    BuiltInParameterGroup
import math
from System import Guid
from rpw import doc, ui
from System.Collections.Generic import *

def GetDescription(length_feet):
    comma = "."
    if round(length_feet, 1) >= 0.0:
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

def LoadParams(params, params_exist):
    try:
        app = doc.Application
        category_set_elements = app.Create.NewCategorySet()
        category_set_elements.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Windows))
        category_set_elements.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Doors))
        category_set_elements.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment))
        app.SharedParametersFilename = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt"
        SharedParametersFile = app.OpenSharedParameterFile()
        map = doc.ParameterBindings
        it = map.ForwardIterator()
        it.Reset()
        with Transaction(doc, 'KPLN_Отметки. Добавление параметров') as t:
            t.Start()
            while it.MoveNext():
                d_Definition = it.Key
                d_Name = it.Key.Name
                d_Binding = it.Current
                d_catSet = d_Binding.Categories
                for param, bool in zip(params, params_exist):
                    if d_Name == param:
                        if d_Binding.GetType() == InstanceBinding:
                            if str(d_Definition.ParameterType) == "Text":
                                if d_Definition.VariesAcrossGroups:
                                    if d_catSet.Contains(Category.GetCategory(doc, BuiltInCategory.OST_Windows)) and d_catSet.Contains(Category.GetCategory(doc, BuiltInCategory.OST_Doors)) and d_catSet.Contains(Category.GetCategory(doc, BuiltInCategory.OST_MechanicalEquipment)):
                                        bool == True
            for dg in SharedParametersFile.Groups:
                if dg.Name == "АР_Отверстия":
                    for param, bool in zip(params, params_exist):
                        if not bool:
                            externalDefinition = dg.Definitions.get_Item(param)
                            newIB = app.Create.NewInstanceBinding(category_set_elements)
                            doc.ParameterBindings.Insert(externalDefinition, newIB, BuiltInParameterGroup.PG_DATA)
                            doc.ParameterBindings.ReInsert(externalDefinition, newIB, BuiltInParameterGroup.PG_DATA)
            map = doc.ParameterBindings
            it = map.ForwardIterator()
            it.Reset()
            while it.MoveNext():
                for param in params:
                    d_Definition = it.Key
                    d_Name = it.Key.Name
                    d_Binding = it.Current
                    if d_Name == param:
                        d_Definition.SetAllowVaryBetweenGroups(doc, True)
            t.Commit()
    except Exception as e:
        print(str(e))

# Основные переменные
collector_elements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType()
paramsList = ["00_Отметка_Относительная", "00_Отметка_Абсолютная"]
paramsExist = [False, False]
project_base_point = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().FirstElement()
bp_height = project_base_point.get_BoundingBox(None).Min.Z
guid_btm_hole = Guid("e4fef752-7335-4c6f-b91d-1ed181beaf3d") # КП_А_Вырез под проемом
hole_param_names = ["Вырезание под проемом", "Вырезание пола снизу"]
sill_param_names = ["Подоконник_Высота", "Высота подоконника"]
el_param_list = []
symb_param_list = []
success = bool

# Проверка наличия параметров
try:
    collector_elements.FirstElement().LookupParameter("00_Отметка_Абсолютная").AsString()
    collector_elements.FirstElement().LookupParameter("00_Отметка_Относительная").AsString()
except:
    LoadParams(paramsList, paramsExist)

# Проверка версии семейства
def FindSharedParameter(element):
    try:
        if element.get_Parameter(guid_btm_hole).AsDouble() != None:
            return True
    except:
        return False

# Поиск параметра высоты подоконника (присутствует только в Г-образных блоках)
def FindSill(element):
    try:
        try:
            if element.Symbol.LookupParameter(sill_param_names[0]).AsDouble() != None:
                return True
        except:
            if element.Symbol.LookupParameter(sill_param_names[1]).AsDouble() != None:
                return True
    except:
        return False

# Основная часть скрипта
with Transaction(doc, 'KPLN_Отметки. Запись отметок') as t:
    t.Start()
    try:
        for element in collector_elements:
            if element.SuperComponent != None:
                baseElement = element.SuperComponent
                if baseElement.SuperComponent != None:
                    baseElement = baseElement.SuperComponent
            else:
                baseElement = element
            famName = element.Symbol.FamilyName
            if famName.startswith("120_Окно") or famName.startswith("121_Блок") or famName.startswith("120_Блок"):
                b_box = element.get_BoundingBox(None)
                window_base = element.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).AsDouble()
                for i in element.Parameters:
                    el_param_list.append(i.Definition.Name)
                for i in element.Symbol.Parameters:
                    symb_param_list.append(i.Definition.Name)

                if FindSharedParameter(element):
                    window_btm_hole = element.get_Parameter(guid_btm_hole).AsDouble()
                    temp = 0
                else:
                    try:
                        window_btm_hole = element.LookupParameter("Вырезание под проемом").AsDouble()
                        temp = window_btm_hole
                    except:
                        window_btm_hole = element.LookupParameter("Вырезание пола снизу").AsDouble()
                        temp = window_btm_hole

                if FindSill(element):
                    try:
                        window_sill = element.Symbol.LookupParameter("Подоконник_Высота").AsDouble()
                    except:
                        window_sill = element.Symbol.LookupParameter("Высота подоконника").AsDouble()

                    value = "Низ в зоне окна на отм. " + GetDescription(window_base + window_sill) + " мм от ур.ч.п., низ в зоне двери на отм. " + GetDescription(window_base - window_btm_hole) + " мм от ур.ч.п."
                    abs_value = "Низ в зоне окна на отм. " + GetDescription(b_box.Min.Z - bp_height + window_sill + window_btm_hole - 2*temp) + " мм, низ в зоне двери на отм. " + GetDescription(b_box.Min.Z - bp_height - temp) + " мм"
                    element.LookupParameter("00_Отметка_Относительная").Set(value)
                    element.LookupParameter("00_Отметка_Абсолютная").Set(abs_value)
                else:
                    value = "Низ на отм. " + GetDescription(window_base - window_btm_hole) + " мм от ур.ч.п."
                    abs_value = "Низ на отм. " + GetDescription(b_box.Min.Z - bp_height - temp) + " мм"
                    element.LookupParameter("00_Отметка_Относительная").Set(value)
                    element.LookupParameter("00_Отметка_Абсолютная").Set(abs_value)
        success = True
    except Exception as e:
        print(str(e))
    t.Commit()

if success:
    ui.forms.Alert("Значения высотных отметок записаны.", title = "Готово!")