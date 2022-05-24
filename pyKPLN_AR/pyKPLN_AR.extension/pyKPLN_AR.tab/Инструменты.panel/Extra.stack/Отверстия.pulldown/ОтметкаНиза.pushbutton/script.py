# -*- coding: utf-8 -*-
"""
OpeningsHeight

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
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
from rpw import doc
from System.Collections.Generic import *

def GetHeight(element):
    values = ["КП_Р_Высота", "INGD_Длина"]
    for j in element.Parameters:
        if j.StorageType == StorageType.Double:
            if j.Definition.Name in values:
                return j.AsDouble()
    return None

def GetDescription(length_feet):
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

def LoadParams(params, params_exist):
    try:
        app = doc.Application
        category_set_elements = app.Create.NewCategorySet()
        category_set_elements.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Windows))
        category_set_elements.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Doors))
        category_set_elements.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment))
        originalFile = app.SharedParametersFilename
        app.SharedParametersFilename = "Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\lib\\ФОП_Scripts.txt"
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
                if dg.Name == "АРХИТЕКТУРА - Дополнительные":
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
collector_elements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType()
paramsList = ["00_Отметка_Относительная", "00_Отметка_Абсолютная"]
paramsExist = [False, False]
default_offset_bp = 0.00
project_base_point = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().FirstElement()
bp_height = project_base_point.get_BoundingBox(None).Min.Z

# Проверка наличия параметров
try:
    collector_elements.FirstElement().LookupParameter("00_Отметка_Абсолютная").AsString()
    collector_elements.FirstElement().LookupParameter("00_Отметка_Относительная").AsString()
except:
    LoadParams(paramsList, paramsExist)

# Основная часть скрипта
with Transaction(doc, 'KPLN_Отметки. Запись отметок') as t:
    t.Start()

    # Запись относительной отметки
    for element in collector_elements:
        if element.SuperComponent != None:
            baseElement = element.SuperComponent
            if baseElement.SuperComponent != None:
                baseElement = baseElement.SuperComponent
        else:
            baseElement = element
        try:
            famName = element.Symbol.FamilyName
            if famName.startswith("199_Отверстие прямоугольное")\
                    or famName == "199_AR_OSW"\
                    or famName.startswith("199_Отверстие в стене прямоугольное")\
                    or famName == ("Отверстие в стене под лючок"):
                if element.LookupParameter("Расширение границ") != None:
                    bound_expand = element.LookupParameter("Расширение границ").AsDouble()
                    base_height = baseElement.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
                    value = "Низ на отм. " + GetDescription(base_height - bound_expand) + " мм от ур.ч.п."
                    element.LookupParameter("00_Отметка_Относительная").Set(value)
                else:
                    base_height = baseElement.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
                    value = "Низ на отм. " + GetDescription(base_height) + " мм от ур.ч.п."
                    element.LookupParameter("00_Отметка_Относительная").Set(value)
            if famName.startswith("199_Отверстие круглое") or famName == "199_AR_ORW":
                if element.LookupParameter("Расширение границ") != None:
                    bound_expand = element.LookupParameter("Расширение границ").AsDouble()
                    base_height = baseElement.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
                    element_height = GetHeight(element)/2
                    value = "Центр на отм. " + GetDescription(base_height + element_height - bound_expand) + " мм от ур.ч.п."
                    element.LookupParameter("00_Отметка_Относительная").Set(value)
                else:
                    base_height = baseElement.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble()
                    element_height = GetHeight(element)/2
                    value = "Центр на отм. " + GetDescription(base_height + element_height) + " мм от ур.ч.п."
                    element.LookupParameter("00_Отметка_Относительная").Set(value)
        except Exception as e:
            print(str(e))

    # Запись абсолютной отметки
    if element.SuperComponent != None:
        baseElement = element.SuperComponent
        if baseElement.SuperComponent != None:
            baseElement = baseElement.SuperComponent
    else:
        baseElement = element
    for element in collector_elements:
        try:
            famName = element.Symbol.FamilyName 
            if famName.startswith("199_Отверстие прямоугольное")\
                    or famName == "199_AR_OSW"\
                    or famName.startswith("199_Отверстие в стене прямоугольное")\
                    or famName == ("Отверстие в стене под лючок"):
                down = element.LookupParameter("offset_down").AsDouble()
                b_box = element.get_BoundingBox(None)
                try:
                    frame = element.LookupParameter("Наличник_Ширина").AsDouble()
                except:
                    frame = 0

                temp = max(down, bound_expand, frame)
                if temp == bound_expand or temp == frame:
                    boundingBox_Z_min = b_box.Min.Z + frame - default_offset_bp
                if temp == down:
                    boundingBox_Z_min = b_box.Min.Z + down - bound_expand - default_offset_bp

                value = "Низ на отм. " + GetDescription(boundingBox_Z_min - bp_height + frame) + " мм"
                element.LookupParameter("00_Отметка_Абсолютная").Set(value)
            if famName.startswith("199_Отверстие круглое"):
                down = element.LookupParameter("offset_down").AsDouble()
                up = element.LookupParameter("offset_up").AsDouble()
                b_box = element.get_BoundingBox(None)
                boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
                value = "Центр на отм. " + GetDescription(boundingBox_Z_center - bp_height) + " мм"
                element.LookupParameter("00_Отметка_Абсолютная").Set(value)
            if famName.startswith("501_Гильза_АР_Стена") or famName == "199_AR_ORW":#temp20200515
                down = element.LookupParameter("offset_down").AsDouble()
                up = element.LookupParameter("offset_up").AsDouble()
                b_box = element.get_BoundingBox(None)
                boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
                value = "Центр на отм. " + GetDescription(boundingBox_Z_center - bp_height) + " мм"
                element.LookupParameter("00_Отметка_Абсолютная").Set(value)
        except:
            famName = element.Symbol.FamilyName 
            if famName.startswith("199_Отверстие прямоугольное")\
                    or famName == "199_AR_OSW"\
                    or famName.startswith("199_Отверстие в стене прямоугольное")\
                    or famName == ("Отверстие в стене под лючок"):
                down = element.LookupParameter("SYS_OFFSET_DOWN").AsDouble()
                b_box = element.get_BoundingBox(None)
                bound_expand = element.LookupParameter("Расширение границ").AsDouble()
                try:
                    frame = element.LookupParameter("Наличник_Ширина").AsDouble()
                except:
                    frame = 0

                temp = max(down, bound_expand, frame)
                if temp == bound_expand or temp == frame:
                    boundingBox_Z_min = b_box.Min.Z + frame - default_offset_bp
                if temp == down:
                    boundingBox_Z_min = b_box.Min.Z + down - bound_expand - default_offset_bp

                value = "Низ на отм. " + GetDescription(boundingBox_Z_min - bp_height) + " мм"
                element.LookupParameter("00_Отметка_Абсолютная").Set(value)
            if famName.startswith("199_Отверстие круглое"):
                down = element.LookupParameter("SYS_OFFSET_DOWN").AsDouble()
                up = element.LookupParameter("SYS_OFFSET_UP").AsDouble()
                b_box = element.get_BoundingBox(None)
                boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
                value = "Центр на отм. " + GetDescription(boundingBox_Z_center - bp_height) + " мм"
                element.LookupParameter("00_Отметка_Абсолютная").Set(value)
            if famName.startswith("501_Гильза_АР_Стена") or famName == "199_AR_ORW":
                down = element.LookupParameter("SYS_OFFSET_DOWN").AsDouble()
                up = element.LookupParameter("SYS_OFFSET_UP").AsDouble()
                b_box = element.get_BoundingBox(None)
                boundingBox_Z_center = (b_box.Min.Z + down + b_box.Max.Z - up) / 2
                value = "Центр на отм. " + GetDescription(boundingBox_Z_center - bp_height) + " мм"
                element.LookupParameter("00_Отметка_Абсолютная").Set(value)

    t.Commit()
