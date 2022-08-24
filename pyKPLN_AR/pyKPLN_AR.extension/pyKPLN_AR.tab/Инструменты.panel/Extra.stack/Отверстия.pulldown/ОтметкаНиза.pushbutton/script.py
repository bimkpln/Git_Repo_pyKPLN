# -*- coding: utf-8 -*-


__author__ = 'Tsimafei Kutsko'
__title__ = "Высотная отметка"
__doc__ = 'Запись значений высоты проема (относительная и абсолютная)\n' \
    '      «00_Отметка_Относительная» - высота проема относительно уровня ч.п. 1-го этажа\n' \
    '      «00_Отметка_Абсолютная» - высота проема относительно ч.п. связанного уровня\n' \


import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, \
    BuiltInParameter, Category, ParameterType, InstanceBinding, Transaction,\
    BuiltInParameterGroup
import math
import os
from System import Guid
from System.Collections.Generic import *
from rpw.ui.forms import Alert


def GetHeight(element):
    try:
        return element.get_Parameter(guidDiamParam).AsDouble()
    except AttributeError:
        return element.get_Parameter(guidHeightParam).AsDouble()


def GetDescription(length_feet):
    comma = "."
    if round(length_feet, 5) >= 0:
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


# Основные переменные
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
comParamsFilePath = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt"
trueCategory = BuiltInCategory.OST_MechanicalEquipment
elemsColl = FilteredElementCollector(doc).\
    OfCategory(trueCategory).\
    WhereElementIsNotElementType()
paramsList = ["00_Отметка_Относительная", "00_Отметка_Абсолютная"]
guidHeightParam = Guid("da753fe3-ecfa-465b-9a2c-02f55d0c2ff1")
guidDiamParam = Guid("9b679ab7-ea2e-49ce-90ab-0549d5aa36ff")
basePoint = FilteredElementCollector(doc).\
    OfCategory(BuiltInCategory.OST_ProjectBasePoint).\
    WhereElementIsNotElementType().\
    FirstElement()
basePointElev = basePoint.get_BoundingBox(None).Min.Z


# Подгружаю параметры
if os.path.exists(comParamsFilePath):
    try:
        # Создаю спец класс CategorySet и добавляю в него зависимости
        # (категории)
        catSetElements = app.Create.NewCategorySet()
        catSetElements.Insert(doc.Settings.Categories.get_Item(trueCategory))

        # Забираю все парамтеры проекта в список
        prjParamsNamesList = []
        paramBind = doc.ParameterBindings
        fIterator = paramBind.ForwardIterator()
        fIterator.Reset()
        while fIterator.MoveNext():
            prjParamsNamesList.append(fIterator.Key.Name)

        # Забираю все парамтеры проекта в список
        prjParamsNamesList = []
        paramBind = doc.ParameterBindings
        fIterator = paramBind.ForwardIterator()
        fIterator.Reset()
        while fIterator.MoveNext():
            d_Definition = fIterator.Key
            d_Name = fIterator.Key.Name
            d_Binding = fIterator.Current
            d_catSet = d_Binding.Categories
            if d_Name in paramsList\
                    and d_Binding.GetType() == InstanceBinding\
                    and d_Definition.ParameterType == ParameterType.Text\
                    and d_catSet.Contains(
                        Category.GetCategory(
                            doc,
                            trueCategory
                        )
                    ):
                prjParamsNamesList.append(fIterator.Key.Name)

        # Забираю все параметры из спец ФОПа
        app.SharedParametersFilename = comParamsFilePath
        sharedParamsFile = app.OpenSharedParameterFile()

        # Добавляю недостающие парамтеры в проект
        with Transaction(doc, 'КП_Добавить параметры') as t:
            t.Start()
            for defGroups in sharedParamsFile.Groups:
                if defGroups.Name == "АР_Отверстия":
                    for extDef in defGroups.Definitions:
                        # Добавляю параметры (если они не были ранее загружены)
                        if extDef.Name not in prjParamsNamesList:
                            paramBind = doc.ParameterBindings
                            newIB = app.\
                                Create.\
                                NewInstanceBinding(catSetElements)
                            paramBind.Insert(
                                extDef,
                                newIB,
                                BuiltInParameterGroup.PG_DATA
                            )

                            # Разворачиваю проход по параметрам проекта
                            revFIterator = doc.\
                                ParameterBindings.\
                                ReverseIterator()
                            while revFIterator.MoveNext():
                                if extDef.Name == revFIterator.Key.Name:

                                    # Включаю вариативность между экземплярами
                                    # групп в Revit
                                    revFIterator.Key.SetAllowVaryBetweenGroups(
                                        doc,
                                        True
                                    )
                                    break
            t.Commit()
    except Exception as e:
        Alert(
            "Ошибка при загрузке параметров:\n[{}]".format(str(e)),
            title="Загрузчик параметров", header="Ошибка"
        )
else:
    Alert(
        "Файл общих параметров не найден:{}".format(comParamsFilePath),
        title="Загрузчик параметров",
        header="Ошибка"
    )


def SetDiscr(elem, baseElem, isRound, isRelative):
    """Относительная и абсолютная отметки"""
    if isRound:
        elemHeight = GetHeight(elem) / 2
        prefDiscr = "Центр на отм. "
    else:
        elemHeight = 0
        prefDiscr = "Низ на отм. "
    try:
        boundExpand = elem.LookupParameter("Расширение границ").AsDouble()
    except AttributeError:
        boundExpand = 0

    # Относительная отметка
    if isRelative:
        sufDiscr = " мм от ур.ч.п."
        trueElevation = baseElem.\
            get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).\
            AsDouble()
        value = prefDiscr +\
            GetDescription(trueElevation + elemHeight - boundExpand) +\
            sufDiscr
        elem.LookupParameter("00_Отметка_Относительная").Set(value)

    # Абсолютная отметка
    else:
        sufDiscr = " мм"
        try:
            downOffset = elem.LookupParameter("SYS_OFFSET_DOWN").AsDouble()
        # Отработка старых параметров
        except AttributeError:
            downOffset = element.LookupParameter("offset_down").AsDouble()
        bBox = elem.get_BoundingBox(None)
        try:
            frame = elem.LookupParameter("Наличник_Ширина").AsDouble()
        except AttributeError:
            frame = 0
        maxExpand = max(downOffset, boundExpand, frame)
        if maxExpand == boundExpand:
            trueElevation = bBox.Min.Z + frame
        elif maxExpand == frame:
            trueElevation = bBox.Min.Z + maxExpand
        else:
            trueElevation = bBox.Min.Z + downOffset - boundExpand
        value = prefDiscr +\
            GetDescription(trueElevation + elemHeight - basePointElev) +\
            sufDiscr
        elem.LookupParameter("00_Отметка_Абсолютная").Set(value)


# Основная часть скрипта
with Transaction(doc, 'КП_Запись отметок') as t:
    t.Start()

    for element in elemsColl:
        if element.SuperComponent is not None:
            baseElement = element.SuperComponent
            if baseElement.SuperComponent is not None:
                baseElement = baseElement.SuperComponent
        else:
            baseElement = element
        try:
            famName = element.Symbol.FamilyName

            # Прямоугольные отверстия
            if famName.startswith("199_Отверстие прямоугольное")\
                    or famName == "199_AR_OSW"\
                    or famName.startswith("199_Отверстие в стене прямоугольное")\
                    or famName == ("Отверстие в стене под лючок"):
                # Запись относительной отметки
                SetDiscr(element, baseElement, False, True)
                # Запись абсолютной отметки
                SetDiscr(element, baseElement, False, False)

            # Круглые отверстия
            if famName.startswith("199_Отверстие круглое")\
                    or famName.startswith("199_Отверстие в стене круглое")\
                    or famName == "199_AR_ORW":
                # Запись относительной отметки
                SetDiscr(element, baseElement, True, True)
                # Запись абсолютной отметки
                SetDiscr(element, baseElement, True, False)
        except Exception as e:
            print(str(e))

    t.Commit()
