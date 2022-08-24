# -*- coding: utf-8 -*-
"""
ROOMS_SharedParameters

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Параметры"
__doc__ = 'Загрузка опционального параметра из ФОП и заполнение.\n' \
          'X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt' \
"""
Архитекурное бюро KPLN

"""
from pyrevit.framework import clr
import re
import os
from Autodesk.Revit.DB import BuiltInCategory, Transaction,\
    BuiltInParameterGroup, ParameterType, FilteredElementCollector
from rpw.ui.forms import Alert


def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
trueCategory = BuiltInCategory.OST_Rooms
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
# Список параметров, которые не меняются по экземплярам группы
dropAllowVaryList = ["ПОМ_Корпус", "ПОМ_Секция", "ПОМ_Номер квартиры"]
comParamsFilePath = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt"
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
            prjParamsNamesList.append(fIterator.Key.Name)

        # Забираю все параметры из спец ФОПа
        app.SharedParametersFilename = comParamsFilePath
        sharedParamsFile = app.OpenSharedParameterFile()

        # Добавляю недостающие парамтеры в проект
        with Transaction(doc, 'КП_Добавить параметры') as t:
            t.Start()
            for defGroups in sharedParamsFile.Groups:
                if defGroups.Name == "АР_Квартирография":
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
                                if extDef.Name == revFIterator.Key.Name\
                                        and extDef.Name in dropAllowVaryList:
                                    # Выключаю вариативность между экземплярами
                                    # групп в Revit
                                    revFIterator.Key.SetAllowVaryBetweenGroups(
                                        doc,
                                        False
                                    )
                                    break
                                elif extDef.Name == revFIterator.Key.Name\
                                        and extDef.ParameterType == ParameterType.Text:
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


# Запись значения в составной параметр
with Transaction(doc, 'Записать значение') as t:
    t.Start()
    try:
        for r in rooms:
            a = r.LookupParameter('ПОМ_Секция').AsString()
            b = r.LookupParameter('Назначение').AsString()
            if a and b != None:
                value = 'Секция ' + a + '. ' + b
                r.LookupParameter('BIM_Составной параметр').Set(value)
    except Exception as e:
        print(e)
    t.Commit()
