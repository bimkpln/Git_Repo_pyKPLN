# -*- coding: utf-8 -*-
"""
ROOMS_SharedParameters

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Параметры"
__doc__ = 'Загрузка всех необходимых параметров для работы с квартирографией в проект.\n' \
          'X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt' \
"""
Архитекурное бюро KPLN

"""
from pyrevit.framework import clr
import re
import os
from Autodesk.Revit.DB import BuiltInCategory, Transaction,\
    BuiltInParameterGroup, ParameterType
from rpw.ui.forms import Alert


def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
# Список параметров, которые не меняются по экземплярам группы
dropAllowVaryList = ["ПОМ_Корпус", "ПОМ_Секция", "ПОМ_Номер квартиры"]
comParamsFilePath = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt"
# Подгружаю параметры
if os.path.exists(comParamsFilePath):
    try:
        # Создаю спец класс CategorySet для помещений
        catSetRooms = app.Create.NewCategorySet()
        catSetRooms.Insert(
            doc.
            Settings.
            Categories.
            get_Item(BuiltInCategory.OST_Rooms))

        # Создаю спец класс CategorySet для сведений о проекте
        catSetPrjInformation = app.Create.NewCategorySet()
        catSetPrjInformation.Insert(
            doc.
            Settings.
            Categories.
            get_Item(BuiltInCategory.OST_ProjectInformation))

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
        trueExtDef = []
        with Transaction(doc, 'КП_Добавить параметры') as t:
            t.Start()
            for defGroups in sharedParamsFile.Groups:
                if defGroups.Name == "АР_Квартирография":
                    for extDef in defGroups.Definitions:
                        # Добавляю параметры (если они не были ранее загружены)
                        if extDef.Name not in prjParamsNamesList:
                            trueExtDef.append(extDef)

            trueExtDef = sorted(trueExtDef, key=lambda t: t.Name)
            for extDef in trueExtDef:
                paramBind = doc.ParameterBindings
                if (extDef.Name.startswith("ТЭП_")):
                    newIB = app.\
                        Create.\
                        NewInstanceBinding(catSetPrjInformation)
                else:
                    newIB = app.\
                        Create.\
                        NewInstanceBinding(catSetRooms)

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
