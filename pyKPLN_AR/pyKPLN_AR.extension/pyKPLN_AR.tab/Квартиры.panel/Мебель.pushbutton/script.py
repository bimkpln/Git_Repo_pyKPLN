# -*- coding: utf-8 -*-
"""
FurnitureChanger

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Мебель"
__doc__ = 'Заменяет выбранные на виде 2D семейства мебели и сантехники на трехмерные с сохранением положения'\

"""
Архитекурное бюро KPLN

"""
from rpw import doc, uidoc, DB, UI, ui
from rpw.ui.forms import CommandLink, TaskDialog, Alert
from Autodesk.Revit.DB import BuiltInParameter, ParameterType,\
    FilteredElementCollector, BuiltInCategory, Transaction,\
    Category, ElementLevelFilter, XYZ, Line, ElementTransformUtils
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import forms
from pyrevit import script

newElements = []
angleList =[]
structuralType = StructuralType.NonStructural
warning = bool
error = bool
succes = bool
warningIds = []
errorIds = []
dblParameters = []
dblValues = []
intParameters = []
intValues = []
categories = [BuiltInCategory.OST_Furniture, BuiltInCategory.OST_PlumbingFixtures]

class MassSelectorFilter (UI.Selection.ISelectionFilter):
    def AllowElement(self, element=DB.Element):
        if element.Category.Id.IntegerValue == -2001160 or element.Category.Id.IntegerValue == -2000080: # Формирование условия выбора (только мебель и сантехника)
            if element.Symbol.Family.get_Parameter(BuiltInParameter.FAMILY_SHARED).AsInteger() == 0: # Игнорирование общих семейств
                return True
        else:
            return False
    def AllowReference(self, refer, point):
        return False

def PickElement():
    elements = []
    succes = False
    try:
        pick_filter = MassSelectorFilter()
        with forms.WarningBar(title='Выберите элементы на виде:'):
            ref_elements = uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element, pick_filter, "Выберите элементы на виде")
        for i in ref_elements:
            try:
                elements.append(doc.GetElement(i))
                succes = True
            except: 
                pass
    except:
        pass
    return elements

# Окно выбора
commands= [CommandLink('Заменить 2D на 3D', return_value='2D_to_3D'),
            CommandLink('Заменить 3D на 2D', return_value='3D_to_2D')]
dialog = TaskDialog('Выберите на виде экземппляры семейств мебели и сантехники, которые необходимо заменить.',
               title="Заменить семейства",
               title_prefix=False,
               content = "ВНИМАНИЕ! Рекомендуется предварительно извлечь экземпляры семейств из групп, иначе при замене группы будут расформированы.",
               commands=commands,
               show_close=True)
dialog_out = dialog.show()

# Определение текущих настроек выбранных семейств
pickedElements = PickElement()
for i in range(0, len(pickedElements)):
    paramList = pickedElements[i].Parameters
    for param in paramList:
        if str(param.StorageType) == 'Double' and param.HasValue == True and param.IsReadOnly == False:
            dblParameters.append(param.Definition.Name)
        if str(param.StorageType) == 'Integer' and param.HasValue == True and param.IsReadOnly == False:
            intParameters.append(param.Definition.Name)
    for paramName in dblParameters:
        try:
            value = pickedElements[i].LookupParameter(paramName).AsDouble()
            elem = i
            name = paramName
            dblValues.append([elem, name, value])
        except:
            pass
    for paramName in intParameters:
        try:
            value = pickedElements[i].LookupParameter(paramName).AsInteger()
            elem = i
            name = paramName
            intValues.append([elem, name, value])
        except:
            pass

if dialog_out == '2D_to_3D':
    with Transaction(doc, 'Загрузить семейства') as t:
        t.Start()
        for element in pickedElements:
            famName = element.Symbol.FamilyName
            try:
                if famName.startswith("140_") and not '_3d' in famName:
                    doc.LoadFamily("X:\\BIM\\3_Семейства\\1_АР\\000_Архитектурная концепция\\140_Мебель\\{}_3d.rfa".format(famName))
                if famName.startswith("700_") and not '_3d' in famName:
                    doc.LoadFamily("X:\\BIM\\3_Семейства\\1_АР\\000_Архитектурная концепция\\700_Сантехнические приборы\\{}_3d.rfa".format(famName))
            except:
                pass
        t.Commit()

    with Transaction(doc, 'Заменить 2D семейства на 3D') as t:
        t.Start()
        symbols = []
        symbolNames = []
        symbolTypes = []
        symbolDict = []
        for category in categories:
            for symbol in DB.FilteredElementCollector(doc).OfCategory(category):
                try:
                    if '_3d' in symbol.FamilyName:
                        symbols.append(symbol)
                        symbolNames.append(symbol.FamilyName)
                        symbolTypes.append(symbol.LookupParameter("Имя типа").AsString())
                        symbolDict.append(str(symbol.FamilyName + ' : ' + symbol.LookupParameter("Имя типа").AsString()))
                except:
                    pass
        # Сопоставление семейств
        for element in pickedElements:
            elemType = element.Name
            elemName = element.Symbol.FamilyName
            location = element.Location.Point
            level = doc.GetElement(element.LevelId)
            angle = element.Location.Rotation
            angleList.append(angle)
            if '_3d' in elemName:
                famSymbol = element.Symbol
            else:
                try:
                    if '{}_3d'.format(elemName) in symbolNames:
                        index = symbolNames.index('{}_3d'.format(elemName))
                        famType = elemType
                        warning = True
                        if elemType in symbolTypes:
                            type = str('{}_3d'.format(elemName) + ' : ' + elemType)
                            for i in symbolDict:
                                if type == i:
                                    index = symbolDict.index(i)
                                    warning = False
                        if '{}_3d'.format(elemType) in symbolTypes:
                            type = str('{}_3d'.format(elemName) + ' : ' + '{}_3d'.format(elemType))
                            for i in symbolDict:
                                if type == i:
                                    index = symbolDict.index(i)
                                    warning = False
                        famSymbol = symbols[index]
                        famSymbol.Activate()
                        error = False
                    else:
                        famSymbol = element.Symbol
                        error = True
                except:
                    famSymbol = element.Symbol
                    error = True
                # Удаление 2д-семейств
                doc.Delete(element.Id)
                # Размещение 3д-семейств по тому же месту
                try:
                    new_element = doc.Create.NewFamilyInstance(location, famSymbol, structuralType)
                    newElements.append(new_element)
                except Exception as e:
                    print(e)
                # Формирование списка ненайденных типоразмеров
                if warning == True:
                    warningIds.append([new_element.Id, famType])
                if error == True:
                    errorIds.append([new_element.Id, elemName])
        t.Commit()

if dialog_out == '3D_to_2D':
    with Transaction(doc, 'Заменить 3D семейства на 2D') as t:
        t.Start()
        symbols = []
        symbolNames = []
        symbolTypes = []
        symbolDict = []
        for category in categories:
            for symbol in DB.FilteredElementCollector(doc).OfCategory(category):
                try:
                    if '_3d' not in symbol.FamilyName:
                        symbols.append(symbol)
                        symbolNames.append(symbol.FamilyName)
                        symbolTypes.append(symbol.LookupParameter("Имя типа").AsString())
                        symbolDict.append(str(symbol.FamilyName + ' : ' + symbol.LookupParameter("Имя типа").AsString()))
                except:
                    pass
        # Сопоставление семейств
        for element in pickedElements:
            elemType = element.Name
            elemName = element.Symbol.FamilyName
            location = element.Location.Point
            level = doc.GetElement(element.LevelId)
            angle = element.Location.Rotation
            angleList.append(angle)
            if '_3d' not in elemName:
                famSymbol = element.Symbol
            else:
                try:
                    if elemName[:-3] in symbolNames:
                        index = symbolNames.index(elemName[:-3])
                        famType = elemType
                        warning = True
                        if elemType in symbolTypes:
                            type = str(elemName[:-3] + ' : ' + elemType)
                            for i in symbolDict:
                                if type == i:
                                    index = symbolDict.index(i)
                                    warning = False
                        if elemType[:-3] in symbolTypes:
                            type = str(elemName[:-3] + ' : ' + elemType[:-3])
                            for i in symbolDict:
                                if type == i:
                                    index = symbolDict.index(i)
                                    warning = False
                        famSymbol = symbols[index]
                        famSymbol.Activate()
                        error = False
                    else:
                        famSymbol = element.Symbol
                        error = True
                except:
                    famSymbol = element.Symbol
                    error = True
                # Удаление 3д-семейств
                doc.Delete(element.Id)
                # Размещение 2д-семейств по тому же месту
                try:
                    new_element = doc.Create.NewFamilyInstance(location, famSymbol, structuralType)
                    newElements.append(new_element)
                except Exception as e:
                    print(e)
                # Формирование списка ненайденных типоразмеров
                if warning == True:
                    warningIds.append([new_element.Id, famType])
                if error == True:
                    errorIds.append([new_element.Id, elemName])
        t.Commit()

# Поворот вновь созданных экземпляров в исходное положение
with Transaction(doc, 'Восстановить положение') as t:
    t.Start()
    try:
        for i in range(0, len(newElements)):
            angle = angleList[i]
            center = newElements[i].Location.Point
            p1 = XYZ(center[0], center[1], 0)
            p2 = XYZ(center[0], center[1], 1)
            axis = Line.CreateBound(p1, p2)
            ElementTransformUtils.RotateElement(doc, newElements[i].Id, axis, angle)
    except:
        pass
    t.Commit()

# Восстановление значений параметров экземпляров
with Transaction(doc, 'Восстановить настройки') as t:
    t.Start()
    try:
        for i in range(0, len(newElements)):
            for value in dblValues:
                if i == value[0]:
                    newElements[i].LookupParameter(value[1]).Set(value[2])
            for value in intValues:
                if i == value[0]:
                    newElements[i].LookupParameter(value[1]).Set(value[2])
    except:
        pass
    t.Commit()

output = script.get_output()

# Создание и вызов сообщения
try:
    try:
        for i in errorIds:
            output.print_html('Семейство <b><span style="color:FireBrick;">{}</span></b> '.format(i[1]) + output.linkify(i[0]) +  ' отсутствует в 3D библиотеке.')
            flag = True
        if flag:
            print('Данные семейства не были заменены!')
    except:
        pass
    for i in warningIds:
        output.print_html('Типоразмер <b><span style="color:FireBrick;">{}</span></b> '.format(i[1]) + output.linkify(i[0]) +  ' отсутствует в 3D библиотеке.')
        flag = True
    if flag:
        print('Данные экземпляры были заменены без сохранения настроек, необходимо скорректировать типоразмеры самостоятельно!')
except:
    if succes == True:
        ui.forms.Alert("Семейства были успешно заменены.", title = "Готово!")
    else:
        ui.forms.Alert("Действие отменено. Замена семейств не выполнена.", title = "Внимание!")
