# -*- coding: utf-8 -*-
"""
Button

"""
__author__ = 'Khlapov Dmitry'
__title__ = "Выбрать\nэлементы"
__doc__ = 'Пробная кнопка' \

"""
Архитекурное бюро KPLN

"""

import clr

from Autodesk.Revit.DB import FilteredElementCollector as FEC
from Autodesk.Revit.DB import ElementCategoryFilter, ElementMulticategoryFilter, BuiltInCategory, Wall, ElementId, FamilyInstance, WallType, Element
#from Autodesk.Revit.UI.Selection import Selection
from System.Collections.Generic import List
from System import Type
from Autodesk.Revit.DB import BuiltInParameter, Parameter, Category, ElementType

doc = __revit__.ActiveUIDocument.Document
sel = __revit__.ActiveUIDocument.Selection

import clr
clr.AddReference('RevitAPI')

#Проверка параметра - системный или нет
def parameter_check(parameter):
	types = ['системный',
		'несистемный']
	if isinstance(parameter, Parameter):
		b_parameter = parameter.Definition.BuiltInParameter
		check = b_parameter == BuiltInParameter.INVALID
		print 'Параметр с Id {} - {}'.format(parameter.Id, types[check])
		print 'Значение его перечисления равно {}\n'.format(b_parameter)		
        else:
            print 'Некорректный аргумент\n' \
              'Необходим аргумент класса Parameter\n'

#Чудо-принтер:
def bprint(array='', depth=0):
    start_symbol = '\t' * depth
    if isinstance(array, dict):
        for key, value in array.items():
            bprint(key, depth)
            bprint(value, depth + 1)
    elif hasattr(array, '__iter__') and not isinstance(array, str):
        for element in array:
            bprint(element, depth)
    else:
        print '{}{}'.format(start_symbol, array)
       
#element_ids = sel.GetElementIds()
#ids = [elem.IntegerValue for elem in element_ids]
#bprint(ids)
#print(type(ids))
#print(type(element_ids))
#print(type(sel))

#wall_types = FEC(doc).OfClass(WallType)

#marks_correct = []
#marks_uncorrect = ''
#problem_parameter = ''

#for wall in wall_types:
#    parameter = wall.Parameter[BuiltInParameter.ALL_MODEL_TYPE_MARK]
#    if isinstance(parameter, Parameter):
#        marks_correct.append(parameter)
#    else:
#        marks_uncorrect += wall.Category.Name
#        problem_parameter = parameter

#sum_ids = sum([elem.Id.IntegerValue for elem in marks_correct])
#print '{}, {}, {}'.format(sum_ids, marks_uncorrect, problem_parameter)

#constrs_type = FEC(doc).OfClass(ElementType)

#params_Ids = {-1010103, -1010109, -1010108}

#for elem in constrs_type:
#    parameter1 = elem.get_Parameter()







wall_types = FEC(doc).OfClass(WallType)
wall_params = [wall.Parameter[BuiltInParameter.ALL_MODEL_TYPE_MARK] for wall in wall_types]

marks_correct = []
marks_uncorrect = []

for wall in wall_params:
    #parameter = wall.Parameter[BuiltInParameter.ALL_MODEL_TYPE_MARK]
    if isinstance(wall, Parameter):
        marks_correct.append(wall)
    else:
        marks_uncorrect.append(wall)

sum_ids = sum([elem.Id.IntegerValue for elem in marks_correct])
print([mark.Category.Name for mark in marks_uncorrect])
print(sum_ids)









#filter = ElementMulticategoryFilter(
#    List[BuiltInCategory]([
#        BuiltInCategory.OST_Doors,
#        BuiltInCategory.OST_Windows,
#        BuiltInCategory.OST_GenericModel]),
#    True)
#elements = FEC(doc).WhereElementIsViewIndependent() \
#    .OfClass(FamilyInstance) \
#    .WherePasses(filter) \
#    .ToElements()
#elementIds = [element.UniqueId for element in elements] #Генератор списков
#elements_new = [doc.GetElement(elementId) for elementId in elementIds]
#elements_category_Id = [element.Category.Id for element in elements_new]

#bprint(clr.References)
#bprint(elements_category_Id)
#print("Число элементов: " + str(len(elements)))
#print "Минимальное значение Id категории: " + str(min(elements_category_Id))

#print ElementId(BuiltInCategory.OST_Walls)


constrs_type = FEC(doc).OfClass(ElementType)

params_Ids = [-1010103, -1010109, -1010108]
b_parameters = BuiltInParameter.GetValues(BuiltInParameter)
params = []
params_collector = []
params_collector_invalid = []

for param in b_parameters:
    if int(param) in params_Ids:
        params.append(param)

for elem in constrs_type:
    for param in params:
        parameter = elem.Parameter[param]
        if isinstance(parameter, Parameter):
            params_collector.append(parameter)
        else:
            params_collector_invalid.append(parameter)

print("Корректные параметры: " + str(len(params_collector)))
print("Некорректные параметры: " + str(len(params_collector_invalid)))




import clr
clr.AddReference('RevitAPI') #not necessary
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from Autodesk.Revit import DB

doc = __revit__.ActiveUIDocument.Document

#Чудо-принтер:
def bprint(array='', depth=0):
    start_symbol = '\t' * depth
    if isinstance(array, dict):
        for key, value in array.items():
            bprint(key, depth)
            bprint(value, depth + 1)
    elif hasattr(array, '__iter__') and not isinstance(array, str):
        for element in array:
            bprint(element, depth)
    else:
        print '{}{}'.format(start_symbol, array)

#Получение выделенных элементов из модели:
def get_selected_elements(doc):
    """API change in Revit 2016 makes old method throw an error"""
    try:
        # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)
       
#Проверка StorageType параметра:
def get_parameter_value_v1(parameter):
    if isinstance(parameter, DB.Parameter):
        storage_type = parameter.StorageType
        if storage_type == DB.StorageType.Integer:
            return parameter.AsInteger()
        elif storage_type == DB.StorageType.Double:
            return parameter.AsDouble()
        elif storage_type == DB.StorageType.String:
            return parameter.AsString()
        elif storage_type == DB.StorageType.ElementId:
            return parameter.AsElementId()

def get_parameter_value_v2(parameter):
    if isinstance(parameter, DB.Parameter):
        storage_type = parameter.StorageType
        if storage_type:
            exec 'parameter_value = parameter.As{}()'.format(storage_type)
            return parameter_value

selection = get_selected_elements(doc)
element = selection[0]
parameters = element.Parameters

for parameter in parameters:
    name = parameter.Definition.Name
    value = get_parameter_value_v2(parameter)
    print '{} = {}'.format(name, value)




#Получение выделенных элементов из модели:
def get_selected_elements(doc):
    """API change in Revit 2016 makes old method throw an error"""
    try:
        # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)
       
#Проверка StorageType параметра:
def get_parameter_value_v1(parameter):
    if isinstance(parameter, DB.Parameter):
        storage_type = parameter.StorageType
        if storage_type == DB.StorageType.Integer:
            return parameter.AsInteger()
        elif storage_type == DB.StorageType.Double:
            return parameter.AsDouble()
        elif storage_type == DB.StorageType.String:
            return parameter.AsString()
        elif storage_type == DB.StorageType.ElementId:
            return parameter.AsElementId()

def get_parameter_value_v2(parameter):
    if isinstance(parameter, DB.Parameter):
        storage_type = parameter.StorageType
        if storage_type:
            exec 'parameter_value = parameter.As{}()'.format(storage_type)
            return parameter_value

selection = get_selected_elements(doc)

mc_filter = DB.ElementMulticategoryFilter(
    List[DB.BuiltInCategory]([
        DB.BuiltInCategory.OST_Walls,
        DB.BuiltInCategory.OST_GenericModel
    ]))

element = selection[0]
#print(type(element))
parameters = [element.GetDependentElements(mc_filter)]
#print(type(parameters))
#bprint([(param, doc.GetElement(param).Category.Name) for param in parameters])
for param in parameters:
    print(str(param, '/n'))