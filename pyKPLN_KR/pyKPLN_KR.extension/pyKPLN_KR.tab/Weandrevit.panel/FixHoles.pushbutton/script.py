# -*- coding: utf-8 -*-
"""
Button

"""
__author__ = 'Khlapov Dmitry'
__title__ = "Починить\nотверстия"
__doc__ = 'Производит починку отверстий, которые начали прорезать арматуру по площади' \

"""
Архитекурное бюро KPLN

"""
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from Autodesk.Revit.DB import *
from rpw import revit, db, ui
from pyrevit import script, forms

doc = __revit__.ActiveUIDocument.Document

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

        panels = FEC(doc) \
    .OfCategory(DB.BuiltInCategory.OST_CurtainWallPanels) \
    .WhereElementIsNotElementType()

selection = get_selected_elements(doc)

if len(selection) == 0:
	ui.forms.Alert("Выберите стену!", title="Внимание", exit=True)	
if len(selection) > 1:
	ui.forms.Alert("Выберите одну стену!", title="Внимание", exit=True)	

wall = selection[0]
if not isinstance(wall, Wall):
	ui.forms.Alert("Выберите стену, была выбрана не стена!", title="Внимание", exit=True)	

holes = FEC(doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(FamilyInstance).ToElements()
holes_in_wall = []
  
for hole in holes:
    parameter_symbol = hole.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
    if '231' in parameter_symbol.Element.Name:
        if JoinGeometryUtils.AreElementsJoined(doc, wall, hole):
            holes_in_wall.append(hole)
counter = 0

with Transaction(doc, "FixHoles") as t:
    t.Start()
    for hole in holes_in_wall:
        JoinGeometryUtils.UnjoinGeometry(doc, wall, hole)
        counter += 1

    doc.Regenerate()

    parameter_offset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    offset_value = parameter_offset.AsDouble()
    offset_value_new = offset_value - 0.01
    parameter_offset.Set(offset_value_new)

    doc.Regenerate()

    parameter_offset.Set(offset_value)

    doc.Regenerate()

    for hole in holes_in_wall:
        JoinGeometryUtils.JoinGeometry(doc, hole, wall)
    t.Commit()
    if t.GetStatus() == TransactionStatus.Committed:
        ui.forms.Alert("Отработано отверстий: {}".format(counter), title="Успешно", exit=False)	
    else: 
	    ui.forms.Alert("Во время выполнения произошла ошибка!", title="Внимание", exit=True)	