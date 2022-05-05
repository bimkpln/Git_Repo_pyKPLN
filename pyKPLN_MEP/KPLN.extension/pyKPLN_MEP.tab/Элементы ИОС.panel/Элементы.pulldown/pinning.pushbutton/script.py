#coding: utf-8

__title__ = "Прикрепление эл-тов"
__author__ = 'Tima Kutsko'
__doc__ = "Прикрепление осей, сеток и связанных моделей"


from Autodesk.Revit.DB import *
from rpw import revit, DB, db, ui
from pyrevit import script, forms


#variables
doc = revit.doc
output = script.get_output()
checkBoxElements = []
cnt = 0


#classes
class ElementOption:
    def __init__(self, name, value):
        self.name = name
        if self.name == "Оси":
            self.value = BuiltInCategory.OST_Grids
        elif self.name == "Уровни":
            self.value = BuiltInCategory.OST_Levels
        elif self.name == "Базовая точка проекта":
            self.value = BuiltInCategory.OST_ProjectBasePoint
        elif self.name == "Точка съёмки":
            self.value = BuiltInCategory.OST_SharedBasePoint
        else:
            self.value = value


#functions
def create_check_boxes(elements, grids, levels, point_1, point_2): 

    elements_options = [ElementOption(e.Name.split(':')[0], e) for e in sorted(elements, key=lambda x: x.Name)]
    elements_options.extend([ElementOption(grid, grid) for grid in grids])
    elements_options.extend([ElementOption(level, level) for level in levels])
    elements_options.extend([ElementOption(point, point) for point in point_1])
    elements_options.extend([ElementOption(point, point) for point in point_2])
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect = True,
                                                    title='Выбери элементы для прикрепления',
                                                    width=300,
                                                    button_name='Выбрать')
    return elements_checkboxes


#main code
linkCollector = FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).WhereElementIsNotElementType().ToElements()
gridCollector = ["Оси"]
levelCollector = ["Уровни"]
basePoint = ["Базовая точка проекта"]
surveyPoint = ["Точка съёмки"]
check_box = create_check_boxes(linkCollector, gridCollector, levelCollector, basePoint, surveyPoint)


try:
    #main part of code
    for item in check_box:
        try:
            checkBoxElements.extend(FilteredElementCollector(doc).OfCategory(item.value).WhereElementIsNotElementType().ToElements())
        except:
            checkBoxElements.append(item.value)
    with db.Transaction('pyKPLN_Прикрепление основных эл-в'):
        for element in checkBoxElements:
            if not element.Pinned:
                pinElement = element.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(1)
                cnt += 1
        output.print_md("Прикреплено **{}** элемента(-ов)".format(cnt))
except:
    script.exit()
    
