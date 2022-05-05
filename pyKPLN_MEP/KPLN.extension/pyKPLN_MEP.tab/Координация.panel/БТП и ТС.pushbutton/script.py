#coding: utf-8

__title__ = "Проверка координат"
__author__ = 'Tima Kutsko'
__doc__ = "Отслеживание истинных и относительных координат базовой точки проекта и точки съёмки"



import Autodesk
from Autodesk.Revit.DB import *
from rpw import revit, ui, db
from rpw.ui.forms import Console, CommandLink, TaskDialog
from pyrevit import script, DB


#variables
doc = revit.doc
output = script.get_output()
table_params = [] 


#definitions
def project_base_point(lst, document):
    view = doc.ActiveView
    for point in lst:   
        bound_box_max_point = point.get_BoundingBox(view).Max * 304.8
        north_south_point = point.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble() * 304.8 
        east_west_point = point.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble() * 304.8    
        elevation_point = point.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble() * 304.8 
        angleton_point = point.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsDouble() * 57.2958
        return bound_box_max_point, north_south_point, east_west_point, elevation_point, angleton_point, who_created_element(lst[0], document)


def project_survey_point(lst):
    for point in lst:
        bound_box_max_point = '-'        
        north_south_point = point.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble() * 304.8 
        east_west_point = point.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble() * 304.8    
        elevation_point = point.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble() * 304.8 
        angleton_point = '-'
        who_created_PSP = '-'       
        return bound_box_max_point, north_south_point, east_west_point, elevation_point, angleton_point, who_created_PSP


def who_created_element(element, document):     
    try:        
        if document.IsWorkshared:        
            wti = DB.WorksharingUtils.GetWorksharingTooltipInfo(document, element.Id)
            lastChanger = wti.LastChangedBy 
        else:
            lastChanger = 'Нет рабочих наборов!'
    except:
        lastChanger = 'Выгружен!'        
    return lastChanger



#forms
commands= [CommandLink('Проверить координаты в данном проекте?', return_value='isCheck'),
            CommandLink('Проверить координаты в связанных проектах?', return_value='isLinkCheck')]
dialog = TaskDialog('Выбери формат теста',																
                    commands=commands,
                    buttons=['Cancel'],
                    footer='Контроль уровня элементов ИОС',
                    show_close=True)
input_result = dialog.show()


#main code
if input_result == 'isLinkCheck':
    linkModels = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.RevitLinkInstance)
    for link in linkModels:                
        try:
            linkDoc = link.GetLinkDocument()            
            project_base_point_list =FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
            project_survey_point_list = FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()
            table_params.append(["БТП в " + link.Name.split(':')[0][0:-5]] + list(project_base_point(project_base_point_list, linkDoc)))
            table_params.append(["ТС в " + link.Name.split(':')[0][0:-5]] + list(project_survey_point(project_survey_point_list)))
        except:
            pass 
    output.print_table(table_data=table_params, title="Параметры БТП и ТС", columns=["Имя файла", "Истинные координаты БТП", "С/Ю", "В/З", "Отметка", "Угол от истинного севера", "Последний редактор"])         

if input_result == 'isCheck':
    project_base_point_list =FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
    project_survey_point_list = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()
    table_params.append(["БТП в " + doc.Title.split(':')[0][0:-4]] + list(project_base_point(project_base_point_list, doc)))
    table_params.append(["ТС в " + doc.Title.split(':')[0][0:-4]] + list(project_survey_point(project_survey_point_list)))
    output.print_table(table_data=table_params, title="Параметры БТП и ТС", columns=["Имя файла", "Истинные координаты БТП", "С/Ю", "В/З", "Отметка", "Угол от истинного севера", "Последний редактор"])  