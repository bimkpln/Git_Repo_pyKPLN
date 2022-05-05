#coding: utf-8

__title__ = "Высотн. отм.\nОТВЕРСТИЙ в стенах_v.0"
__author__ = 'Tima Kutsko'
__doc__ = "Запись отметок в параметр 'Комментарии' шахт/ниш/отверстий или проёмов для старых семейств (не работает с менеджером отверстий!)"


import Autodesk
from Autodesk.Revit.DB import *
from rpw import revit, DB, db, ui
from pyrevit import script
from System import Guid



#variables
doc = revit.doc
output = script.get_output()



#main code
ui.forms.Alert('Выбери по одному экземпляру шахты/ниши/отверстия и проёма для расчета и присвоения им отметки. Параметр для маркировки отметок - "Комментарии"', 
                title='Отметки шахт/ниш/отверстий и проёмов', header='ВНИМАНИЕ!!!')
selected_references = ui.Pick.pick_element(msg='Выбери экземпляры для анализа', multiple=True)
family_name_set = set()
family_category_id_set = set()
element_collector = list()

for reference in selected_references:	
    family_instance = doc.GetElement(reference.id)	
    family_name = family_instance.Symbol.FamilyName
    family_name_set.add(family_name)
    family_category_id = family_instance.Category.Id	
    family_category_id_set.add(family_category_id)



with db.Transaction("pyKPLN_Маркировка шахт/проёмов"):
    for category_id in family_category_id_set:
        element_collector.extend(FilteredElementCollector(doc).OfCategoryId(category_id).WhereElementIsNotElementType().ToElements())
        for element in element_collector:
            if element.Symbol.FamilyName in family_name_set:
                try:
                    diam = element.get_Parameter(Guid("9b679ab7-ea2e-49ce-90ab-0549d5aa36ff")).AsDouble()
                    middle_Z = element.Location.Point.Z * 304.8
                    if middle_Z < 0:
                        comment = "Отм. оси: " + str('{:.1f}'.format(middle_Z)) # The format of output is without a dot
                    else:
                        comment = "Отм. оси: " + "+" + str('{:.1f}'.format(middle_Z)) # The format of output is without a dot
                except:
                    min_Z = element.get_Geometry(Options()).GetBoundingBox().Min.Z * 304.8
                    if min_Z < 0:
                        comment = "Отм. низа: " + str('{:.1f}'.format(min_Z)) # The format of output is without a dot
                    else:
                        comment = "Отм. низа: " + "+" + str('{:.1f}'.format(min_Z)) # The format of output is without a dot
                
                #seting data to element
                element.LookupParameter("Комментарии").Set(comment)