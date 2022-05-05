# -*- coding: utf-8 -*-

__title__ = "СИТИ_Анализ изоляции ВК/ОВ на активном виде"
__author__ = 'Tima Kutsko'
__doc__ = "Заполнение параметра СИТИ_Секция, СИТИ_Система/СИТИ_Этаж и СИТИ_Код по классификатору в изоляцию воздуховодов трубопроводов\n"\
          "Актуально только для ВК, ОВ!"



from Autodesk.Revit.DB import *
from rpw import revit, db, ui
from pyrevit import script
from System import Guid


# definitions

def insulation(elements_insulation):
    errorEl = False
    for current_insulation in elements_insulation:
        for current_guid in guids_list:
            current_element = doc.GetElement(current_insulation.HostElementId)
            try:
                element_current_guid_param = current_element.get_Parameter(current_guid).AsString()
            except:
                errorEl = True
                continue
            if element_current_guid_param is None:
                element_current_guid_param = doc.GetElement(current_element.GetTypeId()).get_Parameter(current_guid).AsString()
            try:
                if not element_current_guid_param is None:
                    try:
                        current_insulation.get_Parameter(current_guid).Set(str(element_current_guid_param))
                    except:
                        doc.GetElement(current_insulation.GetTypeId()).get_Parameter(current_guid).Set(str(element_current_guid_param))
            except:
                output.print_md("**Элемент, у которого нет нужных праметров:** {} - {} \
                                    **Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".format(current_insulation.Name, output.linkify(current_insulation.Id)))
                script.exit()
        if errorEl:
            output.print_md("**Проблемный элемент:** {} \
                                    **Элемент будет пропущен, необходимо проверить вручную!**".format(output.linkify(current_insulation.Id)))
            errorEl = False


""" def insulation(elements_insulation):   #for lookup 
    for current_insulation in elements_insulation:
        for current_guid in guids_list:
            current_element = doc.GetElement(current_insulation.HostElementId)
            element_current_guid_param = current_element.LookupParameter(current_guid).AsString()
            if element_current_guid_param is None:
                try:	
                    element_current_guid_param = doc.GetElement(current_element.GetTypeId()).LookupParameter(current_guid).AsString() 
                except:
                    pass   #you can get elements with out data in try parameter
            try:
                if not element_current_guid_param is None:				
                    try:
                        current_insulation.LookupParameter(current_guid).Set(str(element_current_guid_param)) 
                    except:
                        doc.GetElement(current_insulation.GetTypeId()).LookupParameter(current_guid).Set(str(element_current_guid_param))
            except:
                output.print_md("**Элемент, у которого нет нужных праметров:** {} - {} \
                                    **Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".format(current_insulation.Name, output.linkify(current_insulation.Id)))
                script.exit() """


# main code
doc = revit.doc
output = script.get_output()
# СИТИ_Система
guid_system_param = Guid("98b30303-a930-449c-b6c2-11604a0479cb")
# СИТИ_Этаж
guid_level_param = Guid("79c6dee3-bd92-4fdf-ae12-c06683c61946")
# СИТИ_Секция
guid_section_param = Guid("4e8c4436-b231-4f9c-a659-181a2b55eda4")
# СИТИ_Комплект_документации
guid_complect_param = Guid("1dfc5c47-4308-4443-8229-15d1392e3f57")
guids_list = [guid_section_param, guid_system_param,
              guid_level_param, guid_complect_param]


""" guids_list = ["КП_ПП_Принадлежность к корпусу"]  #for lookup """


## insulations
mep_elements_insulation = list()
pipe_insulation = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements()
duct_insulation = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
mep_elements_insulation.extend(pipe_insulation)
mep_elements_insulation.extend(duct_insulation)

with db.Transaction("СИТИ_Анализ изоляции ОВ/ВК"):
    if len(mep_elements_insulation) > 0:
        insulation(mep_elements_insulation)
        ui.forms.Alert('Выполнено успешно!', title='Анализ изоляции ОВ/ВК')
    else:
        print("В проекте нет изоляции труб/воздуховодов. Только для ОВ/ВК!")