# -*- coding: utf-8 -*-

__title__ = "ДИВ_Проверка перед выгрузкой (на активном виде!)"
__author__ = 'Tima Kutsko'
__doc__ = "Проверка модели на наличее пустых параметров (на активном виде!)\n"\
      "Актуально для всех разделов ИОС!"


from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector
from rpw import revit, ui
from pyrevit import script
from System import Guid, Enum


# definitions
# getting builincategory from element's category
def get_BuiltInCategory(category):
    if Enum.IsDefined(BuiltInCategory, category.Id.IntegerValue):
        BuiltInCat_id = category.Id.IntegerValue
        BuiltInCat = Enum.GetName(BuiltInCategory, BuiltInCat_id)
        return BuiltInCat
    else:
        return BuiltInCategory.INVALID


# main code
doc = revit.doc
output = script.get_output()
views_list = list()
flag = False

# check model elements
# ДИВ_Секция_Текст
guid_section_param = Guid("a3513e5d-b1f8-4f60-ae09-5c0fb579458a")
# ДИВ_RBS_Код по классификатору
guid_RBS_param = Guid("25377442-92fe-4eaa-a4fd-349302493a2c")
# ДИВ_Система_Текст
guid_system_param = Guid("1b2a7570-5cf6-4db9-9490-d8a147466c53")
# ДИВ_Этаж_Текст
guid_level_param = Guid("830e785e-ece4-499d-943f-ffb314fe5031")
# ДИВ_Имя_классификатора
guid_name_param = Guid("0f3cab6f-71c9-4c1e-b92b-c3d2838f4775")
# ДИВ_Наименование по классификатору
guid_type_param = Guid("ff594263-a975-400c-8915-343aef95ceb3")
# ДИВ_Комплект_документации
guid_complect_param = Guid("c03f8a1d-47bd-471b-92c5-9d4c1e862202")
guids_list_strong = [guid_RBS_param]
guids_list_var = [guid_section_param,
                  guid_system_param,
                  guid_level_param]


# get all elements of current project
categoryList = ["OST_DuctCurves", "OST_PipeCurves",
                "OST_FlexPipeCurves", "OST_FlexDuctCurves",
                "OST_PipeInsulations", "OST_DuctInsulations",
                "OST_DuctAccessory", "OST_LightingDevices",
                "OST_DuctFitting", "OST_DuctTerminal",
                "OST_PipeAccessory", "OST_PipeFitting",
                "OST_CableTray", "OST_Conduit",
                "OST_MechanicalEquipment", "OST_ElectricalEquipment",
                "OST_CableTrayFitting", "OST_ConduitFitting",
                "OST_LightingFixtures", "OST_ElectricalFixtures",
                "OST_DataDevices", "OST_FireAlarmDevices",
                "OST_SecurityDevices", "OST_NurseCallDevices",
                "OST_CommunicationDevices", "OST_PlumbingFixtures"]
all_elements_in_doc = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
for current_element in all_elements_in_doc:
    if current_element.Category:
        if get_BuiltInCategory(current_element.Category) in categoryList:
            try:
                super_component = current_element.SuperComponent
            except:
                super_component = None
            if super_component is None:
                try:
                    for current_guid in guids_list_strong:
                        element_current_guid_param = current_element.get_Parameter(current_guid).AsString()
                        if element_current_guid_param is None:
                            element_current_guid_param = doc.GetElement(current_element.GetTypeId()).get_Parameter(current_guid).AsString()
                        if not element_current_guid_param:
                            flag = True
                            output.print_md("**Элемент с пустым параметром:** {} - {}".format(current_element.Name, output.linkify(current_element.Id)))
                    var_flag = 0
                    for current_guid in guids_list_var:
                        element_current_guid_param = current_element.get_Parameter(current_guid).AsString()
                        if element_current_guid_param is None:
                            element_current_guid_param = doc.GetElement(current_element.GetTypeId()).get_Parameter(current_guid).AsString()
                        if not element_current_guid_param:
                            var_flag += 1
                    if var_flag > 1:
                        flag = True
                        output.print_md("**Элемент с пустым параметром:** {} - {}".format(current_element.Name, output.linkify(current_element.Id)))
                except:
                    flag = True
                    output.print_md("**Элемент, у которого нет нужных праметров:** {} - {} \
                            **Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".format(current_element.Name, output.linkify(current_element.Id)))
                    script.exit()

# output form
if not flag:
    ui.forms.Alert('Все элементы готовы к выгрузке!',
                   title='Проверка элементов на активном виде')
else:
    print("_____________________end______________________")