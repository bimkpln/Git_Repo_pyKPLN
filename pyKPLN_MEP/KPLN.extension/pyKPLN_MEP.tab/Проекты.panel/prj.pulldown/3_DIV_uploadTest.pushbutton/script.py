# -*- coding: utf-8 -*-

__title__ = "СИТИ_Проверка перед выгрузкой (на активном виде!)"
__author__ = 'Tima Kutsko'
__doc__ = "Проверка модели на наличее пустых параметров (на активном виде!)\n"\
      "Актуально для всех разделов ИОС!"


from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector,BuiltInParameter
from rpw import revit, ui, db
from pyrevit import script, forms
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
# СИТИ_Диаметр
guid_diameter_param = Guid("d03cddce-6d51-442c-80bd-5ee41ac3e55c")
# СИТИ_Размер
guid_size_param = Guid("a1e634ab-3018-4e25-8ebf-0c65af534fbe")
# СИТИ_Секция
guid_section_param = Guid("4e8c4436-b231-4f9c-a659-181a2b55eda4")
# СИТИ_Система
guid_system_param = Guid("98b30303-a930-449c-b6c2-11604a0479cb")
# СИТИ_Этаж
guid_level_param = Guid("79c6dee3-bd92-4fdf-ae12-c06683c61946")
# СИТИ_Классификатор
guid_RBS_param = Guid("32bf9389-a9b8-4db4-8d67-2fce3844b607")
# СИТИ_Имя_классификатора
guid_name_param = Guid("5ec9640c-c857-4c3b-9119-c6fd79f820e2")
# СИТИ_Описание
guid_type_param = Guid("ffb1aec5-f095-4324-b034-c469c874fd3a")
# СИТИ_Комплект_документации
guid_complect_param = Guid("1dfc5c47-4308-4443-8229-15d1392e3f57")
# СИТИ_Марка/артикул
guid_code_param = Guid("0a77b421-e8a0-4b5c-86a7-10aaf75b04ba")
# СИТИ_Завод_изготовитель
guid_plant_param = Guid("985879f9-2fce-473a-9a70-b26b6316e807")
guids_list_strong = [guid_RBS_param,
                     guid_name_param,
                     guid_type_param,
                     guid_complect_param,
                     guid_code_param,
                     guid_plant_param]
guids_list_var = [guid_system_param,
                  guid_level_param]
guids_list_sec = [guid_section_param]                 
guids_list_linear = [guid_diameter_param, guid_size_param]
BuiltInCategoryList = [BuiltInCategory.OST_DuctCurves,BuiltInCategory.OST_PipeCurves,BuiltInCategory.OST_FlexPipeCurves,BuiltInCategory.OST_FlexDuctCurves,BuiltInCategory.OST_PipeInsulations,BuiltInCategory.OST_DuctInsulations,BuiltInCategory.OST_DuctAccessory,BuiltInCategory.OST_LightingDevices,
BuiltInCategory.OST_CableTray,BuiltInCategory.OST_Conduit,BuiltInCategory.OST_MechanicalEquipment,BuiltInCategory.OST_ElectricalEquipment,BuiltInCategory.OST_CableTrayFitting,BuiltInCategory.OST_ConduitFitting,BuiltInCategory.OST_LightingFixtures,BuiltInCategory.OST_ElectricalFixtures,BuiltInCategory.OST_DataDevices,BuiltInCategory.OST_FireAlarmDevices,
BuiltInCategory.OST_SecurityDevices,BuiltInCategory.OST_NurseCallDevices,BuiltInCategory.OST_CommunicationDevices,BuiltInCategory.OST_PlumbingFixtures,BuiltInCategory.OST_PipeAccessory,BuiltInCategory.OST_PipeFitting,BuiltInCategory.OST_GenericModel]

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
all_elements_in_doc = FilteredElementCollector(doc, doc.ActiveView.Id).\
                      WhereElementIsNotElementType().ToElements()
#1___________________________ПРОВЕРКА ЭЛЕМЕНТОВ С КОДОМ 9999 __________________________________________________________
count = 0
with forms.ProgressBar(title='Проверка элементов с кодом "9999"...',cancellable=True) as pb:
    with db.Transaction("Заполнение кода 9999"):
        Kod9999Dict = dict()
        for cat in BuiltInCategoryList:
            for i in FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements():
                partition = i.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString()
                try:
                    famName = i.Symbol.Family.Name
                except:
                    famName = i.Name
                try:
                    if i.Category.Name == 'Соединительные детали кабельных лотков':
                        if '01_' in partition[:5] or 'болт' in famName.lower() or 'гайка' in famName.lower():
                            if i.Name not in Kod9999Dict:
                                Kod9999Dict[i.Name] = i
                                try:
                                    i.Symbol.get_Parameter(guid_RBS_param).Set('9999')
                                except:
                                    doc.GetElement(i.Id).get_Parameter(guid_RBS_param).Set('9999')
                            else:
                                continue
                        if 'СЛУ' in famName and i.SuperComponent != None:
                            if i.Name not in Kod9999Dict:
                                Kod9999Dict[i.Name] = i
                                try:
                                    i.Symbol.get_Parameter(guid_RBS_param).Set('9999')
                                except:
                                    doc.GetElement(i.Id).get_Parameter(guid_RBS_param).Set('9999')
                            else:
                                continue
                    if '001_' in famName.lower()[:5] or '_ниша' in famName.lower() or '501_' in famName.lower()[:5] or i.Category.Name == 'Обобщенные модели' or 'крепление' in famName.lower() or 'вводной кабель' in i.Name.lower():
                        if i.Name not in Kod9999Dict:
                            Kod9999Dict[i.Name] = i
                            try:
                                i.Symbol.get_Parameter(guid_RBS_param).Set('9999')
                            except:
                                doc.GetElement(i.Id).get_Parameter(guid_RBS_param).Set('9999')
                        else:
                            continue
                except:
                    continue
            count += 1
            pb.update_progress(count, len(BuiltInCategoryList))
#1__________________________________________________________________________________________________________
for current_element in all_elements_in_doc:
        try:
            noneCodeParamData = current_element.\
                                get_Parameter(guid_RBS_param).\
                                AsString()
            if noneCodeParamData is None:
                noneCodeParamData = doc.\
                                    GetElement(current_element.GetTypeId()).\
                                    get_Parameter(guid_RBS_param).\
                                    AsString()
        except Exception:
            continue
        if noneCodeParamData == "9999":
            continue
        current_elem_cat = get_BuiltInCategory(current_element.Category)
        if current_elem_cat in categoryList:
            # try:
            #     super_component = current_element.SuperComponent
            # except:
            #     super_component = None
            # if super_component is None:
            try:
                var_flag = 0
                print_flag = 0
                print_flag_error = 0
                strong_flag = 0
                var_flag_sec = 0
                duct_flag = 0

                for current_guid in guids_list_strong:
                    element_current_guid_param = current_element.\
                                                    get_Parameter(current_guid).\
                                                    AsString()
                    if element_current_guid_param is None:
                        element_current_guid_param = doc.\
                                                        GetElement(current_element.GetTypeId()).\
                                                        get_Parameter(current_guid).\
                                                        AsString()
                    if not element_current_guid_param:
                        strong_flag += 1
                    elif element_current_guid_param.startswith("Е")\
                            or element_current_guid_param.startswith("В")\
                            and len(element_current_guid_param.split("."))>4:
                        flag = True
                        output.print_md("Запрещено использовать **кириллицу**.\
                                        Ошибка в элементе {}\
                                        **Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".
                                        format(output.linkify(current_element.Id)))
                        script.exit()

                for current_guid in guids_list_sec:
                    element_current_guid_param = current_element.\
                                                    get_Parameter(current_guid).\
                                                    AsString()
                    if element_current_guid_param is None:
                        element_current_guid_param = doc.\
                                                        GetElement(current_element.GetTypeId()).\
                                                        get_Parameter(current_guid).\
                                                        AsString()
                    if not element_current_guid_param:
                        var_flag_sec += 1

                for current_guid in guids_list_var:
                    element_current_guid_param = current_element.\
                                                    get_Parameter(current_guid).\
                                                    AsString()
                    if element_current_guid_param is None:
                        element_current_guid_param = doc.\
                                                        GetElement(current_element.GetTypeId()).\
                                                        get_Parameter(current_guid).\
                                                        AsString()
                    if not element_current_guid_param:
                        var_flag += 1

                if str(current_elem_cat) == "OST_PipeCurves" or\
                        str(current_elem_cat) == "OST_Conduit":
                    element_current_guid_param = current_element.\
                                                get_Parameter(guid_diameter_param).\
                                                AsString()
                    if element_current_guid_param is None:
                        element_current_guid_param = doc.\
                                                        GetElement(current_element.GetTypeId()).\
                                                        get_Parameter(guid_diameter_param).\
                                                        AsString()
                    if not element_current_guid_param:
                        strong_flag += 1

                for current_guid in guids_list_linear:
                    if str(current_elem_cat) == "OST_DuctCurves":
                        element_current_guid_param = current_element.\
                                                        get_Parameter(current_guid).\
                                                        AsString()
                        if element_current_guid_param is None:
                            element_current_guid_param = doc.\
                                                        GetElement(current_element.GetTypeId()).\
                                                        get_Parameter(current_guid).\
                                                        AsString()
                        if not element_current_guid_param:
                            duct_flag += 1        

                if str(current_elem_cat) == "OST_CableTray":
                    element_current_guid_param = current_element.\
                                                get_Parameter(guid_size_param).\
                                                AsString()
                    if element_current_guid_param is None:
                        element_current_guid_param = doc.\
                                                        GetElement(current_element.GetTypeId()).\
                                                        get_Parameter(guid_size_param).\
                                                        AsString()
                    if not element_current_guid_param:
                        strong_flag += 1

                if duct_flag > 1:
                    print_flag += 1

                if strong_flag > 0:
                    print_flag += 1

                if var_flag_sec > 0:
                    print_flag += 1

                if  var_flag > 1:
                    print_flag += 1

                if var_flag_sec is 0:
                    if  var_flag is 0:
                        print_flag_error += 1
                        print_flag = 0

                if print_flag > 0:
                    flag = True
                    output.print_md("**Элемент с пустым параметром:** {} - {}".format(current_element.Name, output.linkify(current_element.Id)))

                if print_flag_error > 0:
                    flag = True
                    output.print_md("**Элемент у которого заполнены параметры и системы и этажа:** {} - {}".format(current_element.Name, output.linkify(current_element.Id)))
            except:
                flag = True
                print(current_element.Id)
                output.print_md("**Элемент, у которого нет нужных праметров:** {} - {} \
                                **Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".
                                format(current_element.Name,
                                        output.linkify(current_element.Id)))
                script.exit()

# output form
if not flag:
    ui.forms.Alert('Все элементы готовы к выгрузке!',
                   title='Проверка элементов на активном виде')
else:
    print("_____________________end______________________")