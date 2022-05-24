# coding: utf-8

__title__ = "ОВ2. Заполнение параметров для спецификации по нескольким категориям v1.1"
__author__ = 'Kapustin Roman'
__doc__ = "Заполняет:\n 1.Для воздуховодв длину и габариты, считая воздуховоды MxN и NxM как MxN \n 2.Для элементов в шт. число. \n 3.Для изоляции площадь и толщину стенки.\n 4. Единицы измерения \n 5. Запас для линейных элементов \n 6. Порядок по категориям для сортировки"


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
                              BuiltInParameter
import System
from rpw import revit, db
from pyrevit import script
import codecs
from System import Guid
from rpw.ui.forms import*


# main code: input part
doc = revit.doc
output = script.get_output()

commands = [CommandLink('Электролитный', return_value='Электролитный'),
            CommandLink('Новые проекты', return_value='Новые проекты')]
dialog = TaskDialog('Выберите тип файла:',
                    commands=commands,
                    show_close=True)
dialog_out = dialog.show()

if dialog_out == 'Новые проекты':
    # КП_И_Запас
    supply = Guid("62ff7f1d-f5cd-4e5e-9f50-b7b3e18f2c06")
    # КП_И_КолСпецификация
    guid_length_param = Guid("df684fd7-6956-4d85-8475-709481838383")
    # КП_Размер_Текст
    guid_size_param = Guid("7d38dcf5-47ef-46d0-b4a8-9babee9cc11b")
    # КП_И_Толщина стенки
    guid_fitSize_param = Guid("381b467b-3518-42bb-b183-35169c9bdfb3")
    # КП_О_Единица измерения
    guid_unit_param = Guid("4289cb19-9517-45de-9c02-5a74ebf5c86d")
    # КП_И_Порядковый номер
    guid_num_param = Guid("7a44592a-c698-4276-af5f-75be9b131b40")
    # КП_О_Сортировка
    guid_sort_param = Guid("82fad8c7-4f46-4da8-a69a-035a34b5f015")
elif dialog_out == 'Электролитный':
    supply = "Запас"
    #КП_Длина погонная - параметр проекта
    guid_length_param = Guid("2aa6d2a0-e0a6-468e-8a04-e1d10daea4f4")
    # КП_Размер_Текст
    guid_size_param = Guid("7d38dcf5-47ef-46d0-b4a8-9babee9cc11b")
    # КП_И_Толщина стенки
    guid_fitSize_param = Guid("381b467b-3518-42bb-b183-35169c9bdfb3")
    # КП_О_Единица измерения
    guid_unit_param = Guid("4289cb19-9517-45de-9c02-5a74ebf5c86d")
    # КП_И_Порядковый номер
    guid_num_param = Guid("7a44592a-c698-4276-af5f-75be9b131b40")
    # КП_О_Сортировка
    guid_sort_param = Guid("82fad8c7-4f46-4da8-a69a-035a34b5f015")
else:
    script.exit()
catLits_range = [BuiltInCategory.OST_DuctCurves,
                 BuiltInCategory.OST_FlexDuctCurves]
catList_quantity = [BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_DuctFitting]
catList_insulations = [BuiltInCategory.OST_DuctLinings,
                       BuiltInCategory.OST_DuctInsulations]
with db.Transaction('Заполнение корректной длины для спецификации по нескольким категориям'):
    components = [Label('Выберите параметры для заполнения:'),
                  CheckBox('checkbox1', 'КП_И_Количество в спецификацию', default=True),
                  CheckBox('checkbox2', 'КП_Размер_Текст', default=True),
                  CheckBox('checkbox3', 'КП_И_Порядковый номер', default=True),
                  CheckBox('checkbox4', 'КП_И_Запас', default=True),
                  Label('Размер запаса:'),
                  TextBox('textbox1', Text="1.15"),
                  CheckBox('checkbox5', 'КП_И_Толщина стенки', default=True),
                  CheckBox('checkbox6', 'КП_О_Единица измерения', default=True),
                  CheckBox('checkbox7', 'КП_О_Сортирование', default=False),
                  Separator(),
                  Button('Далее')]
    form = FlexForm('ОВ_Спецификации', components)
    form.show()
    supply_value = float(form.values.get('textbox1'))
    if supply_value < 1 or supply_value >= 2:
        Alert('Ввден некорректный запас!',
              title='Ошибка',
              header='',
              exit=True)
    checkbox1 = form.values.get('checkbox1')
    checkbox2 = form.values.get('checkbox2')
    checkbox3 = form.values.get('checkbox3')
    checkbox4 = form.values.get('checkbox4')
    checkbox5 = form.values.get('checkbox5')
    checkbox6 = form.values.get('checkbox6')
    checkbox7 = form.values.get('checkbox7')
    if form.values != {}:
        # Перенос длины в категории в метрах
        # Воздуховоды, гибкие воздуховоды
        for cat in catLits_range:
            elList_range = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_range:
                if checkbox1:
                    elRange_range = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString()
                    Range = (float(elRange_range))/1000
                    if dialog_out == 'Электролитный':
                        el.get_Parameter(guid_length_param).Set(Range/304.8)
                    else:
                        el.get_Parameter(guid_length_param).Set(Range)

            #Запись в параметр КП_Размер_Текст размера воздуховодов
                if checkbox2:
                    elRange_size_duct = el.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
                    if "x" in elRange_size_duct:
                        size_duct_rectangle = elRange_size_duct.split("x")
                        duct_width = size_duct_rectangle[0]
                        duct_height = size_duct_rectangle[1]
                        if int(duct_width) > int(duct_height):
                            duct_width_correct = duct_width
                            duct_height_correct = duct_height
                        else:
                            duct_width_correct = duct_height
                            duct_height_correct = duct_width
                        size_duct_rectangle_correct = duct_width_correct + "x" + duct_height_correct
                        el.get_Parameter(guid_size_param).Set(size_duct_rectangle_correct)
                    else:
                        el.get_Parameter(guid_size_param).Set(elRange_size_duct)
                if checkbox4:
                    #Запись запаса для длины
                    try:
                        if type(supply) == System.String:
                            new_supply_duct = doc.\
                                GetElement(el.GetTypeId()).\
                                LookupParameter(supply).\
                                AsDouble()
                        else:
                            new_supply_duct = doc.\
                                GetElement(el.GetTypeId()).\
                                get_Parameter(supply).\
                                AsDouble()
                    except:
                        print(el.Id)
                    if new_supply_duct < 1:
                        new_supply_duct = supply_value
                        doc.GetElement(el.GetTypeId()).get_Parameter(supply).Set(new_supply_duct)
                #Запись единиц измерения
                if checkbox6:
                    unit_duct = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                    if not unit_duct:
                        try:
                            doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("м")
                        except:
                            try:
                                el.get_Parameter(guid_unit_param).Set("м")
                            except:
                                output.print_md('У элемента **{} с id {}** не заполнен ед. изм!'.format(el.Name,output.linkify(el.Id)))
                #Запись номера для спецификации
                if checkbox3:
                    num_duct = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if num_duct < 1:
                        new_num_duct = 4
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_duct)
        #Задает 1 для категорий:
            #Оборудование, воздухораспределители, арматура воздуховодов, соеденительные детали воздуховодов.
        num = 1
        for cat in catList_quantity:
            elList_quantity = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_quantity:
                if checkbox1:
                    if dialog_out == 'Электролитный':
                        el.get_Parameter(guid_length_param).Set(1/304.8)
                    else:
                        el.get_Parameter(guid_length_param).Set(1)
                    
                if checkbox4:
                    #Запись запаса для длины
                    if type(supply) == System.String:
                        new_supply_quantity = doc.\
                            GetElement(el.GetTypeId()).\
                            LookupParameter(supply).\
                            AsDouble()
                    else:
                        new_supply_quantity = doc.\
                            GetElement(el.GetTypeId()).\
                            get_Parameter(supply).\
                            AsDouble()
                    if new_supply_quantity < 1:
                        new_supply_quantity = 1
                        if type(supply) == System.String:
                            doc.\
                                GetElement(el.GetTypeId()).\
                                LookupParameter(supply).\
                                Set(new_supply_quantity)
                        else:
                            doc.\
                                GetElement(el.GetTypeId()).\
                                get_Parameter(supply).\
                                Set(new_supply_quantity)
                #Запись единиц измерения
                if checkbox6:
                    unit_quantity = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                    if not unit_quantity:
                        try:
                            doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("шт.")
                        except:
                            try:
                                el.get_Parameter(guid_unit_param).Set("шт.")
                            except:
                                output.print_md('У элемента **{} с id {}** не заполнен ед. изм!'.format(el.Name,output.linkify(el.Id)))
                #Запись номера для спецификации
                if checkbox3:
                    num_all = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if num_all < 1:
                        new_num_all = num
                        try:
                            doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_all)
                        except:
                            output.print_md('У элемента **{} с id {}** параметр КП_И_Порядковый номер заблочен!'.format(el.Name,output.linkify(el.Id)))
            if num < 3:
                num += 1
            else:
                num = 6
        # #Задает для изоляции площадь:
        for cat in catList_insulations:
            elList_insulations = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_insulations:
                if checkbox1:
                    elRange_size_insulations = el.get_Parameter(BuiltInParameter.RBS_DUCT_CALCULATED_SIZE).AsString()
                    elRange_range_insulations = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString()
                    elRange_range_fiting = el.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble()
                    if elRange_range_fiting > 0:
                        if "x" in elRange_size_insulations:
                            elSize_rectangle = elRange_size_insulations.split('x')
                            square_rectangle = ((float(elSize_rectangle[0])/1000)+(float(elSize_rectangle[1])/1000))*2*(float(elRange_range_insulations))/1000
                            if dialog_out == 'Электролитный':
                                el.get_Parameter(guid_length_param).Set(square_rectangle/304.8)
                            else:
                                el.get_Parameter(guid_length_param).Set(square_rectangle)
                        else:
                            elSize_circle = elRange_size_insulations.split('ø')
                            square_circle = (float(elSize_circle[1])/1000)*3.14*(float(elRange_range_insulations))/1000
                            if dialog_out == 'Электролитный':
                                el.get_Parameter(guid_length_param).Set(square_circle/304.8)
                            else:
                                el.get_Parameter(guid_length_param).Set(square_circle)
                if checkbox5:
                    #Перенос толщины изоляции в КП_И_Толщина стенки
                    if cat == BuiltInCategory.OST_DuctInsulations:
                        elRange_fit_insulations = el.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_DUCT).AsValueString()
                    else:
                        elRange_fit_insulations = el.get_Parameter(BuiltInParameter.RBS_LINING_THICKNESS_FOR_DUCT).AsValueString()
                    elRange_fit_insulations_int = int(elRange_fit_insulations.split(' ')[0])/304.8
                    if elRange_fit_insulations_int > 0:
                        el.get_Parameter(guid_fitSize_param).Set(elRange_fit_insulations_int)
                if checkbox4:
                    #Запись запаса для длины
                    if type(supply) == System.String:
                        new_supply_Insulations = doc.\
                                GetElement(el.GetTypeId()).\
                                LookupParameter(supply).\
                                AsDouble()
                    else:
                        new_supply_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(supply).AsDouble()
                    if new_supply_Insulations < 1:
                        new_supply_Insulations = supply_value
                        if type(supply) == System.String:
                            doc.\
                                GetElement(el.GetTypeId()).\
                                LookupParameter(supply).\
                                Set(new_supply_Insulations)
                        else:
                            doc.\
                                GetElement(el.GetTypeId()).\
                                get_Parameter(supply).\
                                Set(new_supply_Insulations)
                #Запись единиц измерения
                if checkbox6:
                    unit_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                    if not unit_Insulations:
                        try:
                            doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("м2")
                        except:
                            try:
                                el.get_Parameter(guid_unit_param).Set("м2")
                            except:
                                output.print_md('У элемента **{} с id {}** не заполнен ед. изм!'.format(el.Name,output.linkify(el.Id)))
                #Запись номера для спецификации
                if checkbox3:
                    num_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if num_Insulations < 1:
                        new_num_Insulations = 5
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_Insulations)
        # Запись в параметр КП_О_Сортировка сумму параметров Семейство,Типоразмер,КП_Размер текст,КП_И_Толщина стенки
        if checkbox7:
            for all_cat in [catLits_range,
                            catList_quantity,
                            catList_insulations]:
                for cat in all_cat:
                    elList_all = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
                    for el in elList_all:
                        try:
                            el_name_sort = el.Name
                            el_family_name_sort = doc.GetElement(el.GetTypeId()).FamilyName
                            el_size_sort = el.get_Parameter(guid_size_param).AsString()
                            try:
                                try:
                                    el_fit_sort = round(el.get_Parameter(guid_fitSize_param).AsDouble() * 304.8, 1).ToString()
                                    sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort+'/'+el_fit_sort
                                except:
                                    sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort
                            except:
                                try:
                                    sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort
                                except:
                                    sort_param = el_family_name_sort+'/'+el_name_sort

                            el.get_Parameter(guid_sort_param).Set(sort_param)
                        except AttributeError as attr:
                            output.print_md(
                                "Ошибка {} у элемента {}. Работа остановлена!".
                                format(attr.ToString(), output.linkify(el.Id))
                            )
                            script.exit()
