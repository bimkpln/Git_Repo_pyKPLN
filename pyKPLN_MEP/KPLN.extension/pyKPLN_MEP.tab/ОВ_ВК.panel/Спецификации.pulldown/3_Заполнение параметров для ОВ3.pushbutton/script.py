# coding: utf-8

__title__ = "ОВ1/3. Заполнение параметров для спецификации по нескольким категориям"
__author__ = 'Kapustin Roman'
__doc__ = "Заполняет:\n 1.Для воздуховодв длину и габариты, считая воздуховоды MxN и NxM как MxN \n 2.Для элементов в шт. число. \n 3.Для изоляции площадь и толщину стенки."


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter
from rpw import revit, db
from pyrevit import script
import codecs
from System import Guid
from rpw.ui.forms import*


# main code: input part
doc = revit.doc
output = script.get_output()
# КП_И_Количество в спецификацию
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
commands= [CommandLink('Revit2018', return_value='2018'),
            CommandLink('Revit2020', return_value='2020')]
dialog = TaskDialog('Выберите версию Revit:',
                commands=commands,
                show_close=True)
dialog_out = dialog.show()
if dialog_out == '2020':
    # КП_И_Запас
    supply = Guid("62ff7f1d-f5cd-4e5e-9f50-b7b3e18f2c06")
elif dialog_out == '2018':
    supply = Guid("5585ac90-f2d5-462b-a474-b97eb1161756")
else:
    script.exit()

catLits_range = [BuiltInCategory.OST_FlexPipeCurves,BuiltInCategory.OST_PipeCurves]
catList_quantity = [BuiltInCategory.OST_MechanicalEquipment,BuiltInCategory.OST_PipeAccessory,BuiltInCategory.OST_PipeFitting]
catList_insulations = [BuiltInCategory.OST_PipeInsulations]
with db.Transaction('Заполнение корректной длины для спецификации по нескольким категориям'):
    components = [Label('Выберите параметры для заполнения:'),
                CheckBox('checkbox1', 'КП__И__Количество в спецификацию',default=True),
                CheckBox('checkbox2', 'КП__Размер__Текст',default=True),
                CheckBox('checkbox3', 'КП__И__Порядковый номер',default=True),
                CheckBox('checkbox4', 'КП__И__Запас',default=True),
                Label('Размер запаса:'),
                TextBox('textbox1', Text="1.15"),
                CheckBox('checkbox5', 'КП__И__Толщина стенки (для изоляции)',default=True),
                CheckBox('checkbox6', 'КП__О__Единица измерения',default=True),
                CheckBox('checkbox7', 'КП__О__Сортирование',default=False),
                Separator(),
                Button('Далее')]
    form = FlexForm('ОВ1/3_Спецификации', components)
    form.show()
    supply_value = float(form.values.get('textbox1'))
    if supply_value < 1 or supply_value >= 2:
        Alert('Ввден некорректный запас!', title='Ошибка', header='', exit=True)
    checkbox1 = form.values.get('checkbox1')
    checkbox2 = form.values.get('checkbox2')
    checkbox3 = form.values.get('checkbox3')
    checkbox4 = form.values.get('checkbox4')
    checkbox5 = form.values.get('checkbox5')
    checkbox6 = form.values.get('checkbox6')
    checkbox7 = form.values.get('checkbox7')
    if form.values != {}:
        # Перенос длины в категории в метрах
        #     Трубы, гибкие трубы 
        for cat in catLits_range:
            elList_range = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_range:
                if checkbox1:
                    elRange_range = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString()
                    Range = (float(elRange_range))/1000
                    el.get_Parameter(guid_length_param).Set(Range)
            #Запись в параметр КП_Размер_Текст размера труб
                if checkbox2:
                    elRange_size_curve = el.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
                    if " мм" in elRange_size_curve:
                        size_curve_mm = elRange_size_curve.split(" мм")
                        el.get_Parameter(guid_size_param).Set(size_curve_mm[0])
                    elif "мм" in elRange_size_curve:
                        size_curve_mm = elRange_size_curve.split(" ")
                        el.get_Parameter(guid_size_param).Set(size_curve_mm[0])
                    elif "ш" in elRange_size_curve:
                            size_curve_mm = elRange_size_curve.split("ш")
                            el.get_Parameter(guid_size_param).Set(size_curve_mm[0])
                    else:
                        el.get_Parameter(guid_size_param).Set(elRange_size_curve)
                if checkbox4:
                    #Запись запаса для длины
                    new_supply_duct = doc.GetElement(el.GetTypeId()).get_Parameter(supply).AsDouble()
                    if new_supply_duct < 1:
                        new_supply_duct = supply_value
                        doc.GetElement(el.GetTypeId()).get_Parameter(supply).Set(new_supply_duct)
                #Запись единиц измерения
                if checkbox6:
                    try:
                        unit_duct = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                        if not unit_duct:
                            doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("м")
                    except:
                        unit_duct = el.get_Parameter(guid_unit_param).AsString()
                        if not unit_duct:
                            try:
                                el.get_Parameter(guid_unit_param).Set("м")
                            except:
                                output.print_md('У элемента **{} с id {}** не заполнен ед. изм!'.format(el.Name,output.linkify(el.Id)))
                #Запись номера для спецификации
                if checkbox3:
                    num_duct = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if not num_duct:
                        new_num_duct = 3
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_duct)
                #Задает 1 для категорий:
                #Оборудование,  арматура труб,  соеденительные детали труб.
        num = 1
        for cat in catList_quantity:
            elList_quantity = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_quantity:
                if checkbox1:
                    el.get_Parameter(guid_length_param).Set(1)
                if checkbox4:
                    #Запись запаса для длины
                    new_supply_quantity = doc.GetElement(el.GetTypeId()).get_Parameter(supply).AsDouble()
                    if new_supply_quantity < 1:
                        new_supply_quantity = 1
                        doc.GetElement(el.GetTypeId()).get_Parameter(supply).Set(new_supply_quantity)
                #Запись единиц измерения
                if checkbox6:
                    try:
                        unit_quantity = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                        if not unit_quantity:
                            doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("шт.")
                    except:
                        try:
                            unit_quantity = el.get_Parameter(guid_unit_param).AsString()
                            if not unit_quantity:
                                el.get_Parameter(guid_unit_param).Set("шт.")
                        except:
                            output.print_md('У элемента **{} с id {}** не заполнен ед. изм!'.format(el.Name,output.linkify(el.Id)))
                #Запись номера для спецификации
                if checkbox3:
                    num_all = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if not num_all:
                        new_num_all = num
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_all)                    
            if num < 2:
                num += 1
            else:
                num = 5
        # Изоляция
        for cat in catList_insulations:
            elList_insulations = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_insulations:   
                # Длина
                elRange_range_insulations = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString()
                if checkbox1:
                    if elRange_range_insulations > '0':
                        Range_insulations = (float(elRange_range_insulations))/1000
                        el.get_Parameter(guid_length_param).Set(Range_insulations)                        
                    else:
                        el.get_Parameter(guid_length_param).Set(0) 
                # Размер
                if checkbox2:
                    elRange_size_curve = el.get_Parameter(BuiltInParameter.RBS_PIPE_CALCULATED_SIZE).AsString()
                    if elRange_size_curve > 0:
                        if '-' in elRange_size_curve:
                            el.get_Parameter(guid_size_param).Set('Изоляция фасонных элементов')
                        elif '/' in elRange_size_curve:
                            el.get_Parameter(guid_size_param).Set('Изоляция фасонных элементов')
                        elif " мм" in elRange_size_curve:
                            size_curve_mm = elRange_size_curve.split(" мм")
                            el.get_Parameter(guid_size_param).Set(size_curve_mm[0])
                        elif "мм" in elRange_size_curve:
                            size_curve_mm = elRange_size_curve.split(" ")
                            el.get_Parameter(guid_size_param).Set(size_curve_mm[0])
                        elif "ш" in elRange_size_curve:
                            size_curve_mm = elRange_size_curve.split("ш")
                            el.get_Parameter(guid_size_param).Set(size_curve_mm[0])                        
                        else:
                            el.get_Parameter(guid_size_param).Set(elRange_size_curve)
                if checkbox5:
                    #Перенос толщины изоляции в КП_И_Толщина стенки
                    elRange_fit_insulations = el.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_PIPE).AsValueString()
                    elRange_fit_insulations_int = int(elRange_fit_insulations.split(' ')[0])/304.8
                    if elRange_fit_insulations_int > 0:
                        el.get_Parameter(guid_fitSize_param).Set(elRange_fit_insulations_int)
                if checkbox4:
                    #Запись запаса для длины
                    new_supply_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(supply).AsDouble()
                    if new_supply_Insulations < 1:
                        new_supply_Insulations = supply_value
                        doc.GetElement(el.GetTypeId()).get_Parameter(supply).Set(new_supply_Insulations)
                #Запись единиц измерения
                if checkbox6:
                    try:
                        unit_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                        if not unit_Insulations:
                            doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("м")
                    except:
                        unit_Insulations = el.get_Parameter(guid_unit_param).AsString()
                        try:
                            if not unit_Insulations:
                                el.get_Parameter(guid_unit_param).Set("м")
                        except:
                            output.print_md('У элемента **{} с id {}** не заполнен ед. изм!'.format(el.Name,output.linkify(el.Id)))
                #Запись номера для спецификации
                if checkbox3:
                    num_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if not num_Insulations:
                        new_num_Insulations = 4
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_Insulations)
        # Запись в параметр КП_О_Сортировка сумму параметров Семейство,Типоразмер,КП_Размер текст,КП_И_Толщина стенки
        if checkbox7:
            for all_cat in [catLits_range,catList_quantity,catList_insulations]:
                for cat in all_cat:
                    elList_all = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
                    for el in elList_all:
                        el_name_sort = el.Name
                        el_family_name_sort = doc.GetElement(el.GetTypeId()).FamilyName
                        el_size_sort = el.get_Parameter(guid_size_param).AsString()
                        el_fit_sort = el.get_Parameter(guid_fitSize_param).AsValueString()                   
                        try:    
                            try:
                                sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort+'/'+el_fit_sort
                            except:
                                sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort
                        except:
                            try:
                                sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort
                            except:
                                sort_param = el_family_name_sort+'/'+el_name_sort

                        el.get_Parameter(guid_sort_param).Set(sort_param)