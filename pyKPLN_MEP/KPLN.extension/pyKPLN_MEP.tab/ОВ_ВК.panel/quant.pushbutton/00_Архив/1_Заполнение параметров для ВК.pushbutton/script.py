# coding: utf-8

__title__ = "ВК. Заполнение параметров для спецификации по нескольким категориям"
__author__ = 'Kapustin Roman'
__doc__ = "Заполняет:\n 1."


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
                              BuiltInParameter
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
# КП_И_Порядковый номер
guid_num_param = Guid("7a44592a-c698-4276-af5f-75be9b131b40")
# КП_И_Запас
supply = Guid("62ff7f1d-f5cd-4e5e-9f50-b7b3e18f2c06")
# КП_О_Единица измерения
guid_unit_param = Guid("4289cb19-9517-45de-9c02-5a74ebf5c86d")
catLits_range = [BuiltInCategory.OST_FlexPipeCurves,BuiltInCategory.OST_PipeCurves]
catList_quantity = [BuiltInCategory.OST_PlumbingFixtures,BuiltInCategory.OST_MechanicalEquipment,BuiltInCategory.OST_PipeAccessory,BuiltInCategory.OST_PipeFitting]
catList_insulations = [BuiltInCategory.OST_PipeInsulations]
with db.Transaction('Заполнение корректной длины для спецификации по нескольким категориям'):   
    components = [Label('Выберите параметры для заполнения:'),
                CheckBox('checkbox1', 'КП__И__Количество в спецификацию',default=True),
                CheckBox('checkbox2', 'КП__Размер__Текст',default=True),
                CheckBox('checkbox3', 'КП__И__Порядковый номер',default=True),
                CheckBox('checkbox4', 'КП__И__Запас',default=True),
                CheckBox('checkbox5', 'КП__О__Единица измерения',default=True),
                Separator(),
                Button('Далее')]
    form = FlexForm('ВК_Спецификации', components)
    form.show()
    checkbox1 = form.values.get('checkbox1')
    checkbox2 = form.values.get('checkbox2')
    checkbox3 = form.values.get('checkbox3')
    checkbox4 = form.values.get('checkbox4')
    checkbox5 = form.values.get('checkbox5')
    if form.values != {}:      
        for cat in catLits_range:
            elList_range = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_range:
                # Перенос длины в категории в метрах
                #     Трубы, гибкие трубы   
                if checkbox1:
                    elRange_range = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString()
                    Range = (float(elRange_range))/1000
                    el.get_Parameter(guid_length_param).Set(Range)
                # Запись в параметр КП_Размер_Текст размер труб
                if checkbox2:
                    elRange_size_curve = el.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString()
                    if " мм" in elRange_size_curve:
                        size_curve_mm = elRange_size_curve.split(" мм")
                        el.get_Parameter(guid_size_param).Set(size_curve_mm[0])
                    elif "мм" in elRange_size_curve:
                        size_curve_mm = elRange_size_curve.split(" ")
                        el.get_Parameter(guid_size_param).Set(size_curve_mm[0])
                    else:
                        el.get_Parameter(guid_size_param).Set(elRange_size_curve)
                #Запись запаса для длины
                if checkbox4:
                    new_supply_curve = doc.GetElement(el.GetTypeId()).get_Parameter(supply).AsDouble()
                    if new_supply_curve < 1:
                        new_supply_curve = 1.1
                        doc.GetElement(el.GetTypeId()).get_Parameter(supply).Set(new_supply_curve)
                #Запись номера для спецификации
                if checkbox3:
                    num_curve = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if num_curve < 1:
                        new_num_curve = 5
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_curve)
                #Запись единиц измерения
                if checkbox5:
                    unit_curve = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                    if not unit_curve:
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("м")
        for cat in catList_insulations:
            elList_range_insulations = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_range_insulations:
                #Перенос длины в категории в метрах
                #Изоляция
                if checkbox1:
                    elRange_range_insulations = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString()
                    Range_insulations = (float(elRange_range_insulations))/1000
                    el.get_Parameter(guid_length_param).Set(Range_insulations)
                #Запись в параметр КП_Размер_Текст размер труб
                if checkbox2:
                    elRange_size_insulations = el.get_Parameter(BuiltInParameter.RBS_PIPE_CALCULATED_SIZE).AsString()
                    if " мм" in elRange_size_insulations:
                        size_insulations_mm = elRange_size_insulations.split(" мм")
                        el.get_Parameter(guid_size_param).Set(size_insulations_mm[0])
                    elif "мм" in elRange_size_insulations:
                        size_insulations_mm = elRange_size_insulations.split(" ")
                        el.get_Parameter(guid_size_param).Set(size_insulations_mm[0])
                    else:
                        el.get_Parameter(guid_size_param).Set(elRange_size_insulations)
                #Запись запаса для длины
                if checkbox4:
                    new_supply_insulations = doc.GetElement(el.GetTypeId()).get_Parameter(supply).AsDouble()
                    if new_supply_insulations < 1:
                        new_supply_insulations = 1.1
                        doc.GetElement(el.GetTypeId()).get_Parameter(supply).Set(new_supply_insulations)
                #Запись номера для спецификации
                if checkbox3:
                    num_insulation = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if num_insulation < 1:
                        new_num_insulation = 6
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_insulation)
                #Запись единиц измерения
                if checkbox5:
                    unit_insulation = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                    if not unit_insulation:
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("м")     
        num = 1
        for cat in catList_quantity:
            elList_quantity = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
            for el in elList_quantity:
                #Задает 1 для категорий:   
                if checkbox1:
                    el.get_Parameter(guid_length_param).Set(1)
                #Запись запаса 1
                if checkbox4:
                    new_supply_quantity = doc.GetElement(el.GetTypeId()).get_Parameter(supply).AsDouble()
                    if new_supply_quantity < 1:
                        new_supply_quantity = 1
                        doc.GetElement(el.GetTypeId()).get_Parameter(supply).Set(new_supply_quantity)                    
                #Запись единиц измерения
                if checkbox5:
                    unit_all = doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).AsString()
                    if not unit_all:
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_unit_param).Set("шт.")
                #Запись номера для спецификации
                if checkbox3:
                    num_all = doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).AsDouble()
                    if num_all < 1:
                        new_num_all = num
                        doc.GetElement(el.GetTypeId()).get_Parameter(guid_num_param).Set(new_num_all)
            num += 1