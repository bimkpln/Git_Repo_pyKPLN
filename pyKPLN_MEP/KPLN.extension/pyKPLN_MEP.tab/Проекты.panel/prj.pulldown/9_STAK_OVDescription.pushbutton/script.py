# -*- coding: utf-8 -*-

__title__ = "СТАК_Заполнение параметров 'КП_И_Описание'"
__author__ = 'Tima Kutsko'
__doc__ = "Автоматическое заполнение параметров для дополнительного\n"\
        "параметра специфкации"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
    Transaction
from rpw import revit
import math


# Функция для нахождения всех экземпляров семейств с именем
def find_family_instances_with_prefix(doc, prefix):
    collector = FilteredElementCollector(doc)
    collector.OfCategory(BuiltInCategory.OST_MechanicalEquipment)
    collector.WhereElementIsNotElementType()
    family_instances = [elem for elem in collector if elem.Symbol.Family.Name.startswith(prefix)]
    return family_instances


# Функция для объединения данных из параметров и записи их в "Параметр3"
def combine_and_set_parameters(instance, param1_name, param2_name, param3_name):
    param1_value = instance.LookupParameter(param1_name).AsString()
    param2_value = instance.LookupParameter(param2_name).AsString()
    combined_value = param1_value + param2_value
    instance.LookupParameter(param3_name).Set(combined_value)


# Основная функция плагина
def main():
    # Начинаем транзакцию
    doc = revit.doc
    transaction = Transaction(doc, "Обновление параметров экземпляров семейств")
    transaction.Start()

    # Поиск экземпляров семейств с заданным префиксом
    family_instances = find_family_instances_with_prefix(doc, "555_")
    family_instances.extend(find_family_instances_with_prefix(doc, "550_"))

    # Обработка найденных экземпляров
    for instance in family_instances:
        g_param = instance.LookupParameter("КП_И_Расход воздуха")
        if g_param is None:
            g_param = doc.GetElement(instance.GetTypeId()).LookupParameter("КП_И_Расход воздуха")
        g = "L=" + math.ceil(g_param.AsDouble() * 101.94).ToString() + " м³/ч"

        prD_param = instance.LookupParameter("КП_И_Потеря давления воздуха")
        if prD_param is None:
            prD_param = doc.GetElement(instance.GetTypeId()).LookupParameter("КП_И_Потеря давления воздуха")
        prD = "P=" + math.ceil(prD_param.AsDouble() * 3.277).ToString() + " Па"

        nP_param = instance.LookupParameter("КП_И_Номинальная мощность")
        if nP_param is None:
            nP_param = doc.GetElement(instance.GetTypeId()).LookupParameter("КП_И_Номинальная мощность")
        nP = "N=" + round(nP_param.AsDouble() / 10764, 3).ToString() + " кВт"

        v_param = instance.LookupParameter("КП_И_Частота вращения двигателя")
        if v_param is None:
            v_param = doc.GetElement(instance.GetTypeId()).LookupParameter("КП_И_Частота вращения двигателя")
        v = "n=" + (v_param.AsInteger()).ToString() + " об/мин"

        u_param = instance.LookupParameter("КП_И_Напряжение")
        if u_param is None:
            u_param = doc.GetElement(instance.GetTypeId()).LookupParameter("КП_И_Напряжение")
        u = "U=" + round(u_param.AsDouble() / 10.764, 0).ToString() + " В"

        instance.LookupParameter("КП_И_Описание").Set(", ".join([g, prD, nP, v, u]))

    # Завершаем транзакцию
    transaction.Commit()

# Запускаем основную функцию
main()