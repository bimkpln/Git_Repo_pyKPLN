# coding: utf-8

__title__ = "СИТИ_Имя_Классификатора"
__author__ = 'Tima Kutsko'
__doc__ = '''Заполняет параметр 'СИТИ_Имя_Классификатора'\n для элементов ИОС
по параметру СИТИ_Классификатор'''


from Autodesk.Revit.DB import FilteredElementCollector, Element, Mechanical
from rpw import revit, db
from pyrevit import script
import codecs
from System import Guid


# main code: input part
doc = revit.doc
output = script.get_output()
# СИТИ_Классификатор
rbs_param = Guid('32bf9389-a9b8-4db4-8d67-2fce3844b607')
# СИТИ_Имя_классификатора
rbs_name_param = Guid('5ec9640c-c857-4c3b-9119-c6fd79f820e2')
codeErrorText = "Код либо не заполнен, либо указан код группы"
paramErrorText = "У элемента/-ов нет параметра/-ов"
error_dict = dict()
file_path_part1 = 'Y:\\Жилые здания\\ДИВНОЕ СитиXXI'
file_path_part2 = '\\7. Стадия РД\\Исходные данные'
file_path_part3 = '\\Классификатор City-ХХI век\\CITY_Классификатор_csv.txt'
file_path = file_path_part1 + file_path_part2 + file_path_part3
element_collector = FilteredElementCollector(doc, doc.ActiveView.Id).\
                    WhereElementIsNotElementType().\
                    ToElements()
# processing data
with db.Transaction("СИТИ_Имя_Классификатора"):
    with codecs.open(file_path, encoding='utf-8') as rbs_list:
        for element in element_collector:
            if type(element) != Element and\
                    type(element) != Mechanical.MechanicalSystem:
                try:
                    inpDataParam = element.get_Parameter(rbs_param).AsString()
                    if inpDataParam is None:
                        inpDataParam = doc.\
                            GetElement(element.GetTypeId()).\
                            get_Parameter(rbs_param).AsString()
                    if inpDataParam is not None:
                        true_data = None
                        for line in rbs_list:
                            if inpDataParam in line.strip().split(';')\
                                    and line.strip().split(';')[-1] == "Item":
                                true_data = line.strip().split(';')[-2]
                        if true_data is None:
                            try:
                                error_dict[codeErrorText].append(element)
                            except KeyError:
                                error_dict[codeErrorText] = list()
                        else:
                            try:
                                element.\
                                get_Parameter(rbs_name_param).\
                                Set(true_data)
                            except Exception:
                                doc.\
                                GetElement(element.GetTypeId()).\
                                get_Parameter(rbs_name_param).\
                                Set(true_data)
                except AttributeError:
                    if type(element) != Element:
                        try:
                            error_dict[paramErrorText].append(element)
                        except KeyError:
                            error_dict[paramErrorText] = list()
# error list output
if len(error_dict.values()) != 0:
    for key, values in error_dict.items():
        print('{}:'.format(key))
        for value in values:
            print('{}'.format(output.linkify(value.Id)))
