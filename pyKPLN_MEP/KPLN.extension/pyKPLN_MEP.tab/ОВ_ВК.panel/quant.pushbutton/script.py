# -*- coding: utf-8 -*-
__title__ = '''Заполнить\nспецификацию'''
__doc__ = '''Плагин для заполнения параметров спецификации по ГОСТ'''

from Autodesk.Revit import DB
import os
from rpw.ui import forms
from pyrevit import script
from System.IO import StreamReader
import System
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

#region Узел ввода
path = r"X:\BIM\5_Scripts\Git_Repo_pyKPLN\pyKPLN_MEP\KPLN.extension\pyKPLN_MEP.tab\ОВ_ВК.panel\Спецификации.pushbutton"
files = [i[:-4] for i in os.listdir(path) if i.split(".")[-1] == "csv" and i != "script.py"] ## CSV

commands = [forms.CommandLink(i, return_value = i) for i in files]
dialog = forms.TaskDialog(
    'Выберите необходимый раздел',
    title_prefix = False,
    content = '',
    commands = commands,
    buttons = ['Cancel'],
    show_close = False)
dialog_out = dialog.show()

if dialog_out == None:
    script.exit()
#endregion

#region Идентификация предварительных настроек
categories = doc.Settings.Categories
categories = {category.Name : category.Id for category in categories}
duct_settings = DB.Mechanical.DuctSettings.GetDuctSettings(doc)
duct_connector_separator = duct_settings.ConnectorSeparator
shared_parameters_names = {str(parameter.GuidValue) : parameter.Name 
                           for parameter in DB.FilteredElementCollector(doc).\
                            OfClass(DB.SharedParameterElement).\
                            ToElements()}
geometry_option = __revit__.Application.Create.NewGeometryOptions()
#endregion

#region Основные методы
def identify_setting_parameter(parameter_name):
    '''Идентификация глобального параметра с настройками
    Используеться матричный формат данных settings = [[setting11, setting12, ...], [setting21, setting22, ...]]

    '''
    _parameters = DB.FilteredElementCollector(doc).\
    OfClass(DB.GlobalParameter)

    for _parameter in _parameters:
        if _parameter.Name == parameter_name:
            return _parameter.GetValue().Value

def get_parameter_by_name(element, parameter_name):
    '''Функция для поиска параметра в элементе по имени или guid'''

    if parameter_name[:5] == "guid:":
        parameter_name = shared_parameters_names[parameter_name[5:]]

    for parameter in element.Parameters: # Для параметров экземпляра
        if parameter.Definition.Name == parameter_name:
            return parameter
    try:
        for parameter in element.Symbol.Parameters: # Для параметров типа
            if parameter.Definition.Name == parameter_name:
                return parameter
    except:
        element = doc.GetElement(element.GetTypeId()) # Для параметров типа в системных семействах
        for parameter in element.Parameters: 
            if parameter.Definition.Name == parameter_name:
                return parameter

def get_numbers_from_string(value):
    '''Записать значения в массив если передается строка. 
    По сути все числа строки записываются в массив. 100х500/д500 = [100, 500, 500]

    '''
    string = ""
    for i in value:
        if i in "1234567890.":
            string += i
        else:
            string += ";"
    value = [float(i) for i in string.split(';') if len(i) != 0]
    if len(value) == 1:
        return value[0]
    else:
        return value

def convert_value(parameter, value = "Извлечь значение параметра"):
    '''Функция конвертирует или извлекает значение параметра

    '''
    if value == "Извлечь значение параметра":
        if str(parameter.StorageType) == "Double":
            return parameter.AsValueString()
        elif str(parameter.StorageType) == "ElementId":
            return parameter.AsValueString()
        elif str(parameter.StorageType) == "String":
            return parameter.AsString()
    else:
        if str(parameter.StorageType) == "Integer":
            if type(value) == str:
                value = get_numbers_from_string(value)
                return value
            else:
                return int(value)
        elif str(parameter.StorageType) == "String":
            return str(value)
        elif str(parameter.StorageType) == "Double":
            if type(value) == str:
                value = get_numbers_from_string(value)
                return value
            else:
                return float(value)
#endregion

#region Вспомогательные методы
def analyze_fitting_size_HVAC(sep, element, value):
    if str(element.MEPModel.PartType) in ["Elbow"]:
        angle_parameter = get_parameter_by_name(element, "guid:a7397d18-200b-4659-b34c-3d8ae1c54317")
        if angle_parameter == None:
            angle_value = "Ошибка !!!"
        else:
            angle_value = convert_value(angle_parameter)
        return value.split(sep)[0] + ", " + str(angle_value)
    
    elif str(element.MEPModel.PartType) in ["TapAdjustable", "Union", "Cap"]: # 
        return value.split(sep)[0]
    elif str(element.MEPModel.PartType) == "Transition":
        return value.split(sep)[0] + " на " + value.split(sep)[1]
    elif str(element.MEPModel.PartType) == "Tee":
        value = list(set(value.split(sep)))
        if len(value) == 1:
            return value[0]
        elif len(value) == 2:
            return value[0] + " , " + value[1]

def convert_value_for_duct_terminals(element, value):
    '''Конвертирует значения для воздухораспределителей
    
    '''
    if convert_value(get_parameter_by_name(element, "КП_О_Единица измерения")) == "шт.":
        return 1000000 ##Потому что переводим в метры квадратные
    elif convert_value(get_parameter_by_name(element, "КП_О_Единица измерения")) == "м²":
        if len(value) != 1:
            value = value[0] * value[1]
            return value
        else:
            value = ((value[0] * value[0]) / 4) * 3.14
            return value

def calculate_element_area(element):
    '''Написано chatGPT'''

    # Get the geometry of the element
    element_geo = element.get_Geometry(geometry_option)

    for i in element_geo:
        geometry_instance = i

    geometry_instance = geometry_instance.GetInstanceGeometry()
    
    # Initialize total area to 0
    total_area = 0
    total_shape = 0

    for connector in element.MEPModel.ConnectorManager.Connectors:
        if str(connector.Shape) == "Rectangular":
            shape = (connector.Height * connector.Width)
            total_shape += shape
        elif str(connector.Shape) == "Round":
            shape = (3.14 * connector.Radius * connector.Radius)
            total_shape += shape
    
    for geo_object in geometry_instance:
        if isinstance(geo_object, DB.Solid):
            # Get the surface area of the solid
            total_area += geo_object.SurfaceArea

    return (total_area - total_shape) / 10.76364864
#endregion

#region Обработка настроек
#       Вариант 1:
# settings = identify_setting_parameter("КП_Спецификация")
# settings = settings.split("\n")
# _ = []
# for setting in settings:
#     setting = [s.strip() for s in setting.split("\t")]
#     _.append(setting)
# settings = _
#       Вариант 2:
# exec("settings = " + dialog_out + ".settings")
#       Вариант 3:
encoding = System.Text.Encoding.GetEncoding(1251)
file_settings = StreamReader(path + "\\" + dialog_out + ".csv")
settings = file_settings.ReadToEnd()
settings = settings.split("\n")
_ = []
for setting in settings:
    setting = [s.strip() for s in setting.split(";")]
    _.append(setting)
settings = _
file_settings.Close()
#endregion 

#region Основной цикл
transaction = DB.Transaction(doc, "КП_Спецификация")
transaction.Start()

list_error_id = []

for setting in settings:
    if setting[0].capitalize() == "Категория":
        elements_of_category = DB.FilteredElementCollector(doc).\
        OfCategoryId(categories[setting[1].capitalize()]).\
        WhereElementIsNotElementType().\
        ToElements()
        for element in elements_of_category:
            try:
                parameter = get_parameter_by_name(element, setting[2])
                value = convert_value(parameter, setting[3])
                parameter.Set(value)
            except:
                list_error_id.append(element.Id)

    if setting[0].capitalize() == "Внутри элемента":
        elements_of_category = DB.FilteredElementCollector(doc).\
        OfCategoryId(categories[setting[1].capitalize()]).\
        WhereElementIsNotElementType().\
        ToElements()
        for element in elements_of_category:
            try:
                parameter_from = get_parameter_by_name(element, setting[2])
                parameter_to = get_parameter_by_name(element, setting[3])
                value_from = convert_value(parameter_from)
                value = convert_value(parameter_to, value_from)
                if setting[4] != "value":
                    exec(setting[4])
                parameter_to.Set(value)
            except:
                list_error_id.append(element.Id)

    if setting[0].capitalize() == "Объеденить параметры":
        elements_of_category = DB.FilteredElementCollector(doc).\
        OfCategoryId(categories[setting[1].capitalize()]).\
        WhereElementIsNotElementType().\
        ToElements()
        for element in elements_of_category:
            try:
                #region Преобразовать параметры
                parameters = setting[2].split("+")
                parameters = [parameter.strip() for parameter in parameters]
                #endregion
                string = ""
                for parameter in parameters:
                    try:
                        parameter = get_parameter_by_name(element, parameter)
                        value = convert_value(parameter)
                    except:
                        pass
                    string += str(value) + "/"
                string = string [:-1]
                parameter_for_write = get_parameter_by_name(element, setting[3])
                parameter_for_write.Set(string) ## Не преобразовываю, потому что подразумеванется, что будет только строка
            except:
                list_error_id.append(element.Id)

    if setting[0].capitalize() == "Записать из хоста":
        elements_of_category = DB.FilteredElementCollector(doc).\
        OfCategoryId(categories[setting[1].capitalize()]).\
        WhereElementIsNotElementType().\
        ToElements()
        for element in elements_of_category:
            try:
                host = doc.GetElement(element.HostElementId)
                parameter_host = get_parameter_by_name(host, setting[2])
                value_host = convert_value(parameter_host)
                parameter = get_parameter_by_name(element, setting[3])
                value = convert_value(parameter, value_host)
                if setting[4] != "value":
                    exec(setting[4])
                parameter.Set(value)
            except:
                list_error_id.append(element.Id)

#region Обработка ошибок
if len(list_error_id) != 0:
    print("При записи значений для следующих элементов произошла ошибка")
    print("Список элементов: ")
    # print(str([int(i) for i in list_error_id]))
else:
    print("Запись прошла без ошибок")
#endregion

transaction.Commit()
#endregion
            



