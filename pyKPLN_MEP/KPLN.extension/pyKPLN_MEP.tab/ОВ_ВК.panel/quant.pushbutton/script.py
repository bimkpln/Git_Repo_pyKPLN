# -*- coding: utf-8 -*-
__title__ = '''Заполнить\nспецификацию'''
__doc__ = '''Плагин для заполнения параметров спецификации по ГОСТ'''
__highlight__ = 'updated'

from Autodesk.Revit import DB
import os
from rpw.ui import forms
from pyrevit import script
from pyrevit import forms as pyforms
from System.IO import StreamReader, StreamWriter
import kpln_logs
import System
import time

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

#region Параметры для логирования в Extensible Storage. Не менять по ходу изменений кода
extStorage_guid = "720080C5-DA99-40D7-9445-E53F288AA139"
extStorage_name = "kpln_ios_quant"
#endregion

if __shiftclick__:
   try:
       obj = kpln_logs.create_obj(extStorage_guid, extStorage_name)
       kpln_logs.read_log(obj)
   except:
       print("Логи запуска программы отсутствуют. Плагин в этом проекте ниразу не запускался")
   script.exit()

#region Узел ввода
path = path = __file__[:-10]
files = [i[:-4] for i in os.listdir(path) if i.split(".")[-1] == "csv" and i != "script.py"] ## CSV

commands = [forms.CommandLink(i, return_value = i) for i in files]
dialog_out = pyforms.SelectFromList.show(files, button_name='Выберите необходимую конфигурацию')

if dialog_out == None:
    script.exit()
#endregion

#region Идентификация предварительных настроек
categories = doc.Settings.Categories
categories = {category.Name : category.Id for category in categories}

duct_settings = DB.Mechanical.DuctSettings.GetDuctSettings(doc)
duct_connector_separator = duct_settings.ConnectorSeparator
pipe_settings = DB.Plumbing.PipeSettings.GetPipeSettings(doc)
pipe_connector_separator = pipe_settings.ConnectorSeparator

shared_parameters_names = {str(parameter.GuidValue) : parameter.Name 
                           for parameter in DB.FilteredElementCollector(doc).\
                            OfClass(DB.SharedParameterElement).\
                            ToElements()}

geometry_option = __revit__.Application.Create.NewGeometryOptions()
foot_to_mm = 0.0032808 ## Коэфициент, на который нужно поделить футы, чтобы получить миллиметры

# Переменные для кода
list_error_id = {} # Обработчик ошибок
time_log = {} # Логи времени выполнения
#endregion


#region Основные методы
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
        elif str(parameter.StorageType) == "Integer":
            return parameter.AsInteger()
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
    
    elif str(element.MEPModel.PartType) in ["TapAdjustable", "Union", "Cap", "PipeFlange"]: # 
        return value.split(sep)[0]
    elif str(element.MEPModel.PartType) in ["Transition", "MultiPort", "SpudPerpendicular", "SpudAdjustable"]:
        return value.split(sep)[0] + "/" + value.split(sep)[1]
    elif str(element.MEPModel.PartType) in ["Tee", "Cross"]:
        value = list(set(value.split(sep)))
        if len(value) == 1:
            return value[0]
        elif len(value) == 2:
            return value[0] + "/" + value[1]

def convert_value_for_duct_terminals(element, value):
    '''Конвертирует значения для воздухораспределителей
    
    '''
    if convert_value(get_parameter_by_name(element, "КП_О_Единица измерения")) == "шт.":
        stock = get_parameter_by_name(element, "КП_И_Запас")
        stock_value = convert_value(stock)
        return 1000000 / float(stock_value) ##Потому что переводим в метры квадратные и запас учитывается в конфигурации
    elif convert_value(get_parameter_by_name(element, "КП_О_Единица измерения")) == "м²":
        if type(value) == list:
            value = value[0] * value[1]
            return value
        else:
            value = ((value * value) / 4) * 3.14
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

def calcucate_pipe_thickness(element): #Изменено
   OD = get_parameter_by_name(element, "Внешний диаметр").AsDouble()
   ID = get_parameter_by_name(element, "Внутренний диаметр").AsDouble()

   return ((float(OD) - float(ID)) / 2)

# def calculate_size_lightning_protection(element):
#     width = convert_value(get_parameter_by_name(element, "Ширина"))
#     height = convert_value(get_parameter_by_name(element, "Высота"))
    
#     min_size = width if int(width) < int(height) else height
#     max_size = width if int(width) > int(height) else height

#     return min_size + "x" + max_size

## Переменные, которые вычисляют значения для функции real_pipe_size()
pipe_collector = DB.FilteredElementCollector(doc).\
OfCategory(DB.BuiltInCategory.OST_PipeCurves).\
WhereElementIsElementType()
flex_pipe_collector = DB.FilteredElementCollector(doc).\
OfCategory(DB.BuiltInCategory.OST_FlexPipeCurves).\
WhereElementIsElementType()
d_inner_param = "guid:4779ced4-8b54-4308-aefa-5c8f418da90c"
d_outer_param = "guid:d5799102-d7da-404a-bd64-1c50984bce7d"
d_nomin_param = "guid:c758aee7-e324-4335-8c22-16d458f8737e"
try:
    pipe__diams_dict = {i.Id : {d_inner_param : get_parameter_by_name(i, d_inner_param).AsDouble(),
                                d_outer_param : get_parameter_by_name(i, d_outer_param).AsDouble(),
                                d_nomin_param : get_parameter_by_name(i, d_nomin_param).AsDouble()} for i in pipe_collector}
except:
    print("Для Трубы не определены параметры: " + "С_Внешний диаметр трубы" + " " + "С_Внутренний диаметр трубы" + " " + "С_Услоный диаметр трубы")
    print("По умолчанию в КП_Размер_Текст будет записан условный диаметр трубы")
    print("-----------------------------------------------------------")
try:
    flex_pipe__diams_dict = {i.Id : {d_inner_param : get_parameter_by_name(i, d_inner_param).AsDouble(),
                                d_outer_param : get_parameter_by_name(i, d_outer_param).AsDouble(),
                                d_nomin_param : get_parameter_by_name(i, d_nomin_param).AsDouble()} for i in flex_pipe_collector}
except:
    print("Для Гибкие трубы не определены параметры: " + "С_Внешний диаметр трубы" + " " + "С_Внутренний диаметр трубы" + " " + "С_Услоный диаметр трубы")
    print("По умолчанию в КП_Размер_Текст будет записан условный диаметр трубы")
    print("-----------------------------------------------------------")
def real_pipe_size(element):
    '''
    Функция, которая в зависимости от заданных параметров записывает возвращает диаметр трубы
    Внутренний
    Внешний
    или Условный
    '''

    diam = element.Parameter[DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM].AsDouble()
    if element.Category.Name == "Трубы":
        try:
            pipe = pipe__diams_dict[element.PipeType.Id]
        except:
            return diam / foot_to_mm
    elif element.Category.Name == "Гибкие трубы":
        try:
            pipe = flex_pipe__diams_dict[element.FlexPipeType.Id]
        except:
            return diam / foot_to_mm

    if pipe[d_nomin_param] == 0 and pipe[d_outer_param] == 0 and pipe[d_inner_param] == 0:
        return diam / foot_to_mm
    else:
        _max_value = 99999
        ret_value = ""
        for i, j in pipe.items():
            if diam < j and j < _max_value:
                ret_value = i
                _max_value = j

        if ret_value == d_nomin_param:
            return diam / foot_to_mm
        elif ret_value == d_outer_param:
            return element.Parameter[DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER].AsDouble() / foot_to_mm
        elif ret_value == d_inner_param:
            return element.Parameter[DB.BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM].AsDouble() / foot_to_mm
#endregion

#region Внутренние функции для кода
def create_correc_dict(key, value, dict):
    if dict.get(key) == None:
        dict[key] = value
    else:
        if type(dict[key]) in [int, float]:
            dict[key] += value
        elif type(dict[key]) in [list]:
            dict[key].extend(value)
    return dict

def create_string_from_list(list):
    ret_string = ""
    ret_string += "".join(str(i) + "," for i in list)
    ret_string = ret_string[:-1]
    return ret_string
#endregion


#region Обработка настроек
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

def main1():
    global string
    for parameter in parameters:
        if (parameter[0] == "'" or parameter[0] == '"') and (parameter[-1] == "'" or parameter[-1] == '"'):
            parameter = parameter[1:-1]
            value = parameter
        else:
            try:
                parameter = get_parameter_by_name(element, parameter)
                if parameter.Definition.Name == "КП_И_Толщина стенки": ## Отстойный хардкод, потом поправлю
                    value = round(parameter.AsDouble() / foot_to_mm, 1)
                else:
                    value = convert_value(parameter)
            except:
                value = "Не найден"
        string += str(value) + delimiter

print("-----------------------------------------------------------")
for setting in settings:
    if setting[0].capitalize() == "Категория":
        s = time.time()
        elements_of_category = DB.FilteredElementCollector(doc).\
        OfCategoryId(categories[setting[1].capitalize()]).\
        WhereElementIsNotElementType().\
        ToElements()
        for element in elements_of_category:
            try:
                parameter = get_parameter_by_name(element, setting[2])
                value = convert_value(parameter, setting[3])
                parameter.Set(value)
            except Exception as e:
                list_error_id = create_correc_dict(str(e) + "&" + create_string_from_list(setting), [element.Id.IntegerValue], list_error_id)
        f = time.time()
        time_log = create_correc_dict(setting[0].capitalize(), f-s, time_log)

    if setting[0].capitalize() == "Внутри элемента":
        s = time.time()
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
            except Exception as e:
                list_error_id = create_correc_dict(str(e) + "&" + create_string_from_list(setting), [element.Id.IntegerValue], list_error_id)
        f = time.time()
        time_log = create_correc_dict(setting[0].capitalize(), f-s, time_log)

    if setting[0].capitalize() == "Объеденить параметры":
        s = time.time()
        elements_of_category = DB.FilteredElementCollector(doc).\
        OfCategoryId(categories[setting[1].capitalize()]).\
        WhereElementIsNotElementType().\
        ToElements()
        delimiter = setting[4]
        if setting[5] == "":
            filter_family_name = ""
        else:
            filter_family_name = setting[5].strip()
        for element in elements_of_category:
            if element.Category.Name in ["Трубы", "Гибкие трубы", "Воздуховоды", "Гибкие воздуховоды"]:
                family_name = str(element.Parameter[DB.BuiltInParameter.ELEM_TYPE_PARAM].AsValueString())[:len(filter_family_name)]
            else:
                family_name = str(element.Parameter[DB.BuiltInParameter.ELEM_FAMILY_PARAM].AsValueString())[:len(filter_family_name)]
            if filter_family_name == "" or family_name == filter_family_name:
                pass
            else:
                continue
            try:
                #region Преобразовать параметры
                parameters = setting[2].split("+")
                parameters = [parameter.strip() for parameter in parameters]
                #endregion
                string = ""
                main1()
                string = string [:-len(delimiter)] if len(delimiter) != 0 else string
                parameter_for_write = get_parameter_by_name(element, setting[3])
                parameter_for_write.Set(string) ## Не преобразовываю, потому что подразумеванется, что будет только строка
            except Exception as e:
                list_error_id = create_correc_dict(str(e) + "&" + create_string_from_list(setting), [element.Id.IntegerValue], list_error_id)
        f = time.time()
        time_log = create_correc_dict(setting[0].capitalize(), f-s, time_log)

    if setting[0].capitalize() == "Записать из хоста":
        s = time.time()
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
            except Exception as e:
                list_error_id = create_correc_dict(str(e) + "&" + create_string_from_list(setting), [element.Id.IntegerValue], list_error_id)
        f = time.time()
        time_log = create_correc_dict(setting[0].capitalize(), f-s, time_log)

#region Обработка ошибок
string_error = ""
string_time = ""

no_error = True
for type_error, errors_id in list_error_id.items():
    type_error_exception, type_error_config = type_error.split("&")

    string_error += type_error_exception + " " + type_error_config + "\n"
    string_error += create_string_from_list(errors_id) + "\n" + "\n"

    if type_error_exception == "The parameter is read-only.":
        pass
    elif type_error_exception == "f8ee9c2a-ae76-4bc5-a965-3787d57a3de7":
        pass
    elif type_error_exception == "":
        pass
    else:
        no_error = False

if no_error:
    print("Скрипт завершился без ошибок")
    print("-----------------------------------------------------------")
else:
    print("Есть ошибки при работе скрипта. Требуется анализ")
    print("-----------------------------------------------------------")

total_time = round(sum(time_log.values())/60, 1)
print("Общее время выполнения: " + str(total_time) + " минут")
for c, t in time_log.items():
    string_time += c +":"
    string_time += ": " + str(t)
    string_time += "\n"

# Сохранение ошибок
path_to_save = r"Z:\Отдел BIM\Марищенко Богдан\999_ЛогиСкрипта"
file_name = doc.Title + " " + str(__revit__.Application.Username) + ".txt"
file_logs = StreamWriter(path_to_save + "\\" + file_name)
file_logs.Write(string_time + "\n")
file_logs.Write(string_error)
file_logs.Close()
#endregion

#region Запись логов
try:
   obj = kpln_logs.create_obj(extStorage_guid, extStorage_name)
   kpln_logs.write_log(obj, "Запуск скрипта на спецификацию: " + dialog_out)
except:
   print("Что-то пошло не так. Лог не записался. Обратитесь в бим - отдел")
#endregion

transaction.Commit()
#endregion



