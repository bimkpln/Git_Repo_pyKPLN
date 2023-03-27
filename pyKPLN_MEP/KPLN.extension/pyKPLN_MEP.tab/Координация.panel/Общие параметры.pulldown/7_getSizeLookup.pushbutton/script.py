# -*- coding: utf-8 -*-
# Template
from Autodesk.Revit import DB
from rpw.ui import forms
import System
import codecs
from pyrevit import script

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

__title__ = 'Экспорт\nCSV'
__author__ = "@butiryc_acid" #TELEGRAM
__doc__ = '''   Программа экспортирует таблицу поиска из семейства в корректном формате

    Прим.: Для экспорта таблицы откройте необходимое семейство в редакторе семейств'''
# Проверка на то, открыто ли семейство в редакторе семейств
if doc.IsFamilyDocument:
    family_id = doc.OwnerFamily.Id
else:
    forms.Alert('Необходимо открыть семейство в редакторе семейств', title = __title__)
    script.exit()

family_size_table_manager = DB.FamilySizeTableManager.\
    GetFamilySizeTableManager(doc, family_id) # Получение менеджера таблиц

try:
    list_of_size_tables = list(family_size_table_manager.GetAllSizeTableNames())
except:
    forms.Alert('В семействе отсутствуют таблицы поиска', title = __title__)
    script.exit()

selected_size_table_name = forms.SelectFromList(
    '''Выберите csv-таблицу для экспорта''',
    list_of_size_tables
)
selected_size_table = family_size_table_manager.GetSizeTable(selected_size_table_name)

rows, columns = selected_size_table.NumberOfRows, selected_size_table.NumberOfColumns # Размер таблицы

# Получение шапки таблицы:
returned_header = ';'
if int(doc.Application.VersionNumber) <= 2020:
    for i in range(1, columns):
        _column_header = selected_size_table.GetColumnHeader(i)
        returned_header += ''.join([
            str(_column_header.Name), '##',
            str(_column_header.UnitType).replace('UT_',''), '##',
            str(_column_header.DisplayUnitType).replace('DUT_',''), ';'
            ])
elif int(doc.Application.VersionNumber) >= 2021:
    import all_units
    import all_specs
    for i in range(1, columns):
        _column_header = selected_size_table.GetColumnHeader(i)
        try:
            returned_header += ''.join([
                str(_column_header.Name), '##',
                all_specs.dickt[_column_header.GetSpecTypeId().TypeId], '##',
                all_units.dickt[_column_header.GetUnitTypeId().TypeId], ';'
            ])
        except:
            returned_header += ''.join([
                str(_column_header.Name), '##',
                "Undefined", '##',
                "UNDEFINED", ';'
            ])
returned_header = returned_header[:-1] + '\n'
returned_header = returned_header.replace('##Undefined##UNDEFINED', '##OTHER##') # Добавить реплейсы при необходимости
returned_header = returned_header.replace('DECIMAL_FEET', 'GENERAL') # Добавить реплейсы при необходимости
returned_header = returned_header.replace('Airflow', 'AIR_FLOW') # Добавить реплейсы при необходимости

# Алгоритм перебора
def search_algorithm():
    '''Функция для ускорения перебора csv - таблицы'''

    _ret = ''
    for column in range(columns):
        _ret += ''.join([selected_size_table.AsValueString(row, column).ToString(), ';'])
    _ret = _ret[:-1]
    return _ret

returned_string = returned_header
for row in range(rows):
    returned_string += ''.join([search_algorithm(),'\n'])
returned_string = returned_string[:-1]

# Запись:
def convert(path_to_file, encoding_before = 'utf-8', encoding_after = 'windows-1251'):
    '''Функция для преобразования кодировки csv - файлов. По умолчанию преобразовывает UTF-8 в ASCI
        encoding_before : str - Кодировка файла до преобразования 
        encoding_after : str - Кодировка файла после преобразования 
        return : None

    '''
    encoding_before, encoding_after = encoding_before.upper(), encoding_after.upper()
    try:
        data_string = codecs.open(path_to_file, 'r', encoding_before)
        new_content = data_string.read()
        codecs.open(path_to_file, 'w', encoding_after).write(new_content)
    except IOError as report_error:
        print("I/O error: {0}".format(report_error))

try:
    path = ''.join([forms.select_folder(), '\\', selected_size_table_name, '.csv'])
    with System.IO.StreamWriter(path) as sw:
        sw.Write(returned_string)
    convert(path)
    forms.Alert('Таблица поиска сохранена успешно', title = __title__)
except:
    forms.Alert('Не выбран путь для сохранения', title = __title__)
