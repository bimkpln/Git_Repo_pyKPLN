# -*- coding: utf-8 -*-
__title__ = '''Отметка по\nуровню'''
__author__ = '''' 'Tima Kutsko' & "@butiryc_acid" #TELEGRAM '''
__doc__ = '''Расставляет марки высотных отметок
Прим.: Убедитесь, что рабочий набор с осями и уровнями - активен'''

from Autodesk.Revit import DB
from pyrevit import script

import pyrevit.forms
from System import Guid
from System.Windows.Forms import *

from TextFormSelected import InputWindow

#region Функции и константы
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

CATEGORIES = {
    "Арматура воздуховодов": DB.BuiltInCategory.OST_DuctAccessory,
    "Арматура трубопроводов": DB.BuiltInCategory.OST_PipeAccessory,
    "Воздуховоды": DB.BuiltInCategory.OST_DuctCurves,
    "Воздухораспределитель": DB.BuiltInCategory.OST_DuctTerminal,
    "Оборудование": DB.BuiltInCategory.OST_MechanicalEquipment,
    "Сантехнические приборы": DB.BuiltInCategory.OST_PlumbingFixtures,
    "Соединительные детали воздуховодов": DB.BuiltInCategory.OST_DuctFitting,
    "Соединительные детали трубопроводов": DB.BuiltInCategory.OST_PipeFitting,
    "Трубы": DB.BuiltInCategory.OST_PipeCurves
}
CATEGORIES_NAMES = list(CATEGORIES.keys())
VIEW = doc.ActiveView
TAG_NAME = "015_Обозначение_ВысотнаяОтметка_(Об)"
FAMILY_NAME = "503_Технический_ОтметкаУровня(Об)"
LEVELS = DB.FilteredElementCollector(doc).\
            OfCategory(DB.BuiltInCategory.OST_Levels).\
            WhereElementIsNotElementType().\
            ToElements()

def horisontal_checker(element):
    '''Проверка на то, является ли кривая горизонтальной,
    либо наклонной. В дальнейшем предлагается добавить коэффициент угла уклона

    '''
    try:
        direction = element.Location.Curve.Direction
        if abs(direction[2]): # Условное место коэффициента
            return True
        else:
            return False
    except:
        return True

def location_finder(element):
    '''Функция по поиску координаты точки, в зависимости от того, является ли она криволинейным элементом, либо 
    Какой-либо другой позицией

    '''
    try:
        origin = element.Location.Curve.Origin
        origin = DB.XYZ(origin.X, origin.Y, 0)
        return origin
    except:
        origin = element.Location.Point
        origin = DB.XYZ(origin.X, origin.Y, 0)
        return origin

def find_family_by_name(name):
    '''Поиск необходимых семейств по их имени
    
    '''
    true_family = False # Без данной переменной не наботало
    list_of_family_symbol = DB.FilteredElementCollector(doc).\
        OfClass(DB.FamilySymbol).\
        ToElements()
    for family_symbol in list_of_family_symbol:
        if family_symbol.Family.Name == name:
            true_family = family_symbol
            break
    return true_family

def string_by_range(name, diapason):
    '''Функция - транслятор. С помощью диапазона преобразует строку

    '''
    returned_string = ''
    if name[diapason[0]] == "-":
        for i in diapason:
            returned_string += name[i+1]
        return '-' + returned_string
    else:
        for i in diapason:
            returned_string += name[i]
        return returned_string

def converting_level(level):
    try:
        i, f = level.split('.')
        if not '-' in i:
            i = '+' + i
        f += '0' * (3 - len(f))
        return ''.join([i, '.', f])
    except:
        if '-' in level:
            level += '0' * (3 - len(level))
            return ''.join(['-0.', level[1:]])
        else:
            level += '0' * (3 - len(level))
            return ''.join(['+0.', level])
#endregion

#region Проверка наличия семейств
if find_family_by_name(FAMILY_NAME) == False or find_family_by_name(TAG_NAME) == False:
    print(''.join(["В данном проекте отсутствуют необходимые семейства: ", TAG_NAME, " или ", FAMILY_NAME]))
    print("Вы можете добавать их по указанному пути: ")
    print("Путь к семейству '015_Обозначение_ВысотнаяОтметка_(Об)':  ")
    print(r"X:\BIM\3_Семейства\0_Общие семейства\1_Марки\015_Аннотации - аналоги системных")
    print("Путь к семейству '503_Технический_ОтметкаУровня(Об) ':  ")
    print(r"X:\BIM\3_Семейства\4_ОВиК\2_Вспомогательные семейства\01_BIM")
    script.exit()
#endregion

#region Входной узел учета уровней 
try:
    LEVELS_FILTERED = pyrevit.forms.select_levels('Выберите необходимые уровни')
    if len(LEVELS_FILTERED) == 0:
        print("Не выбран ни один уровень в проекте")
        script.exit()
except:
    script.exit()
#endregion

#region Входной узел для выделения в уровне
example_level__name = max([i.Name for i in LEVELS_FILTERED])
my_for_diapason_analisys = InputWindow(example_level__name)
Application.Run(my_for_diapason_analisys)

if my_for_diapason_analisys.dis_operation == True:
    script.exit()

if not my_for_diapason_analisys.from_level:
    temp_new_pos = int(my_for_diapason_analisys.selected_start)
    temp_end_pos = temp_new_pos + int(len(my_for_diapason_analisys.selected_text))
    DIAPASON = [int(i) for i in range(temp_new_pos, temp_end_pos)]
#endregion

#region 4. Основной алгоритм
t = DB.Transaction(doc, "pyKPLN_Создание элементов")
t.Start()

TECHNICAL_ELEMENT = find_family_by_name(FAMILY_NAME)
if not TECHNICAL_ELEMENT.IsActive:
    TECHNICAL_ELEMENT.Activate()

#region ПОИСК ЦЕНТРАЛЬНОЙ ТОЧКИ

#endregion


for level in LEVELS_FILTERED:

    # Блок создания интерсектора
    level_elevation = level.Parameter[DB.BuiltInParameter.LEVEL_ELEV].AsDouble()
    outline_intersector = DB.Outline(
        DB.XYZ(-1000, -1000, level_elevation),
        DB.XYZ(1000, 1000, level_elevation)
        )
    bounding_box_filter = DB.BoundingBoxIntersectsFilter(outline_intersector)
    collector = DB.FilteredElementCollector(doc, VIEW.Id).\
        WhereElementIsNotElementType().\
        WherePasses(bounding_box_filter).\
        ToElements()

    # Фильтрация, создание и маркировка. Основной цикл
    structural_type = DB.Structure.StructuralType.NonStructural
    level_elevation = level.Parameter[DB.BuiltInParameter.LEVEL_ELEV].AsDouble() * 304.8 / 1000 # Конвертируем только тут
    level_elevation = converting_level(str(level_elevation))

    for element in collector:
        if element.Category.Name in CATEGORIES_NAMES and horisontal_checker(element):
            #Добавляем новый элемент
            new_element = doc.Create.NewFamilyInstance(
                location_finder(element),
                TECHNICAL_ELEMENT,
                level,
                structural_type
            )

            #region ПЕРЕДАЧА ПАРАМЕТРОВ
            #КП_И_Отметка от уровня
            FLOOR_HEIGHT = Guid("9bf9520b-59aa-4c4d-a52e-5d8209041792")
            new_element.get_Parameter(FLOOR_HEIGHT).Set(level_elevation)

            #КП_О_Имя Системы
            SYSTEM_NAME = Guid("21213449-727b-4c1f-8f34-de7ba570973a")
            try:
                if element.get_Parameter(SYSTEM_NAME).AsString() != None:
                    new_system = element.get_Parameter(SYSTEM_NAME).AsString()
                else:
                    new_system = "Нет системы"
                new_element.get_Parameter(SYSTEM_NAME).Set(new_system)
            except:
                pass

            #КП_И_Имя уровня
            LEVEL_NAME = Guid("d8bcada1-fade-45af-87c5-508119731db4")
            if not my_for_diapason_analisys.from_level:
                new_element.get_Parameter(LEVEL_NAME).Set(string_by_range(level.Name, DIAPASON))
            else:
                FLOOR_PARAM = Guid("9eabf56c-a6cd-4b5c-a9d0-e9223e19ea3f")
                try:
                    temp = level.get_Parameter(FLOOR_PARAM).AsString()
                    new_element.get_Parameter(LEVEL_NAME).Set(temp)
                except:
                    new_element.get_Parameter(LEVEL_NAME).Set("Параметр не заполнен")
            #endregion

            # Добавить наполнением параметрами
            try:
                tag = DB.IndependentTag.Create(
                    doc,
                    VIEW.Id,
                    DB.Reference(new_element),
                    False,
                    DB.TagMode.TM_ADDBY_CATEGORY,
                    DB.TagOrientation.Horizontal,
                    DB.XYZ(
                        location_finder(element)[0],
                        location_finder(element)[1],
                        level.Parameter[DB.BuiltInParameter.LEVEL_ELEV].AsDouble()
                        )
                )
                tag.ChangeTypeId(find_family_by_name(TAG_NAME).Id)
            except:
                print("Данный вид не предназначен для маркирования")
            break
t.Commit()
#endregion
