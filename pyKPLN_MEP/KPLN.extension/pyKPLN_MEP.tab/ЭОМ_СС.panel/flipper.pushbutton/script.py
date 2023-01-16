# -*- coding: utf-8 -*-
__title__ = 'Флиппер'
__doc__ = '''Данный скрипт переворачивает элементы категорий* относительно опорной плоскости, на которой они расположены.
Объекты могут быть выбраны заранее, либо после запуска скрипта. Выбор не обязательно отфильтровывать. Перевернуться только определенные категории*.

*Осветительные приборы, Датчики, Электрооборудование

Прим.: Подробное описание смотри на MOODLE'''
__author__ = "@butiryc_acid" #TELEGRAM

from Autodesk.Revit import DB, UI
from pyrevit import script

from AppForm import KPLN_Alarm

#region ФУНКЦИИ, КЛАССЫ, КОНСТАНТЫ:
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

CATEGORY = ["Осветительные приборы", "Датчики", "Электрооборудование"]

class _ISelectionFilter(UI.Selection.ISelectionFilter):
    '''Интерфейс RevitAPI
    Используется для фильтрации по категории

    '''
    def AllowElement(self, element):
        if element.Category.Name in CATEGORY:
            return True
        return False
#endregion

#region ВЫБОРКА ЭЛЕМЕНТОВ
element_ids = list(uidoc.Selection.GetElementIds()) # Проверка на "выбранность"
if len(element_ids) != 0:
    lighting_fixtures = [doc.GetElement(i) for i in element_ids if doc.GetElement(i).Category.Name in CATEGORY]
    if len(lighting_fixtures) == 0:
        KPLN_Alarm("В выбранных элементах нет подходящих категорий")
        script.exit()
elif len(element_ids) == 0:
    filter_ = _ISelectionFilter()
    object_type = UI.Selection.ObjectType.Element
    try:
        reference_lighting_fixtures = uidoc.Selection.PickObjects(object_type, filter_, "Выберите элементы")
    except:
        KPLN_Alarm("Выбор элементов отменен")
        script.exit()
    if len(reference_lighting_fixtures) == 0:
        KPLN_Alarm("Не выбран ни один экземпляр")
        script.exit()
    lighting_fixtures = [doc.GetElement(i) for i in reference_lighting_fixtures]
#endregion

#region ПРОВЕРКА НА ПЕРЕВОРОТ
if not all([element.CanFlipWorkPlane for element in lighting_fixtures]):
    KPLN_Alarm(("".join(["Не все элементы могут быть перевернуты"])))
    script.exit()
#endregion

#region ОСНОВНОЙ ЦИКЛ. ПЕРЕВОРОТ ОБЪЕКТОВ
transaction = DB.Transaction(doc, "".join(["pyKPLN_", __title__]))
transaction.Start()

try:
    for element in lighting_fixtures:
        element.IsWorkPlaneFlipped = not(element.IsWorkPlaneFlipped)
except:
    KPLN_Alarm("Произошла непредвиденная ошибка\nОбратитесь в BIM-Отдел")
    transaction.RollBack()
    script.exit()

transaction.Commit()
KPLN_Alarm("Элементы успешно повернуты")
#endregion