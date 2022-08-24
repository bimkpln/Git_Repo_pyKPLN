# -*- coding: utf-8 -*-


__title__ = "Задать пределы"
__doc__ = 'Определение параметров «offset_down» и «offset_up»\n' \
    'Это позволяет показать экземпляры семейств на планах '\
    'поверх стен. Визуально меняется порядок прорисовки, однако '\
    'по факту меняется длина спец. элементов экземпляров семейств.'


import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, \
    BuiltInParameter
from rpw import doc, db
from pyrevit import script
from System import Guid

levels = []
output = script.get_output()
# Ширина
guid_width = Guid("8f2e4f93-9472-4941-a65d-0ac468fd6a5d")
# Высота
guid_height = Guid("da753fe3-ecfa-465b-9a2c-02f55d0c2ff1")
# Диаметр
guid_diam = Guid("9b679ab7-ea2e-49ce-90ab-0549d5aa36ff")


def get_level_higher(max):
    closest_level = None
    for level in levels:
        if level.Elevation >= max:
            if closest_level is None:
                closest_level = level
            elif closest_level.Elevation > level.Elevation:
                closest_level = level
    return closest_level


def get_level_lower(min):
    closest_level = None
    for level in levels:
        if level.Elevation <= min:
            if closest_level is None:
                closest_level = level
            elif closest_level.Elevation < level.Elevation:
                closest_level = level
    return closest_level


for level in FilteredElementCollector(doc).\
        OfCategory(BuiltInCategory.OST_Levels).\
        WhereElementIsNotElementType().\
        ToElements():
    levels.append(level)


with db.Transaction(name="КП_Задать пределы"):
    floorHeight = 250
    levelHeight = 2750
    for elem in FilteredElementCollector(doc).\
            OfCategory(BuiltInCategory.OST_MechanicalEquipment).\
            WhereElementIsNotElementType():
        # Параметры, для записи значений
        value_up = 0
        value_down = 0

        # Основной процесс
        try:
            famName = elem.Symbol.FamilyName
            if famName.startswith("199_Отверстие")\
                    or famName.startswith("501_MEP_TRW")\
                    or famName.startswith("501_MEP_TSW"):

                try:
                    elemLevel = doc.GetElement(elem.LevelId)
                    elevParam = elem.\
                        get_Parameter(BuiltInParameter.
                                      INSTANCE_ELEVATION_PARAM).\
                        AsDouble()

                    try:
                        elemMaxElev = elemLevel.Elevation +\
                            elevParam +\
                            elem.get_Parameter(guid_height).AsDouble()
                    except AttributeError:
                        elemMaxElev = elemLevel.Elevation +\
                            elevParam +\
                            elem.get_Parameter(guid_diam).AsDouble()

                    elemMinElev = elemLevel.Elevation + elevParam

                    # SET OFFSET UP
                    elemHigherLev = get_level_higher(elemMaxElev)
                    if elemHigherLev is not None:
                        value_up = elemHigherLev.Elevation\
                            - elemMaxElev\
                            - floorHeight / 304.8

                        if value_up >= 0:
                            elem.LookupParameter("SYS_OFFSET_UP").Set(value_up)
                    else:
                        output.print_md(
                            ("У элемента с id: {} - не был определен уровень выше, поэтому значение оффсета равно 2000 мм. Отверстие может появиться на планах выше (если они есть). **Проверь вручную!**").
                            format(
                                output.linkify(elem.Id)
                            )
                        )
                        elem.LookupParameter("SYS_OFFSET_UP").Set(2000 / 304.8)

                    # SET OFFSET DOWN
                    elemLowerLev = get_level_lower(elemMinElev)
                    if elemLowerLev is not None:
                        value_down = elemMinElev - elemLowerLev.Elevation
                        if round(value_down, 5) >= 0:
                            elem.\
                                LookupParameter("SYS_OFFSET_DOWN").\
                                Set(value_down)
                    else:
                        if elevParam >= 0:
                            elem.\
                                LookupParameter("SYS_OFFSET_UP").\
                                Set(elevParam)
                        else:
                            # Добавлена обработка отверстий в примяках, которые
                            # ниже уровня подвала
                            elem.\
                                LookupParameter("SYS_OFFSET_DOWN").\
                                Set(0.0)
                            elem.\
                                LookupParameter("SYS_OFFSET_UP").\
                                Set(abs(elevParam) - floorHeight / 304.8)

                except Exception as exc:
                    output.print_md(
                        ("Ошибка {} у элемента с id: {}").
                        format(
                            exc.ToString(),
                            output.linkify(elem.Id)
                        )
                    )
        except Exception:
            pass

        # Проверка на слишком большие оффсеты
        value_down *= 304.8
        value_up *= 304.8
        if abs(value_up) >= levelHeight\
                or abs(value_down) >= levelHeight:
            output.print_md(
                ("У элемента с id: {} - оффсеты больше высоты стандартного этажа {}. Отверстие может появиться на планах выше. **Проверь вручную!**").
                format(
                    output.linkify(elem.Id),
                    levelHeight
                )
            )
