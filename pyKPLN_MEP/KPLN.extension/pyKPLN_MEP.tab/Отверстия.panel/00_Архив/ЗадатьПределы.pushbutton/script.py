# -*- coding: utf-8 -*-


__title__ = "Задать пределы"
__doc__ = 'Определение параметров «offset_down» и «offset_up»\n' \
          'Это позволяет показать экземпляры семейств на планах\n'\
          'поверх стен. Визуально меняется порядок прорисовки, однако \n'\
          'по факту меняется длина спец. элементов экземпляров семейств.'


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, \
                              BuiltInParameter
from rpw import doc, db
from pyrevit import script
from System import Guid
import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')

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
    for elem in FilteredElementCollector(doc).\
            OfCategory(BuiltInCategory.OST_MechanicalEquipment).\
            WhereElementIsNotElementType():

        try:
            fam_name = elem.Symbol.FamilyName
            if fam_name.startswith("199_Отверстие в стене")\
                    or fam_name.startswith("501_MEP_TRW")\
                    or fam_name.startswith("501_MEP_TSW"):

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
                            - 250 / 304.8

                        if value_up >= 0:
                            elem.LookupParameter("SYS_OFFSET_UP").Set(value_up)
                    else:
                        output.print_md(
                            ("У элемента с id: {} - не был определен уровень выше, поэтому значение оффсета равно 1200 мм. Отверстие может появиться на планах выше (если они есть). **Проверь вручную!**").
                            format(
                                output.linkify(elem.Id)
                            )
                        )
                        elem.LookupParameter("SYS_OFFSET_UP").Set(1200 / 304.8)

                    # SET OFFSET DOWN
                    elemLowerLev = get_level_lower(elemMinElev)
                    if elemLowerLev is not None:
                        value_down = elemMinElev - elemLowerLev.Elevation
                        if value_down >= 0:
                            elem.\
                                LookupParameter("SYS_OFFSET_DOWN").\
                                Set(value_down)
                    else:
                        elem.\
                            LookupParameter("SYS_OFFSET_UP").\
                            Set(elevParam)

                except Exception as exc:
                    output.print_md(("Ошибка {} у элемента с id: {}").
                                    format(exc.ToString(),
                                           output.linkify(elem.Id)))
        except:
            pass
