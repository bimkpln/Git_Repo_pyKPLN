# -*- coding: utf-8 -*-
__title__ = "Отметка низа"
__doc__ = 'Запись высотной отметки отверстия (относительная и абсолютная)\n' \
          '    «00_Отметка_Относительная» - высота проема относительно' \
          'уровня ч.п. 1-го этажа, а точнее базовой точки проекта.\n' \
          '    «00_Отметка_Абсолютная» - высота проема относительно' \
          'ч.п. связанного уровня.\n' \


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, \
    InstanceBinding, BuiltInParameterGroup, Category, Options,\
    ParameterValueProvider, ElementId, FilterStringContains,\
    ElementParameterFilter, FilterStringRule
import math
from rpw import doc, db
from pyrevit import script
from System import Guid
import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')


# Замер производительности
def benchmark(func):
    import time

    def wrapper(*args, **kwargs):
        start = time.time()
        return_value = func(*args, **kwargs)
        end = time.time()
        print('{} Время выполнения: {} секунд.'.format(func, end-start))
        return return_value
    return wrapper


def getHeight(element):
    """Определение высоты элемента, для поиска нижней точки"""

    expand = element.LookupParameter("Расширение границ").AsDouble()
    try:
        upHeight = baseElement.LookupParameter("SYS_OFFSET_UP").AsDouble()
    except AttributeError:
        # Для старых семейств - нет линий отображения
        upHeight = 0
    if upHeight < expand:
        try:
            height = element.get_Parameter(guidHeight).AsDouble()
        except AttributeError:
            height = element.get_Parameter(guidDiam).AsDouble()
    else:
        try:
            height = element.get_Parameter(guidHeight).AsDouble() - expand
        except AttributeError:
            height = element.get_Parameter(guidDiam).AsDouble() - expand * 2
    return height


def getDescription(length_feet):
    comma = "."
    if length_feet >= 0:
        sign = "+"
    else:
        sign = "-"
    length_feet_abs = math.fabs(length_feet)
    length_meters = int(round(length_feet_abs * 304.8 / 5, 0) * 5)
    length_string = str(length_meters)
    if len(length_string) == 7:
        value = length_string[:4] + comma + length_string[4:]
    elif len(length_string) == 6:
        value = length_string[:3] + comma + length_string[3:]
    elif len(length_string) == 5:
        value = length_string[:2] + comma + length_string[2:]
    elif len(length_string) == 4:
        value = length_string[:1] + comma + length_string[1:]
    elif len(length_string) == 3:
        value = "0{}".format(comma) + length_string
    elif len(length_string) == 2:
        value = "0{}0".format(comma) + length_string
    elif len(length_string) == 1:
        value = "0{}00".format(comma) + length_string
    value = sign + value
    return value


def setPrefixByShape(element, isCircle):
    """Указание префикса и высоты элемента, в зависимости от формы"""

    if isCircle:
        prefix = "Центр на отм. "
        element_height = getHeight(element) / 2
    else:
        prefix = "Низ на отм. "
        element_height = getHeight(element)
    return prefix, element_height


def setRelativeElevHole(element, isCircle):
    """Установка относительной отметки отверстий - для ИОС - не актуально!"""

    bBox = baseElement.get_Geometry(Options()).GetBoundingBox()
    upHeight = baseElement.LookupParameter("SYS_OFFSET_UP").AsDouble()
    bBoxDownCoord_Z = (bBox.Max.Z - upHeight - bp_height)
    prefix = setPrefixByShape(element, isCircle)[0]
    element_height = setPrefixByShape(element, isCircle)[1]
    value = prefix +\
        getDescription(bBoxDownCoord_Z - element_height) +\
        " мм от ур.ч.п."
    element.LookupParameter("00_Отметка_Относительная").Set(value)


def setAbsoluteElevHole(element, isCircle):
    """Установка абсолютной отметки отверстий"""

    bBox = baseElement.get_Geometry(Options()).GetBoundingBox()
    try:
        upHeight = baseElement.LookupParameter("SYS_OFFSET_UP").AsDouble()
    except AttributeError:
        # Для старых семейств - нет линий отображения
        upHeight = 0
    bBoxDownCoord_Z = (bBox.Max.Z - upHeight - bp_height)
    prefix = setPrefixByShape(element, isCircle)[0]
    element_height = setPrefixByShape(element, isCircle)[1]
    value = prefix +\
        getDescription(bBoxDownCoord_Z - element_height) +\
        " мм от нуля здания"
    element.LookupParameter("00_Отметка_Абсолютная").Set(value)


def setAbsoluteElevShaft(element):
    """Установка абсолютной отметки шахт"""

    bBox = baseElement.get_Geometry(Options()).GetBoundingBox()
    bBoxDownCoord_Z = (bBox.Min.Z - bp_height)
    prefix = "Низ на отм. "
    value = prefix +\
        getDescription(bBoxDownCoord_Z) +\
        " мм от нуля здания"
    element.LookupParameter("00_Отметка_Абсолютная").Set(value)


collElements = list()
# Имя семейства
provider = ParameterValueProvider(ElementId(-1002002))
evaluator = FilterStringContains()
strContains = ["199_", "501_"]
for s in strContains:
    stringRule = FilterStringRule(provider, evaluator, s, False)
    trueFilter = ElementParameterFilter(stringRule)
    collElements.extend(
        FilteredElementCollector(doc).
        OfCategory(BuiltInCategory.OST_MechanicalEquipment).
        WhereElementIsNotElementType().ToElements()
    )

params = ["00_Отметка_Относительная", "00_Отметка_Абсолютная"]
params_exist = [False, False]
default_offset_bp = 0.00

# КП_Р_Высота
guidHeight = Guid("da753fe3-ecfa-465b-9a2c-02f55d0c2ff1")
# КП_Р_Ширина
guidWidth = Guid("8f2e4f93-9472-4941-a65d-0ac468fd6a5d")
# КП_Р_Диаметр
guidDiam = Guid("9b679ab7-ea2e-49ce-90ab-0549d5aa36ff")

prj_base_point = FilteredElementCollector(doc).\
                     OfCategory(BuiltInCategory.OST_ProjectBasePoint).\
                     WhereElementIsNotElementType().\
                     FirstElement()

bp_height = prj_base_point.get_BoundingBox(None).Min.Z

# ПРОВЕРКА НАЛИЧИЯ ПАРАМЕТРОВ
try:
    collElements[0].LookupParameter("00_Отметка_Абсолютная").AsString()
    collElements[0].LookupParameter("00_Отметка_Относительная").AsString()
    isUpload = True
except:
    isUpload = False

# Загрузка параметров при отсутсвии
if not isUpload:
    try:
        app = doc.Application
        catSetElements = app.Create.NewCategorySet()
        catSetElements.Insert(
            doc.
            Settings.
            Categories.
            get_Item(BuiltInCategory.OST_Windows)
        )

        catSetElements.Insert(
            doc.
            Settings.
            Categories.
            get_Item(BuiltInCategory.OST_Doors)
        )

        catSetElements.Insert(
            doc.
            Settings.
            Categories.
            get_Item(BuiltInCategory.OST_MechanicalEquipment)
        )

        originalFile = app.SharedParametersFilename
        app.SharedParametersFilename = "X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\lib\\ФОП_Scripts.txt"
        SharedParametersFile = app.OpenSharedParameterFile()
        map = doc.ParameterBindings
        it = map.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            d_Definition = it.Key
            d_Name = it.Key.Name
            d_Binding = it.Current
            d_catSet = d_Binding.Categories

            windCheck = d_catSet.Contains(
                Category.GetCategory(doc, BuiltInCategory.OST_Windows)
            )
            doorCheck = d_catSet.Contains(
                Category.GetCategory(doc, BuiltInCategory.OST_Doors)
            )
            mepCheck = d_catSet.Contains(
                Category.GetCategory(
                    doc, BuiltInCategory.OST_MechanicalEquipment
                )
            )

            for param, userBool in zip(params, params_exist):
                if d_Name == param\
                        and d_Binding.GetType() == InstanceBinding\
                        and str(d_Definition.ParameterType) == "Text"\
                        and d_Definition.VariesAcrossGroups\
                        and windCheck\
                        and doorCheck\
                        and mepCheck:
                    userBool = True
        with db.Transaction(name="КП_Высотная отметка. Добавление параметров"):
            for dg in SharedParametersFile.Groups:
                if dg.Name == "АРХИТЕКТУРА - Дополнительные":
                    for param, userBool in zip(params, params_exist):
                        if not userBool:
                            externalDefinition = dg.Definitions.get_Item(param)
                            newIB = app.Create.NewInstanceBinding(
                                catSetElements
                            )
                            doc.ParameterBindings.Insert(
                                    externalDefinition,
                                    newIB,
                                    BuiltInParameterGroup.PG_DATA
                            )

                            doc.ParameterBindings.ReInsert(
                                externalDefinition,
                                newIB,
                                BuiltInParameterGroup.PG_DATA
                            )

        map = doc.ParameterBindings
        it = map.ForwardIterator()
        it.Reset()
        with db.Transaction(name="КП_Высотная отметка. Установка параметров"):
            while it.MoveNext():
                for param in params:
                    d_Definition = it.Key
                    d_Name = it.Key.Name
                    d_Binding = it.Current
                    if d_Name == param:
                        d_Definition.SetAllowVaryBetweenGroups(doc, True)
    except Exception as e:
        print(str(e))


with db.Transaction(name="КП_Высотная отметка. Запись результатов"):
    for element in collElements:
        if element.SuperComponent is not None:
            baseElement = element.SuperComponent
            if baseElement.SuperComponent is not None:
                baseElement = baseElement.SuperComponent
        else:
            baseElement = element

        try:
            fam_name = element.Symbol.FamilyName

            # Прямугольные отверстия
            if fam_name.startswith("199_Отверстие в стене прямоугольное")\
                    or ("501_Отверстие" in fam_name
                        and "_Стена" in fam_name)\
                    or fam_name == "199_AR_OSW"\
                    or fam_name == "501_MEP_TSW":
                setAbsoluteElevHole(baseElement, False)

            # Круглые отверстия
            if fam_name.startswith("199_Отверстие в стене круглое")\
                    or fam_name == "199_AR_ORW"\
                    or fam_name == "501_MEP_TRW":
                setAbsoluteElevHole(baseElement, True)

            # Шахты
            if fam_name.startswith("501_MEP_Отв")\
                    or ("501_Отверстие" in fam_name
                        and "_Перекрытие" in fam_name)\
                    or "шахта" in fam_name.lower():
                setAbsoluteElevShaft(baseElement)

        except Exception as e:
            print("Ошибка у элемента {}. Работа прекарщена".format(element.Id))
            print(str(e))
            script.exit()
