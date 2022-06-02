# -*- coding: utf-8 -*-

__title__ = "Заполнение параметров для расчёта креплений"
__author__ = 'Tima Kutsko'
__doc__ = "Скрипт заполняет параметры проекта, необходимые для анализа элементов креплений через спецификации"


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, XYZ,\
    RevitLinkInstance, Options, FloorFace, BuiltInParameter
import math
from rpw import revit, db, ui
from pyrevit import script, forms


# Classes
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value

    def __nonzero__(self):
        return self.state

    def __bool__(self):
        return self.state


# Functions
def create_check_boxes_by_name(elements):
    # Create check boxes for elements if they have Name property
    elements_options = [
        CheckBoxOption(e.Name, e)
        for e in sorted(elements, key=lambda x: x.Name)
        if not e.GetLinkDocument() is None
    ]

    elements_checkboxes = forms.SelectFromList.show(
        elements_options,
        multiselect=True,
        title='Выбери подгруженные модели, из которой брать помещения:',
        width=600,
        button_name='Запуск!'
    )
    return elements_checkboxes


def GetCenterPoint(el):
    return XYZ(
        (
            el.Location.Curve.GetEndPoint(0).X +
            el.Location.Curve.GetEndPoint(1).X
        ) / 2,
        (
            el.Location.Curve.GetEndPoint(0).Y +
            el.Location.Curve.GetEndPoint(1).Y
        ) / 2,
        (
            el.Location.Curve.Origin.Z +
            el.Location.Curve.GetEndPoint(1).Z
        ) / 2
    )


def GetTrueDistance(elemPoint, elemBBox, trans, floor, minDistancePoint):
    """Возвращает расстояние до ближайшей поверхности"""
    try:
        centrPntPrj = floor.GetVerticalProjectionPoint(
            elemPoint,
            FloorFace.Bottom
        )
        floorGeom = floor.get_Geometry(Options())
        boundingBoxFloor = floorGeom.GetBoundingBox()
        transformFloorMaxPoint = trans.OfPoint(boundingBoxFloor.Max)
        transformFloorMinPoint = trans.OfPoint(boundingBoxFloor.Min)
        checkMax = transformFloorMaxPoint.X > centrPntPrj.X\
            and transformFloorMaxPoint.Y > centrPntPrj.Y
        checkMin = transformFloorMinPoint.X < centrPntPrj.X\
            and transformFloorMinPoint.Y < centrPntPrj.Y
        inArea = checkMax and checkMin

        trueFloor = None
        if inArea:
            plFacePrjPoint = None
            for geo in floorGeom:
                planarFacesList = geo.Faces
                for planarFace in planarFacesList:
                    intersectResult = planarFace.Project(elemPoint)
                    if intersectResult:
                        prjPoint = intersectResult.XYZPoint
                        if plFacePrjPoint is None\
                                or prjPoint.Z < plFacePrjPoint.Z:
                            plFacePrjPoint = prjPoint
            if plFacePrjPoint is not None\
                    and plFacePrjPoint.Z > elemBBox.Min.Z\
                    and (
                        minDistancePoint is None
                        or plFacePrjPoint.Z < minDistancePoint.Z
                    ):
                minDistancePoint = plFacePrjPoint
                trueFloor = floor
        """ if trueFloor:
            print(trueFloor.Id)
            print(elemPoint)
            print(minDistancePoint) """
    except AttributeError:
        pass

    return minDistancePoint


# main code
doc = revit.doc
output = script.get_output()
cable_tray_list = FilteredElementCollector(doc).\
    OfCategory(BuiltInCategory.OST_CableTray).\
    WhereElementIsNotElementType().\
    ToElements()


# parameters of cabel trays
try:
    parameterDict_lengh = {
        p.Definition.Name:
            p.Definition.Name
            for p in cable_tray_list[0].Parameters
            if p.Definition.ParameterType.ToString() == 'Length'
    }

    parameterDict_type = {
        p.Definition.Name:
            p.Definition.Name
            for p in cable_tray_list[0].Parameters
            if p.Definition.ParameterType.ToString() == 'Text'
    }
except Exception:
    ui.forms.Alert("В проекте нет кабельных лотков!",
                   title="Внимание",
                   exit=True)


# input form main part
try:
    components = [
        ui.forms.flexform.Label("Узел ввода данных"),
        ui.forms.flexform.CheckBox(
            'all_elements', 'Все элементы проекта',
            default=False
        ),
        ui.forms.flexform.CheckBox(
            'selection', 'Выбранный фрагмент',
            default=True
        ),
        ui.forms.flexform.Label("Выбери пар-р для присв. длины шпильки:"),
        ui.forms.flexform.ComboBox("paramName_lengh", parameterDict_lengh),
        ui.forms.flexform.Label(
            "Выбери пар-р для присв. типа лотка (положение):"
        ),
        ui.forms.flexform.ComboBox("paramName_type", parameterDict_type),
        ui.forms.flexform.Label(
            "Угол, после которого лоток - вертикальный, °:"
        ),
        ui.forms.flexform.TextBox("agnleCabTray", default="89"),
        ui.forms.flexform.Label("Введи часть имени перекрытия для проверок"),
        ui.forms.flexform.Label("(для быстрого анализа - используй"),
        ui.forms.flexform.Label("плиты перекрытия вместо полов):"),
        ui.forms.flexform.TextBox("nameARCH", Text="Жб"),
        ui.forms.flexform.Separator(),
        ui.forms.flexform.Button("Запуск")
    ]
    form = ui.forms.FlexForm(
        "Заполнение параметров для расчёта креплений ЭОМ",
        components
    )
    form.ShowDialog()

    all_elements = form.values["all_elements"]
    selection = form.values["selection"]
    paramName_lengh = form.values["paramName_lengh"]
    paramName_type = form.values["paramName_type"]
    agnleCabTray = int(form.values["agnleCabTray"])
    nameARCH = form.values["nameARCH"]
except Exception:
    ui.forms.Alert("Ошибка ввода данных. Обратись к разработчику!",
                   title="Внимание",
                   exit=True)


# input form: select links
linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
linkModels_checkboxes = create_check_boxes_by_name(linkModelInstances)


# the selecting and analazing mep elements
selectedElems = list()
if selection:
    elementsSelectReference = ui.Pick.pick_element(
        msg='Выбери фрагмент для анализа',
        multiple=True
    )
    for elem in elementsSelectReference:
        docElement = doc.GetElement(elem.id)
        if docElement.Category.Name == cable_tray_list[0].Category.Name:
            selectedElems.append(docElement)
elif all_elements:
    selectedElems = cable_tray_list
else:
    ui.forms.Alert('Не выбрано количесвто элементов для проверки', exit=True)

# the selecting and analazing arch elements
elemsARCH = dict()
if linkModels_checkboxes:
    link_models = [c.value for c in linkModels_checkboxes if c.state]
    for link in link_models:
        transform = link.GetTotalTransform()
        elemsARCH[link] = FilteredElementCollector(link.GetLinkDocument()).\
            OfCategory(BuiltInCategory.OST_Floors).\
            WhereElementIsNotElementType().\
            ToElements()

with db.Transaction("КП_Анализ лотков"):
    try:
        for elem in selectedElems:
            angleChk = math.acos(elem.CurveNormal.Z) * 180 / 3.14
            if angleChk < agnleCabTray or angleChk == agnleCabTray:

                # filter: wall holder
                if elem.\
                        LookupParameter("КП_ПП_Крепление_К стене").\
                        AsInteger() == 0:
                    # filter: find elements with sleeping param yes/no,
                    # and wake up those elements
                    elem.LookupParameter("КП_ПП_Крепление_К стене").Set(0)

                    # main part of code: find distance
                    minDistPrjPnt = None
                    centerPoint = GetCenterPoint(elem)
                    bBoxElement = elem.\
                        get_Geometry(Options()).\
                        GetBoundingBox()
                    centerBBoxPnt = XYZ(
                        centerPoint.X,
                        centerPoint.Y,
                        bBoxElement.Min.Z
                    )

                    # arch_elements analising
                    for linkKey, floorValues in elemsARCH.items():
                        transform = linkKey.GetTotalTransform()
                        for floor in floorValues:
                            if nameARCH.lower() in floor.Name.lower():
                                minDistPrjPnt = GetTrueDistance(
                                    centerBBoxPnt,
                                    bBoxElement,
                                    transform,
                                    floor,
                                    minDistPrjPnt
                                )

                    # finded the distance between elements (mep vs arch)
                    if minDistPrjPnt is not None:
                        distance = minDistPrjPnt.DistanceTo(centerBBoxPnt)
                        elem.LookupParameter(paramName_lengh).Set(distance)

                    if minDistPrjPnt is None:
                        output.print_md(
                            "Для элемента {} - расчет не прошел!".format(
                                output.linkify(elem.Id)
                                )
                            )
                    elem.LookupParameter(paramName_type).Set("ГЛ")

                else:
                    elem.LookupParameter(paramName_lengh).Set(0)

            else:
                elem.LookupParameter(paramName_type).Set("ВЛ")
    except Exception as main_exc:
        output.print_md(
            "Ошибка {}. Для элемента {} - расчета не произвелся!".
            format(
                main_exc,
                output.linkify(elem.Id)
                )
            )
