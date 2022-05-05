# -*- coding: utf-8 -*-

__title__ = "Заполнение параметров для расчёта креплений"
__author__ = 'Tima Kutsko'
__doc__ = "Скрипт заполняет параметры проекта, необходимые для анализа элементов креплений через спецификации"


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, XYZ,\
                              RevitLinkInstance, Options, FloorFace,\
                              BuiltInParameter
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
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name) if not e.GetLinkDocument() is None]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='Выбери подгруженные модели, из которой брать помещения:',
                                                    width=600,
                                                    button_name='Запуск!')
    return elements_checkboxes


# main code
doc = revit.doc
output = script.get_output()
cable_tray_list = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements()


# parameters of cabel trays
try:
    parameterDict_lengh = {p.Definition.Name: p.Definition.Name for p in cable_tray_list[0].Parameters
                            if p.Definition.ParameterType.ToString() == 'Length'}
    parameterDict_type = {p.Definition.Name: p.Definition.Name for p in cable_tray_list[0].Parameters
                            if p.Definition.ParameterType.ToString() == 'Text'}
except:
    ui.forms.Alert("В проекте нет кабельных лотков!",
                   title="Внимание",
                   exit=True)


## input form^ main part
try:
    components = [ui.forms.flexform.Label("Узел ввода данных"),
                ui.forms.flexform.CheckBox('all_elements', 'Все элементы проекта', default=False),
                ui.forms.flexform.CheckBox('selection', 'Выбранный фрагмент', default=True),
                ui.forms.flexform.Label("Выбери пар-р для присв. длины шпильки:"),
                ui.forms.flexform.ComboBox("paramName_lengh", parameterDict_lengh),
                ui.forms.flexform.Label("Выбери пар-р для присв. типа лотка (положение):"),
                ui.forms.flexform.ComboBox("paramName_type", parameterDict_type),
                ui.forms.flexform.Label("Угол, после которого лоток - вертикальный, °:"),
                ui.forms.flexform.TextBox("agnleCabTray", default="89"),
                ui.forms.flexform.Label("Введи часть имени перекрытия для проверок"),
                ui.forms.flexform.Label("(для быстрого анализа - используй"),
                ui.forms.flexform.Label("плиты перекрытия вместо полов):"),
                ui.forms.flexform.TextBox("nameARCH", Text="Жб"),
                ui.forms.flexform.Separator(),
                ui.forms.flexform.Button("Запуск")]
    form = ui.forms.FlexForm("Заполнение параметров для расчёта креплений ЭОМ", components)
    form.ShowDialog()

    all_elements = form.values["all_elements"]
    selection = form.values["selection"]
    paramName_lengh = form.values["paramName_lengh"]
    paramName_type = form.values["paramName_type"]
    agnleCabTray = int(form.values["agnleCabTray"])
    nameARCH = form.values["nameARCH"]
except:
    ui.forms.Alert("Ошибка ввода данных. Обратись к разработчику!",
                   title="Внимание",
                   exit=True)
# input form: select links
linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
linkModels_checkboxes = create_check_boxes_by_name(linkModelInstances)

## the selecting and analazing mep elements
selected_elements = list()
if selection:
    elementsSelectReference = ui.Pick.pick_element(msg='Выбери фрагмент для анализа', multiple=True)
    for element in elementsSelectReference:
        docElement = doc.GetElement(element.id)
        if docElement.Category.Name == cable_tray_list[0].Category.Name:
            selected_elements.append(docElement)
elif all_elements:
    selected_elements = cable_tray_list
else:
    ui.forms.Alert('Не выбрано количесвто элементов для проверки', exit=True)

# the selecting and analazing arch elements
#elementsARCH = list()
elementsARCH = dict()
if linkModels_checkboxes:
    link_models = [c.value for c in linkModels_checkboxes if c.state]
    for link in link_models:
        transform = link.GetTotalTransform()
        elementsARCH[link] = FilteredElementCollector(link.GetLinkDocument()).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()

with db.Transaction("pyKPLN_Проверка высоты элементов"):
    try:
        for element in selected_elements:
            angleChk = math.acos(element.CurveNormal.Z) * 180 / 3.14
            if angleChk < agnleCabTray or angleChk == agnleCabTray:

                # filter: wall holder
                if element.LookupParameter("КП_ПП_Крепление_К стене").AsInteger() == 0:
                    # filter: find elements with sleeping param yes/no, and wake up those elements
                    element.LookupParameter("КП_ПП_Крепление_К стене").Set(0)

                    # main part of code: find distance
                    centerPoint = XYZ((element.Location.Curve.GetEndPoint(0).X + element.Location.Curve.GetEndPoint(1).X) / 2, (element.Location.Curve.GetEndPoint(0).Y
                                       + element.Location.Curve.GetEndPoint(1).Y) / 2, (element.Location.Curve.Origin.Z + element.Location.Curve.GetEndPoint(1).Z) / 2)
                    bBoxElement = element.\
                                  get_Geometry(Options()).\
                                  GetBoundingBox()
                    centerBBoxPnt = XYZ(centerPoint.X,
                                        centerPoint.Y,
                                        bBoxElement.Min.Z)

                    # arch_elements analising
                    for link_key, floor_values in elementsARCH.items():
                        minDistPrjPnt = None
                        transform = link_key.GetTotalTransform()
                        for floor in floor_values:
                            if nameARCH.lower() in floor.Name.lower():
                                try:
                                    centrPntPrj = floor.GetVerticalProjectionPoint(centerPoint, FloorFace.Top)
                                    boundingBoxFloor = floor.get_Geometry(Options()).GetBoundingBox()
                                    transformFloorMaxPoint = transform.OfPoint(boundingBoxFloor.Max)
                                    transformFloorMinPoint = transform.OfPoint(boundingBoxFloor.Min)
                                    checkMax = transformFloorMaxPoint.X>centrPntPrj.X and transformFloorMaxPoint.Y>centrPntPrj.Y
                                    checkMin = transformFloorMinPoint.X<centrPntPrj.X and transformFloorMinPoint.Y<centrPntPrj.Y
                                    boolean = checkMax and checkMin
                                    if boolean:
                                        if centrPntPrj.Z > bBoxElement.Min.Z:
                                            if minDistPrjPnt is None or centrPntPrj.Z < minDistPrjPnt.Z:
                                                minDistPrjPnt = centrPntPrj
                                                minDist_project_point_floor = floor
                                                true_link = link_key
                                except:
                                    pass
                        
                        ###finded the distance between elements (mep vs arch)
                        if not minDistPrjPnt is None:
                            floorHeigt = true_link.GetLinkDocument().GetElement(minDist_project_point_floor.Id).get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble()
                            distance = minDistPrjPnt.DistanceTo(centerBBoxPnt) - floorHeigt
                            element.LookupParameter(paramName_lengh).Set(distance)
                            element.LookupParameter(paramName_type).Set("ГЛ")

                else:
                    element.LookupParameter(paramName_lengh).Set(0)

            else:
                element.LookupParameter(paramName_type).Set("ВЛ")
    except Exception as main_exc:
        output.print_md('**ОШИБКИ:**')
        output.print_md("**Main exception is:** {} для элемента {}".
                        format(main_exc,
                               output.linkify(element.Id)))
        output.print_md('**Отправь в BIM-отдел**')
