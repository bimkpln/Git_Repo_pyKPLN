# -*- coding: utf-8 -*-

__title__ = "СН17_Заполнение параметров 'КП_О_Этаж', 'ADSK_Номер здания'"
__author__ = 'Tima Kutsko'
__doc__ = "Автоматическое заполнение параметров для классификатора.\n"\
        "Заполянет параметры 'КП_О_Этаж', 'ADSK_Номер здания'\n"\
        "Актуально для всех разделов ИОС!"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import FilteredElementCollector, XYZ, Options,\
    FloorFace, BuiltInParameter, BuiltInCategory, RevitLinkInstance, Floor
from System import Guid
from rpw import revit, db, ui
from pyrevit import script, forms


class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value

    def __nonzero__(self):
        return self.state

    def __bool__(self):
        return self.state


def getTrueLvlName(element):
    try:
        centerPoint = XYZ(
            (element.Location.Curve.GetEndPoint(0).X + element.Location.Curve.GetEndPoint(1).X) / 2,
            (element.Location.Curve.GetEndPoint(0).Y + element.Location.Curve.GetEndPoint(1).Y) / 2,
            (element.Location.Curve.GetEndPoint(0).Z + element.Location.Curve.GetEndPoint(1).Z) / 2
            )
    except:
        centerPoint = element.Location.Point
    boundingBoxElement = element.get_Geometry(Options()).GetBoundingBox()
    center_boundBox_Point = XYZ(centerPoint.X, centerPoint.Y, boundingBoxElement.Min.Z)
    # arch_elements analising
    min_distance = None
    min_floor_level_name = None
    for floor in elementsARCH:
        try:
            if "жб".upper() in floor.Name.upper():
                centerPointPrj = floor.GetVerticalProjectionPoint(
                    centerPoint,
                    FloorFace.Top
                    )
                boundingBoxFloor = floor.\
                    get_Geometry(Options()).\
                    GetBoundingBox()
                transformFloorMaxPoint = transform.OfPoint(boundingBoxFloor.Max)
                transformFloorMinPoint = transform.OfPoint(boundingBoxFloor.Min)
                checkMax = transformFloorMaxPoint.X>centerPointPrj.X and transformFloorMaxPoint.Y>centerPointPrj.Y
                checkMin = transformFloorMinPoint.X<centerPointPrj.X and transformFloorMinPoint.Y<centerPointPrj.Y
                boolean = checkMax and checkMin
                if boolean:
                    if centerPointPrj.Z < boundingBoxElement.Min.Z:
                        # finded the distance between elements (mep vs arch)
                        distance = (centerPointPrj.DistanceTo(center_boundBox_Point)) * 304.8
                        if min_distance is None or distance < min_distance:
                            min_distance = distance
                            min_floor_level_name = floor.get_Parameter(BuiltInParameter.LEVEL_PARAM).AsValueString()
        except:
            pass
    return min_floor_level_name


def insulation(elements_insulation):
    errorEl = False
    for current_insulation in elements_insulation:
        current_element = doc.GetElement(current_insulation.HostElementId)
        try:
            elemGuidParam = current_element.get_Parameter(guid_level_param).AsString()
        except:
            errorEl = True
            continue
        if elemGuidParam is None:
            elemGuidParam = doc.GetElement(current_element.GetTypeId()).get_Parameter(guid_level_param).AsString()
        try:
            if not elemGuidParam is None:
                try:
                    current_insulation.get_Parameter(guid_level_param).Set(str(elemGuidParam))
                except:
                    doc.GetElement(current_insulation.GetTypeId()).get_Parameter(guid_level_param).Set(str(elemGuidParam))
        except:
            output.print_md("**Элемент, у которого нет нужных праметров:** {} - {} \
                **Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".
                format(
                    current_insulation.Name,
                    output.linkify(current_insulation.Id))
                    )
            script.exit()
        if errorEl:
            output.print_md("**Проблемный элемент:** {} \
                **Элемент будет пропущен, необходимо проверить вручную!**".
                format(output.linkify(current_insulation.Id)))
            errorEl = False


def chkBx(elements):
    # Create check boxes for elements if they have Name property.
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name) if not e.GetLinkDocument() is None]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='Выбери файл АР:',
                                                    width=600,
                                                    button_name='Запуск!')
    return elements_checkboxes


# main code
doc = revit.doc
output = script.get_output()
# get all elements of current project
category_list = [
    BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_FlexPipeCurves, BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_CableTrayFitting,
    BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_LightingDevices, BuiltInCategory.OST_NurseCallDevices,
    BuiltInCategory.OST_SecurityDevices, BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_CommunicationDevices
    ]


# КП_О_Этаж
guid_level_param = Guid("9eabf56c-a6cd-4b5c-a9d0-e9223e19ea3f")
guid_number_param = Guid("eaa57141-68d3-4f89-8272-246328f8e77b")


try:
    components = [
        ui.forms.flexform.Label("Узел ввода данных"),
        ui.forms.flexform.CheckBox(
            'all_elements',
            'Все элементы в проекте',
            default=False
            ),
        ui.forms.flexform.CheckBox(
            'all_elements_on_view',
            'Все элементы на активном виде',
            default=True),
        ui.forms.flexform.Separator(),
        ui.forms.flexform.Button("Запуск")
        ]
    form = ui.forms.FlexForm("Автоматическое заполнение параметров для классификатора", components)
    form.ShowDialog()
    all_elements = form.values["all_elements"]
    all_elements_on_view = form.values["all_elements_on_view"]
except:
    ui.forms.Alert(
        "Ошибка ввода данных. Обратись к разработчику!",
        title="Внимание",
        exit=True
        )


# selecting and analazing mep and arch elements
elements_list = list()
if all_elements:
    for current_category in category_list:
        elements_in_project = FilteredElementCollector(doc).OfCategory(current_category).WhereElementIsNotElementType().ToElements()
        elements_list.extend(elements_in_project)
elif all_elements_on_view:
    for current_category in category_list:
        elements_at_active_view = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(current_category).WhereElementIsNotElementType().ToElements()
        elements_list.extend(elements_at_active_view)
else:
    ui.form.Alert(
        "Выбери только один тип выобра элементов!",
        title="Внимание",
        exit=True
        )

# analazing links
linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
linkModels_checkboxes = chkBx(linkModelInstances)
if linkModels_checkboxes:
    elementsARCH = list()
    linkModels = [c.value for c in linkModels_checkboxes if c.state is True]
    with db.Transaction("КПЛН_Заполнение параметров по СН17"):
        for link in linkModels:
            transform = link.GetTotalTransform()
            elementsARCH = FilteredElementCollector(link.GetLinkDocument()).\
                OfClass(Floor).\
                ToElements()

            for current_item in elements_list:
                try:
                    element = doc.GetElement(current_item.id)
                except:
                    element = current_item
                try:
                    super_component = element.SuperComponent
                except:
                    super_component = None
                if super_component is None:
                    trueFloorLvlName = getTrueLvlName(element)
                    if not trueFloorLvlName is None:
                        ## get&set true level param
                        """ lvl_pat = re.compile("(\-?\d\d)\_{1}")
                        lvl_index = re.findall(lvl_pat, str(trueFloorLvlName))[0] """
                        lvl_index = trueFloorLvlName.split("_")[0]
                        index = str(lvl_index)
                        try:
                            element.get_Parameter(guid_level_param).Set(index)
                            element.get_Parameter(guid_number_param).Set("0")
                        except:
                            pass

    mep_elements_insulation = list()
    pipe_insulation = FilteredElementCollector(doc, doc.ActiveView.Id).\
        OfCategory(BuiltInCategory.OST_PipeInsulations).\
        WhereElementIsNotElementType().\
        ToElements()
    duct_insulation = FilteredElementCollector(doc, doc.ActiveView.Id).\
        OfCategory(BuiltInCategory.OST_DuctInsulations).\
        WhereElementIsNotElementType().\
        ToElements()
    mep_elements_insulation.extend(pipe_insulation)
    mep_elements_insulation.extend(duct_insulation)

    if len(mep_elements_insulation) > 0:
        with db.Transaction("КПЛН_Заполнение параметров по СН17 для изоляции"):
            insulation(mep_elements_insulation)

    ui.forms.Alert("Завершено!", title="Завершено")
