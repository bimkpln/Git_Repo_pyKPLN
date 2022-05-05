# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from System import Enum
from System.Collections.Generic import List
from rpw import revit, db, ui
from rpw.ui.forms import CommandLink, TaskDialog
from pyrevit import script, forms, HOST_APP
from pyrevit.forms import WPFWindow
from pyrevitmep.event import CustomizableEvent
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger
import clr
clr.AddReference('RevitAPI')


__doc__ = """Скрипты для замены/проверки уровней элементов ИОС.\n
    1. Сопоставление эл-в модели ИОС по уровням:\n
        \t1.1 По уровням перекрытий - каждый элемент проецируется на перекрытие из файла АР/КР. Из полученного перекрытия извлекается уровень, и присваивается элементам ИОС. Основное условие - необходимо наличие всех уровней в проекте ИОС с идентичными с файлами АР/КР именами.
        \t1.2 По выбранным уровням - выбирается уровень из проекта ИОС (из списка), затем присваивается выбранным вручную элеменатам модели ИОС.
    2. Проверка уровня на наличие элементов.\n
        \t2.1 Выбрать в модели (с разреза) - поэлементный выбор неообходимых уровней непосредственно в модели ИОС.
        \t2.2 Выбор из списка - выбор неообходимых уровней из представленного списка."""
__title__ = "Контроль уровня элементов ИОС"
__author__ = "Tima Kutsko"
__persistentengine__ = True


# definitions
# getting builincategory from element's category
def get_BuiltInCategory(category):
    if Enum.IsDefined(BuiltInCategory, category.Id.IntegerValue):
        BuiltInCat_id = category.Id.IntegerValue
        BuiltInCat = Enum.GetName(BuiltInCategory, BuiltInCat_id)
        return BuiltInCat
    else:
        return BuiltInCategory.INVALID


# creating check-boxes from class
def create_check_boxes(elements):
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name)]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='Узел ввода №2: Выбери модель из списка (нужен файл АР/КР)',
                                                    width=600,
                                                    button_name='Запуск!')
    return elements_checkboxes


# classes
class Commands:
    """
    this class need to create new commads windows
    """
    def __init__(self, dialog_title, button_1, return_value_1, button_2, return_value_2):
        self.dialog_title = dialog_title
        self.button_1 = button_1
        self.return_value_1 = return_value_1
        self.button_2 = button_2
        self.return_value_2 = return_value_2

    def task_dialog(self):
        commands = [CommandLink(self.button_1, return_value=self.return_value_1),
                    CommandLink(self.button_2, return_value=self.return_value_2)]
        dialog = TaskDialog('Выбери формат теста',
                            title=self.dialog_title,
                            commands=commands,
                            buttons=['Cancel'],
                            footer='Контроль уровня элементов ИОС',
                            show_close=True)
        return dialog.show()


class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value  


# main code
doc = revit.doc
output = script.get_output()
uidoc = revit.uidoc
logger = script.get_logger()
is_revit_2020 = int(HOST_APP.version) >= 2020
logger.debug(is_revit_2020)
category_list = [BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves, 
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
                BuiltInCategory.OST_Sprinklers, BuiltInCategory.OST_PlumbingFixtures]


link_models_instance = list()
link_models_collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
for check_link in link_models_collector:
    if check_link.GetLinkDocument():
        link_models_instance.append(check_link)

floorTypeDict = {"Перекрытия": BuiltInCategory.OST_Floors, "Фундамент несущей конструкции": BuiltInCategory.OST_StructuralFoundation}


# first input
first_input = Commands("Первичный ввод",
                    "Сопоставление элементов модели ИОС по уровням", "Сопоставление элементов",
                    "Проверка уровня на наличие элементов", "Проверка уровня"
                    ).task_dialog()


# main part of code
if first_input == 'Сопоставление элементов':
    result_levels = Commands("Выбор метода замены уровня",
                    "Сопоставление элементов модели ИОС по уровням перекрытий", "Сопоставление элементов по перекрытиям",
                    "Сопоставление элементов модели ИОС по  выбранным уровням", "Сопоставление элементов по выбранным уровням"
                    ).task_dialog()

    if result_levels == 'Сопоставление элементов по перекрытиям':
        # getting info logger about user
        log_name = "Сопоставление элементов_По перекрытиям"
        InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
        # main part of code
        # # input form 1
        components = [ui.forms.flexform.Label("Формат работы скрипта"),
                    ui.forms.flexform.CheckBox('all_elements', 'Все элементы на виде', default=True),
                    ui.forms.flexform.CheckBox('selection', 'Выбранный фрагмент', default=False),
                    ui.forms.flexform.Label("Тип перекрытия для проверки:"),
                    ui.forms.flexform.ComboBox("floorType", floorTypeDict),
                    ui.forms.flexform.Label("Введи часть имени перекрытия для проверок"),
                    ui.forms.flexform.Label("(для быстрого анализа - используй"),
                    ui.forms.flexform.Label("плиты перекрытия вместо полов):"),
                    ui.forms.flexform.TextBox("nameARCH", Text="Жб"),
                    ui.forms.flexform.Separator(),
                    ui.forms.flexform.Button("Далее")]
        form = ui.forms.FlexForm("Сопоставление элементов. Узел ввода №1", components)
        form.ShowDialog()
        selection = form.values["selection"]
        all_elements = form.values["all_elements"]
        builtInCat = form.values["floorType"]
        nameARCH = form.values["nameARCH"]

        # # selecting and analazing mep and arch elements 
        elementsARCH = list()
        link_models_checkboxe = create_check_boxes(link_models_instance)
        if link_models_checkboxe:
            link_models = [c.value for c in link_models_checkboxe if c.state is True]
            for link in link_models:
                transform = link.GetTotalTransform()
                elementsARCH.extend(FilteredElementCollector(link.GetLinkDocument()).OfCategory(builtInCat).WhereElementIsNotElementType().ToElements())
        else:
            forms.alert('Нужно выбрать хотя бы одну связанную модель!', title="Внимание", exitscript=True)
        project_levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
        project_level_names = {lvl.Name: lvl.Id for lvl in project_levels}
        selected_elements = list()
        if selection:
            selected_refrences = ui.Pick.pick_element(msg='Выбери фрагмент для анализа', multiple=True)
            for current_refrence in selected_refrences:
                element = doc.GetElement(current_refrence.id)
                if get_BuiltInCategory(element.Category) in [str(c) for c in category_list]:
                    selected_elements.append(element)
        elif all_elements:
            for current_category in category_list:
                selected_elements.extend(FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(current_category).WhereElementIsNotElementType().ToElements())
        elif selection and all_elements:
            forms.alert("Выбери только один тип выобра элементов!", title="Внимание", exitscript=True)


        # main process 1
        counter = 0
        with db.Transaction("pyKPLN_Замена уровня группы элементов"):	
            for element in selected_elements:
                try:
                    centerPoint = XYZ((element.Location.Curve.GetEndPoint(0).X + element.Location.Curve.GetEndPoint(1).X) / 2, (element.Location.Curve.GetEndPoint(0).Y
                                                        + element.Location.Curve.GetEndPoint(1).Y) / 2, (element.Location.Curve.GetEndPoint(0).Z + element.Location.Curve.GetEndPoint(1).Z) / 2)
                except:
                    centerPoint = element.Location.Point
                boundingBoxElement = element.get_Geometry(Options()).GetBoundingBox()
                center_boundBox_Point = XYZ(centerPoint.X, centerPoint.Y, boundingBoxElement.Min.Z)
                
                # # # arch_elements analising
                min_distance = None
                min_distance_project_point_floor = None
                for floor in elementsARCH:
                    try:
                        if nameARCH.upper() in floor.Name.upper():
                            centerPointPrj = floor.GetVerticalProjectionPoint(centerPoint, FloorFace.Top)
                            boundingBoxFloor = floor.get_Geometry(Options()).GetBoundingBox()
                            transformFloorMaxPoint = transform.OfPoint(boundingBoxFloor.Max)
                            transformFloorMinPoint = transform.OfPoint(boundingBoxFloor.Min)
                            checkMax = transformFloorMaxPoint.X>centerPointPrj.X and transformFloorMaxPoint.Y>centerPointPrj.Y
                            checkMin = transformFloorMinPoint.X<centerPointPrj.X and transformFloorMinPoint.Y<centerPointPrj.Y
                            boolean = checkMax and checkMin
                            if boolean:
                                if centerPointPrj.Z < boundingBoxElement.Min.Z:
                                    # # # finded the distance between elements (mep vs arch)
                                    distance = (centerPointPrj.DistanceTo(center_boundBox_Point))
                                    if min_distance is None or distance < min_distance:
                                        min_distance = distance
                                        min_distance_project_point_floor = floor
                    except:
                        pass		
                # # # getting floor level
                if not min_distance_project_point_floor is None	and not min_distance is None:
                    arch_floor_level_name = min_distance_project_point_floor.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM).AsValueString()
                    true_project_level_id = None
                    if arch_floor_level_name in project_level_names.keys():
                        for current_level_name, current_level_id in project_level_names.items():
                            if arch_floor_level_name == current_level_name:
                                true_project_level_id = current_level_id
                        if not true_project_level_id is None:
                            true_project_level_name = doc.GetElement(true_project_level_id).Name
                            if isinstance(element, FamilyInstance) and element.Host is None:
                                    current_element_level_id = element.LevelId
                                    try:
                                        current_element_level_name = doc.GetElement(current_element_level_id).Name
                                    except:
                                        output.print_md("У элемента {} с id **{}** проблемы с основой. Нужно поправить семейство".format(element.Name, output.linkify(element.Id)))	
                                    if true_project_level_name != current_element_level_name and current_element_level_id != ElementId(-1):
                                        counter += 1
                                        element_elevation = element.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble()
                                        element_level_elevation = doc.GetElement(current_element_level_id).Elevation
                                        true_level_elevation = doc.GetElement(true_project_level_id).Elevation
                                        true_elevation_element = element_elevation + element_level_elevation - true_level_elevation
                                        try:
                                            element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(true_project_level_id)
                                        except:
                                            try:
                                                element.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM).Set(true_project_level_id)
                                            except:
                                                continue
                                        element.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(true_elevation_element)
                            elif isinstance(element, MEPCurve):
                                current_element_level_name = element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsValueString()
                                if true_project_level_name != current_element_level_name:
                                    counter += 1
                                    element.ReferenceLevel = doc.GetElement(true_project_level_id)
                        else:
                            output.print_md("Не удалось найти уровень **{}** в проекте ИОС".format(arch_floor_level_name))
                    else:
                        output.print_md("Необходимо проверить наличие уровня **{}** в проекте ИОС. **Работа скрипта остановлена!**".format(arch_floor_level_name))
                        script.exit()
        
        # # output 1
        print("_" * 30)
        output.print_md("**Обработано {} элементов**".format(counter))


    elif result_levels == 'Сопоставление элементов по выбранным уровням':
        # getting info logger about user
        log_name = "Сопоставление элементов_По выбранному уровню" 
        InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
        # main part of code
        # created by Cyril Waechter
        def get_enhanced_ids():
            # Modified category for Revit 2020 and above. When you change level, object doesn't move. eg. fittings.
            enhanced_cat = (
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_ConduitFitting
                )
            return (Category.GetCategory(doc, cat).Id for cat in enhanced_cat)
        

        def get_level_from_object():
            """Ask user to select an object and retrieve its associated level"""
            try:
                ref_id = ui.Pick.pick_element("Select reference object").ElementId
                ref_object = doc.GetElement(ref_id)
                if isinstance(ref_object, MEPCurve):
                    level = ref_object.ReferenceLevel
                else:
                    level = doc.GetElement(ref_object.LevelId)
                return level
            except:
                print("Unable to retrieve reference level from this object")


        def change_level(ref_level):
            # Change reference level and relative offset for each selected object in order to change reference plane without
            # moving the object
            with db.Transaction("pyKPLN_Замена уровня элементов"):
                selection_ids = uidoc.Selection.GetElementIds()
                for id in selection_ids:
                    enhanced_cat_ids = get_enhanced_ids()
                    el = doc.GetElement(id)

                    # Seporate groups
                    if el.GroupId != ElementId(-1):
                        output.print_md("Это {} - элемент группы. **С элементами групп не работает**"
                            .format((output.linkify(el.Id))))
                        continue

                    # Change reference level of objects like ducts, pipes and cable trays
                    if isinstance(el, MEPCurve):
                        el.ReferenceLevel = ref_level

                    # Change reference level of objects like ducts, pipes and cable trays
                    elif isinstance(el, FamilyInstance):
                        # find elements with top and base levels (mep shafts)
                        try:
                            el_base_level_param = el.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
                            el_top_level_param = el.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
                            el_base_level = doc.GetElement(el_base_level_param.AsElementId())
                            el_top_level = doc.GetElement(el_top_level_param.AsElementId())
                            if not el.Category.Id in enhanced_cat_ids:
                                el_base_param_offset = el.get_Parameter(BuiltInParameter.SCHEDULE_BASE_LEVEL_OFFSET_PARAM)
                                el_top_param_offset = el.get_Parameter(BuiltInParameter.SCHEDULE_TOP_LEVEL_OFFSET_PARAM)
                                el_base_newoffset = el_base_param_offset.AsDouble() + el_base_level.Elevation - ref_level.Elevation
                                el_top_newoffset = el_top_param_offset.AsDouble() + el_top_level.Elevation - ref_level.Elevation 
                                el_base_param_offset.Set(el_base_newoffset)
                                el_top_param_offset.Set(el_top_newoffset)
                            el_base_level_param.Set(ref_level.Id)
                            el_top_level_param.Set(ref_level.Id)
                        # find elemetns with only 1 level
                        except:
                            # seporate non Host elements and Host elements
                            el_level_param = el.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
                            if el_level_param.AsValueString():
                                el_level = doc.GetElement(el_level_param.AsElementId())
                                if is_revit_2020 and not el.Category.Id in enhanced_cat_ids:
                                    el_param_offset = el.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                                    el_newoffset = el_param_offset.AsDouble() + el_level.Elevation - ref_level.Elevation
                                    el_param_offset.Set(el_newoffset)
                                elif not is_revit_2020:
                                    el_param_offset = el.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                                    el_newoffset = el_param_offset.AsDouble() + el_level.Elevation - ref_level.Elevation
                                    el_param_offset.Set(el_newoffset)
                                el_level_param.Set(ref_level.Id)
                            else:
                                el_level_param = el.\
                                    get_Parameter(
                                        BuiltInParameter.
                                        INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
                                el_level = doc.GetElement(el_level_param.AsElementId())
                                el_level_param.Set(ref_level.Id)
                    elif isinstance(el, Group):
                        output.print_md("Это {} - группа. **С группами не работает**"
                            .format((output.linkify(el.Id))))



        customizable_event = CustomizableEvent()
        class ReferenceLevelSelection(WPFWindow):
            
            def __init__(self, xaml_file_name):
                WPFWindow.__init__(self, xaml_file_name)
                self.levels = FilteredElementCollector(doc).OfClass(Level)
                self.combobox_levels.DataContext = self.levels

            # noinspection PyUnusedLocal
            def from_list_click(self, sender, e):
                level = self.combobox_levels.SelectedItem
                customizable_event.raise_event(change_level, level)

            # noinspection PyUnusedLocal
            def from_object_click(self, sender, e):
                selection = uidoc.Selection.GetElementIds()
                level = get_level_from_object()
                uidoc.Selection.SetElementIds(selection)
                customizable_event.raise_event(change_level, level)


        if __forceddebugmode__:
            selection = uidoc.Selection.GetElementIds()
            level = get_level_from_object()
            uidoc.Selection.SetElementIds(selection)
            change_level(level)
        else:
            ReferenceLevelSelection('ReferenceLevelSelection.xaml').Show()



elif first_input == 'Проверка уровня':
    # # process 2
    # # input form 2.1
    selected_levels = None
    result_2 = Commands("Проверка уровня", 
                    "Выбрать в модели (с разреза)", "Выбрать в модели (с разреза)", 
                    "Выбрать из списка", "Выбрать из списка"
                    ).task_dialog()
	
	
    if result_2 == 'Выбрать в модели (с разреза)':
        selected_levels = ui.Pick.pick_element(msg='Выбери уровни', multiple=True)
    elif result_2 == 'Выбрать из списка':
        all_model_levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
        list_checkboxe = create_check_boxes(all_model_levels)
        if list_checkboxe:
            selected_levels = [c.value for c in list_checkboxe if c.state == True]
        else:
            forms.alert('Нужно выбрать хотя бы один уровень!', title="Внимание", exitscript=True)
    else:
        script.exit()

    # # process 2
    # # input form 2.2
    result_3 = Commands("Проверка уровня", 
                    "Выдать результаты списком", "Выдать результаты списком", 
                    "Выделить элементы в модели", "Выделить элементы в модели"
                    ).task_dialog()
        
    # getting info logger about user
    log_name = "Контроль уровня_Проверка на наличее элементов" 
    InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
    # main part of code
    if not selected_levels is None:
        element_id_IList = List[ElementId]()
        for current_level in selected_levels:
            all_element_at_current_level = list()
            try:
                level = doc.GetElement(current_level.id)
            except:
                level = current_level
            if level.GetType() == Level:
                level_filter = ElementLevelFilter(level.Id)
                if get_BuiltInCategory(level.Category) == "OST_Levels":
                    for current_category in category_list:
                        collector = FilteredElementCollector(doc).OfCategory(current_category).WhereElementIsNotElementType().WherePasses(level_filter).ToElements()
                        true_collector = list()
                        for current_element in collector:
                            if current_element.SuperComponent is None:
                                true_collector.append(current_element)
                        if len(collector) == 0:
                            try:
                                curves_collector = FilteredElementCollector(doc).OfCategory(current_category).WhereElementIsNotElementType().ToElements()
                                level_name = level.Name
                                for current_element in curves_collector:
                                    current_element_level_name = current_element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsValueString()
                                    if level_name == current_element_level_name:
                                        all_element_at_current_level.append(current_element)
                                        element_id_IList.Add(current_element.Id)	
                            except:
                                pass
                        all_element_at_current_level.extend(true_collector)
                        for current_element in true_collector:
                            element_id_IList.Add(current_element.Id)
                            
                    
                # output 2.1
                if result_3 == 'Выдать результаты списком':
                    if len(all_element_at_current_level) == 0:
                        output.print_md("На уровне **{}** находится **{}** элемента(-ов)".format(level.Name, len(all_element_at_current_level)))
                    else:
                        output.print_md("На уровне **{}** находится **{}** элемента(-ов):".format(level.Name, len(all_element_at_current_level)))
                        for current_element in all_element_at_current_level:
                            output.print_md("Id элемента {}".format(output.linkify(current_element.Id)))
        
        # output 2.2
        if result_3 == 'Выделить элементы в модели':
            if element_id_IList.Count == 0:
                output.print_md("**На выбранных уровнях нет элементов**")
            else:
                uidoc.Selection.SetElementIds(element_id_IList)