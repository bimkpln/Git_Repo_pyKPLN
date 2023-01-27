#coding: utf-8

__title__ = "Заполнение параметров\nсистем"
__author__ = 'Tima Kutsko'
__doc__ = '''1. Заполнение параметра КП_О_Имя Системы для проектов ОВ/ВК.
Необходим при использовании вложенных семейств.\n
2. Заполнение параметра КП_О_Имя Системы для проектов ЭОМ/СС.\n
Необходим при использовании вложенных семейств.'''


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter,\
                              BuiltInCategory
from rpw import revit, db, ui
from rpw.ui.forms import CommandLink, TaskDialog
from pyrevit import script, forms
from System import Guid
from re import split


# classes
class CategoryOption:
    def __init__(self, name, value):
        self.name = name
        self.value = value


# definitions
def createCheckBoxe(catSet):
    categoties_options = [CategoryOption(c.Name, c.Id) for c in catSet]
    catCheckBoxes = forms.\
        SelectFromList.\
        show(categoties_options,
             multiselect=True,
             title='Выбери элементы для прикрепления',
             width=500,
             button_name='Выбрать'
             )
    return catCheckBoxes


def findSysType(sysTypeData):
    trueSysNames = None
    sysTypeData = split(r',', sysTypeData)
    if sysTypeData[0] == '':
        trueSysNames = 'Нет системы'
    else:
        for current_system in system_types:
            systemAbbreav = current_system.\
                            get_Parameter(BuiltInParameter.
                                          RBS_SYSTEM_ABBREVIATION_PARAM).\
                            AsString()
            for currentSysTypeData in sysTypeData:
                if systemAbbreav == currentSysTypeData:
                    if trueSysNames is None:
                        trueSysNames = str(current_system.
                                           get_Parameter(BuiltInParameter.
                                                         ALL_MODEL_TYPE_NAME
                                                         ).
                                           AsString())
                    else:
                        trueSysNames += '/' + str(
                            current_system.
                            get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).
                            AsString())
    return trueSysNames


def collTrueElem(elements_tuple):
    category_name_set = set()
    for currentCat in elements_tuple:
        frstElem = FilteredElementCollector(doc).\
                   OfCategory(currentCat).\
                   WhereElementIsNotElementType().\
                   FirstElement()
        if frstElem:
            category_name_set.add(frstElem.Category)
    check_box = createCheckBoxe(category_name_set)
    if check_box:
        selCatIds = [p.value for p in check_box]
        for currId in selCatIds:
            elemList.extend(FilteredElementCollector(doc).
                            OfCategoryId(currId).
                            WhereElementIsNotElementType().
                            ToElements())
    else:
        script.exit()
    return elemList


# input form
commands = [CommandLink('Элементы ОВ/ВК',
                        return_value='Сопоставление элементов ОВ/ВК'),
            CommandLink('Элементы ЭОМ/СС',
                        return_value='Сопоставление элементов ЭОМ/СС')]
dialog = TaskDialog('Выбери формат теста',
                    title='Выбери тип проекта',
                    commands=commands,
                    buttons=['Cancel'],
                    footer='Заполнение параметров систем',
                    show_close=True)
inputRes = dialog.show()


# main code
doc = revit.doc
output = script.get_output()
elemList = list()
system_types = list()
system_param = None
flag = False
# "КП_О_Имя Системы"
guidSysName = Guid("21213449-727b-4c1f-8f34-de7ba570973a")


if inputRes == 'Сопоставление элементов ОВ/ВК':
    # main part of code
    elemCatTuple = (BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_PipeAccessory,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_PlumbingFixtures,
                    BuiltInCategory.OST_PipeInsulations,
                    BuiltInCategory.OST_DuctInsulations,
                    BuiltInCategory.OST_DuctLinings
                )
    sysTypesCatTuple = (BuiltInCategory.OST_PipingSystem,
                        BuiltInCategory.OST_DuctSystem)
    for currentCat in sysTypesCatTuple:
        system_types.extend(FilteredElementCollector(doc).
                            OfCategory(currentCat).
                            WhereElementIsElementType().
                            ToElements())

    with db.Transaction('pyKPLN_MEP: Заполнение параметра КП_О_Имя Системы для ОВ/ВК'):
        elemCollector = collTrueElem(elemCatTuple)
        for currentElem in elemCollector:
            #false parameter (КП_О_Позиция) checking
            # if not currentElem.get_Parameter(guidSysName):
            #     guidSysName = Guid("ae8ff999-1f22-4ed7-ad33-61503d85f0f4")  # "КП_О_Позиция"
            #after checking part
            try:
                #RBS_SYSTEM_NAME_PARAM
                #RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM
                sysTypeParamData = currentElem.\
                                   get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).\
                                   AsString()
                if sysTypeParamData is not None:
                    trueSysNames = findSysType(sysTypeParamData)
                    try:
                        currentElem.\
                            get_Parameter(guidSysName).\
                            Set(trueSysNames)
                    except:
                        try:
                            doc.GetElement(currentElem.GetTypeId()).\
                                get_Parameter(guidSysName).\
                                Set(trueSysNames)
                        except Exception as exc:
                            output.print_md("Ошибка {} в элементе {}".
                                            format(str(exc),
                                                   output.linkify(currentElem.
                                                                  Id))
                                            )
            except AttributeError:
                system_type_param = currentElem.\
                                    get_Parameter(BuiltInParameter.
                                                  RBS_DUCT_SYSTEM_TYPE_PARAM)
                if system_type_param:
                    sysTypeParamData = system_type_param.AsValueString()
                    try:
                        try:
                            currentElem.\
                                        get_Parameter(guidSysName).\
                                        Set(sysTypeParamData)
                        except:
                            doc.\
                                GetElement(currentElem.GetTypeId()).\
                                get_Parameter(guidSysName).\
                                Set(sysTypeParamData)
                    except:
                        flag = True
                        output.print_md('Элемент **{} с id {}** не имеет нужных параметров!'.
                                        format(currentElem.Name,
                                               output.linkify(currentElem.Id))
                                        )
    if flag:
        ui.forms.Alert('Завершено с ошибками! Ознакомся в появившемся окне',
                       title='Заполнение параметра КП_О_Имя Системы')
    else:
        ui.forms.Alert('Завершено!',
                       title='Заполнение параметра КП_О_Имя Системы')


elif inputRes == 'Сопоставление элементов ЭОМ/СС':
    # main part of code
    elemCatTuple = (BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_Conduit,
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_CableTrayFitting,
                    BuiltInCategory.OST_ConduitFitting,
                    BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_ElectricalFixtures,
                    BuiltInCategory.OST_DataDevices,
                    BuiltInCategory.OST_LightingDevices,
                    BuiltInCategory.OST_NurseCallDevices,
                    BuiltInCategory.OST_SecurityDevices,
                    BuiltInCategory.OST_FireAlarmDevices,
                    BuiltInCategory.OST_CommunicationDevices)

    with db.Transaction('pyKPLN_MEP: Заполнение параметра КП_О_Имя Системы для ЭОМ/СС'):
        elemCollector = collTrueElem(elemCatTuple)
        for currentElem in elemCollector:
            workset_value = currentElem.\
                get_Parameter(
                    BuiltInParameter.
                    ELEM_PARTITION_PARAM).\
                AsValueString()

            trueValue = None
            trueValueList = workset_value.split("ЭОМ_")
            if len(trueValueList) == 2:
                trueValue = trueValueList[1]
            else:
                trueValueList = workset_value.split("СС_")
                if len(trueValueList) == 2:
                    trueValue = trueValueList[1]

            if trueValue is None:
                trueValue = workset_value

            try:
                try:
                    currentElem.\
                        get_Parameter(guidSysName).\
                        Set(trueValue)
                except:
                    doc.\
                        GetElement(currentElem.GetTypeId()).\
                        get_Parameter(guidSysName).\
                        Set(trueValue)
            except:
                output.print_md('Элемент **{} с id {}** не имеет нужных параметров!'.
                                format(currentElem.Name,
                                       output.linkify(currentElem.Id)))
        if flag:
            ui.forms.Alert('Завершено с ошибками! Ознакомся в появившемся окне',
                           title='Заполнение параметра КП_О_Имя Системы')
        else:
            ui.forms.Alert('Завершено!',
                           title='Заполнение параметра КП_О_Имя Системы')