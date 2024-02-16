#coding: utf-8

__title__ = "Заполнение параметров\nсистем"
__author__ = 'Tima Kutsko'
__doc__ = '''1. Заполнение параметра КП_О_Имя Системы для проектов ОВ/ВК.
Необходим при использовании вложенных семейств.\n
2. Заполнение параметра КП_О_Имя Системы для проектов ЭОМ/СС.\n
Необходим при использовании вложенных семейств.'''


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter,\
    BuiltInCategory, StorageType, FamilyInstance
from rpw import revit, db, ui
from rpw.ui.forms import CommandLink, TaskDialog
from pyrevit import script, forms
from System import Guid
from libKPLN import kpln_logs
from re import split


# classes
class CategoryOption:
    def __init__(self, name, value):
        self.name = name
        self.value = value


# definitions
def createCheckBox(catSet):
    categoties_options = [CategoryOption(c.Name, c.Id) for c in catSet]
    catCheckBoxes = forms.\
        SelectFromList.\
        show(categoties_options,
             multiselect=True,
             title='Выбери категории для записи',
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
                    sysData = current_system.\
                        get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).\
                        AsString()
                    if trueSysNames is None:
                        trueSysNames = sysData
                    elif sysData not in trueSysNames:
                        trueSysNames += '/' + sysData
    return trueSysNames


def userCheck(elements_tuple):
    global check_box
    category_name_set = set()
    for currentCat in elements_tuple:
        frstElem = FilteredElementCollector(doc).\
                OfCategory(currentCat).\
                WhereElementIsNotElementType().\
                FirstElement()
    
        if frstElem:
            category_name_set.add(frstElem.Category)

    check_box = createCheckBox(category_name_set)

    return check_box


def collTrueElem(check_box):
    elemList = list()

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


def collTrueElem_WithSubElemsFilter(elemList, onlySubFamInst=False):
    resultList = list()
    for elem in elemList:
        if isinstance(elem, FamilyInstance) and onlySubFamInst:
            if elem.SuperComponent:
                resultList.append(elem)
        else:
            if isinstance(elem, FamilyInstance):
                if not elem.SuperComponent:
                    resultList.append(elem)
            elif not onlySubFamInst:
                resultList.append(elem)

    return resultList


def paramSetter(elemCollector, onlySubFamInst=False):
    global flag
    if onlySubFamInst:
        for currentElem in elemCollector:
            kplnElemParam = currentElem.get_Parameter(guidSysName)
            if kplnElemParam is None:
                kplnElemParam = doc.\
                    GetElement(currentElem.GetTypeId()).\
                    get_Parameter(guidSysName)
            if kplnElemParam is None:
                raise Exception("У элемента нет парамтера КП_О_Имя Системы. Id:" + str(currentElem.Id))

            hostElem = currentElem.SuperComponent
            kplnHostElemParam = hostElem.get_Parameter(guidSysName)
            if kplnHostElemParam is None:
                kplnHostElemParam = doc.\
                    GetElement(hostElem.GetTypeId()).\
                    get_Parameter(guidSysName)

            if not kplnElemParam.IsReadOnly and kplnHostElemParam.HasValue:
                kplnElemParam.Set(kplnHostElemParam.AsString())

    else:
        for currentElem in elemCollector:
            try:
                sysTypeParam = currentElem.\
                    get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM)
            except AttributeError:
                sysTypeParam = currentElem.\
                    get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
                
            if sysTypeParam is not None:
                if sysTypeParam.StorageType == StorageType.String:
                    sysTypeParamData = sysTypeParam.AsString()
                elif sysTypeParam.StorageType == StorageType.ElementId:
                    sysTypeParamData = sysTypeParam.AsValueString()
                else:
                    raise Exception("Не удалось привести параметр. Отправь в BIM-отдел!")

                if sysTypeParamData is not None:
                    kplnSysParam = currentElem.get_Parameter(guidSysName)
                    if kplnSysParam is None:
                        kplnSysParam = doc.\
                            GetElement(currentElem.GetTypeId()).\
                            get_Parameter(guidSysName)

                    if kplnSysParam is None:
                        flag = True
                        output.print_md(
                            'Элемент **{} с id {}** не имеет нужных параметров!'.
                            format(
                                currentElem.Name, 
                                output.linkify(currentElem.Id))
                        )
                    elif not kplnSysParam.IsReadOnly:
                        kplnSysParam.Set(findSysType(sysTypeParamData))


#region Параметры для логирования в Extensible Storage. Не менять по ходу изменений кода
extStorage_guid = "be15305c-5249-4581-a4ca-01784efd8415"
extStorage_field_name = "Last_Run"
extStorage_name = "KPLN_SystemType"
if __shiftclick__:
    try:
       obj = kpln_logs.create_obj(extStorage_guid, extStorage_field_name, extStorage_name)
       kpln_logs.read_log(obj)
    except:
       print("Логи запуска программы отсутствуют. Плагин в этом проекте ниразу не запускался")
    script.exit()
#endregion


# input form
commands = [CommandLink('Элементы ОВ/ВК',
                        return_value='Сопоставление элементов ОВ/ВК'),
            CommandLink('Элементы ЭОМ/СС',
                        return_value='Сопоставление элементов ЭОМ/СС')]
dialog = TaskDialog('Выбери формат обработки данных',
                    title='Выбери тип проекта',
                    commands=commands,
                    buttons=['Cancel'],
                    footer='Заполнение параметров систем',
                    show_close=True)
inputRes = dialog.show()


# main code
doc = revit.doc
output = script.get_output()
system_types = list()
system_param = None
flag = False
# "КП_О_Имя Системы"
guidSysName = Guid("21213449-727b-4c1f-8f34-de7ba570973a")


if inputRes == 'Сопоставление элементов ОВ/ВК':
    # main part of code
    elemCatTuple = (BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_FlexPipeCurves,
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
                    BuiltInCategory.OST_DuctLinings,
                    BuiltInCategory.OST_GenericModel
                )
    sysTypesCatTuple = (BuiltInCategory.OST_PipingSystem,
                        BuiltInCategory.OST_DuctSystem)
    for currentCat in sysTypesCatTuple:
        system_types.extend(FilteredElementCollector(doc).
                            OfCategory(currentCat).
                            WhereElementIsElementType().
                            ToElements())

    checkBox = userCheck(elemCatTuple)
    # Прохожу по НЕ вложенным семействам
    elemCollector = collTrueElem(checkBox)
    with db.Transaction('pyKPLN_MEP: КП_О_Имя Системы для ОВ/ВК_1'):
        paramSetter(collTrueElem_WithSubElemsFilter(elemCollector, False), False)

    # Прохожу по вложенным семействам, фиксирую запуск в ExtStr
    with db.Transaction('pyKPLN_MEP: КП_О_Имя Системы для ОВ/ВК_2'):
        paramSetter(collTrueElem_WithSubElemsFilter(elemCollector, True), True)

        #region Запись логов
        try:
            obj = kpln_logs.create_obj(extStorage_guid, extStorage_field_name, extStorage_name)
            kpln_logs.write_log(obj,"Запуск скрипта заполнения 'КП_И_Имя Системы' на категории: " + "~".join([p.name for p in check_box]))
        except Exception as ex:
            flag = True
            print("Лог не записался. Обратитесь в бим - отдел: " + ex.ToString())
        #endregion

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