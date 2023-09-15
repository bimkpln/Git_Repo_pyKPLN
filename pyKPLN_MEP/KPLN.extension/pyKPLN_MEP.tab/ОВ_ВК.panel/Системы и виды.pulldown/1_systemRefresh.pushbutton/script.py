# coding: utf-8

__title__ = "Обновить\nимя системы"
__author__ = 'Tima Kutsko'
__doc__ = '''1. Сопоставляет сокращение для системы с типом системы.\n
2. Обновляет системы у элементов ОВ/ВК'''


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter,\
                              BuiltInCategory, Plumbing, Mechanical,\
                              MEPSystemClassification, FilterStringRule,\
                              FilterStringContains, ParameterValueProvider,\
                              ElementParameterFilter, ElementId

from rpw import revit, db
from pyrevit import script


# definitions
def benchmark(func):
    import time

    def wrapper(*args, **kwargs):
        start = time.time()
        return_value = func(*args, **kwargs)
        end = time.time()
        print('{} Время выполнения: {} секунд.'.format(func, end-start))
        return return_value
    return wrapper


def getColl(builtInCat, sysColl):
    for sys in sysColl:
        sysName = sys.Name
        provider = ParameterValueProvider(ElementId(BuiltInParameter.
                                                    RBS_SYSTEM_NAME_PARAM))
        evaluator = FilterStringContains()
        stringRule = FilterStringRule(provider, evaluator, sysName, False)
        trueFilter = ElementParameterFilter(stringRule)
        fstSysEl = FilteredElementCollector(doc).\
                   OfCategory(builtInCat).\
                   WherePasses(trueFilter).\
                   WhereElementIsNotElementType().\
                   FirstElement()
        if fstSysEl:
            fstSysElList.append(fstSysEl)


def mSysCr(classification, name):
    mNewSys = Mechanical.MechanicalSystemType.Create(doc, classification, name)
    return mNewSys


def pSysCr(classification, name):
    pNewSys = Plumbing.PipingSystemType.Create(doc, classification, name)
    return pNewSys


def sysChanger(collector):
    for currentElement in collector:
        try:
            trueMEPSys = currentElement.MEPSystem
            trueSysTypeId = trueMEPSys.GetTypeId()
            for transSys in transSysList:
                try:
                    transSystemId = transSys.Id
                    currentElement.SetSystemType(transSystemId)
                    currentElement.SetSystemType(trueSysTypeId)
                    break
                except Exception as exc:
                    if "is not valid HVAC system type" not in str(exc)\
                            and "internal error" not in str(exc)\
                            and "is not valid piping system" not in str(exc):
                        output.print_md("Отправь разработчику: **{} - {}**".
                                        format(exc, currentElement.Id))
        except Exception as exc:
            output.print_md("В элементе {} - ошибка {}".
                            format(output.linkify(currentElement.Id), exc))


# main code
doc = revit.doc
output = script.get_output()
systemTypes = list()


# getting first element of each system
fstSysElList = list()
mechSysColl = FilteredElementCollector(doc).\
              OfClass(Mechanical.MechanicalSystem)
pipeSysColl = FilteredElementCollector(doc).\
              OfClass(Plumbing.PipingSystem)
getColl(BuiltInCategory.OST_DuctCurves, mechSysColl)
getColl(BuiltInCategory.OST_PipeCurves, pipeSysColl)
# getting duct and pipe systems
systemTypesCategoryTuple = (BuiltInCategory.OST_PipingSystem,
                            BuiltInCategory.OST_DuctSystem)
for currentCategory in systemTypesCategoryTuple:
    systemTypes.extend(FilteredElementCollector(doc).
                       OfCategory(currentCategory).
                       WhereElementIsElementType().
                       ToElements())


# find false names systems abbreviation, and generate try abbreviation.
# secondly - create clear systems
with db.Transaction('pyKPLN_MEP: 1. Корректировка сокр. для систем'):
    # working with systems abbreviation
    externalSysNames = list()
    for currentSysType in systemTypes:
        sysTypeAbbr = currentSysType.Abbreviation
        typeSymbols = [s for s in sysTypeAbbr]
        try:
            if typeSymbols[-1] != "_":
                externalSysNames.append(currentSysType.
                                        get_Parameter(BuiltInParameter.
                                                      ALL_MODEL_TYPE_NAME).
                                        AsString())
                externalName = sysTypeAbbr + "_"
                currentSysType.Abbreviation = externalName
        except Exception:
            continue
    # create clear systems
    suppAir = mSysCr(MEPSystemClassification.SupplyAir, "BIM_suppAir")
    exhAir = mSysCr(MEPSystemClassification.ExhaustAir, "BIM_exhAir")
    suppHidr = pSysCr(MEPSystemClassification.SupplyHydronic, "BIM_suppHidr")
    retHidr = pSysCr(MEPSystemClassification.ReturnHydronic, "BIM_retHidr")
    sanitary = pSysCr(MEPSystemClassification.Sanitary, "BIM_sanitary")
    domCW = pSysCr(MEPSystemClassification.DomesticColdWater, "BIM_domCW")
    domHW = pSysCr(MEPSystemClassification.DomesticHotWater, "BIM_domHW")
    transSysList = [suppAir, exhAir, suppHidr, retHidr, sanitary, domCW, domHW]


# working with true systems
with db.Transaction('pyKPLN_MEP: 2. Обновление имени системы'):
    # working with systems abbreviation
    for currentSysType in systemTypes:
        typeName = currentSysType.get_Parameter(BuiltInParameter.
                                                ALL_MODEL_TYPE_NAME).\
                                                AsString()
        sysTypeAbbr = currentSysType.Abbreviation
        firstTruePartSysAbbr = ''
        firstTruePartSysAbbrList = None
        lastTruePartSysAbbr = None
        try:
            sysSplitter = len(sysTypeAbbr.split('_'))
            if sysSplitter > 1 and sysSplitter < 3:
                firstTruePartSysAbbrList = sysTypeAbbr.split('_')[:-1]
                lastTruePartSysAbbr = sysTypeAbbr.split('_')[-1]
            elif sysSplitter > 2:
                firstTruePartSysAbbrList = sysTypeAbbr.split('_')[:-2]
                lastTruePartSysAbbr = sysTypeAbbr.split('_')[-2]
            for atr in firstTruePartSysAbbrList:
                firstTruePartSysAbbr += atr
                firstTruePartSysAbbr += "_"
            if lastTruePartSysAbbr != typeName and sysSplitter > 2:
                trueSysTypeAbbr = firstTruePartSysAbbr + typeName + '_'
                currentSysType.Abbreviation = trueSysTypeAbbr
            elif lastTruePartSysAbbr != typeName and sysSplitter > 1:
                trueSysTypeAbbr = typeName + '_'
                currentSysType.Abbreviation = trueSysTypeAbbr
        except Exception as exc:
            if "is not defined" not in str(exc)\
                    and "iteration over non-sequence of type NoneType"\
                    not in str(exc)\
                    and "index out of range" not in str(exc):
                output.print_md("Отправь меня разработчику: {}".format(exc))
    # working with system types in document
    sysChanger(fstSysElList)
    for transSys in transSysList:
        try:
            doc.Delete(transSys.Id)
        except Exception as exc:
            print(exc)
