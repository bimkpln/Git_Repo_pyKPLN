# -*- coding: utf-8 -*-
__title__ = "Толщина\nвоздуховодов"
__author__ = 'Tima Kutsko'
__doc__ = "Расстановка толщины воздуховодов согласно СП.60 и СП.7"


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter
from pyrevit import script
from rpw import revit, db, ui
from libKPLN import kpln_logs
import re


# definition - Find max item
def maxItemFinder(digitSize):
    sizeList = []
    for s in digitSize:
        sizeList.append(float(s))
    return max(sizeList)


def duct_depth_vent(element):
    lookUpSize = element.get_Parameter(BuiltInParameter.RBS_REFERENCE_FREESIZE).AsString()
    digitSize = re.findall(r'\d+', lookUpSize)
    diamSize = re.search(r'ø\d+', lookUpSize)
    maxItem = maxItemFinder(digitSize)
    # to find a duct fitting adapter - square to round
    if diamSize:
        diamDigit = float(diamSize.group(0)[1:])
    if "ø" in lookUpSize and diamDigit == maxItem:
        if maxItem < 250:
            element.LookupParameter(parameterName).Set(0.5 / 304.8)
        elif maxItem >= 250 and maxItem < 500:
            element.LookupParameter(parameterName).Set(0.6 / 304.8)
        elif maxItem >= 500 and maxItem < 900:
            element.LookupParameter(parameterName).Set(0.7 / 304.8)
        elif maxItem >= 900 and maxItem < 1400:
            element.LookupParameter(parameterName).Set(1.0 / 304.8)
        elif maxItem >= 1400 and maxItem < 1800:
            element.LookupParameter(parameterName).Set(1.2 / 304.8)
        elif maxItem >= 1800:
            element.LookupParameter(parameterName).Set(1.4 / 304.8)
        try:
            nameInsul = element.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE).AsString()		
            if partOfInsulationName in nameInsul:	
                if maxItem < 900:
                    element.LookupParameter(parameterName).Set(0.8 / 304.8)
                elif maxItem >= 900 and maxItem < 1400:
                    element.LookupParameter(parameterName).Set(1.0 / 304.8)
                elif maxItem >= 1400 and maxItem < 1800:
                    element.LookupParameter(parameterName).Set(1.2 / 304.8)
                elif maxItem >= 1800:
                    element.LookupParameter(parameterName).Set(1.4 / 304.8)
        except:
            pass
    else:
        if maxItem < 300:
            element.LookupParameter(parameterName).Set(0.5 / 304.8)
        elif maxItem >= 300 and maxItem < 1250:
            element.LookupParameter(parameterName).Set(0.7 / 304.8)
        elif maxItem >= 1250:
            element.LookupParameter(parameterName).Set(0.9 / 304.8)
        try:
            nameInsul = element.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_TYPE).AsString()
            if partOfInsulationName in nameInsul:
                if maxItem < 1250:
                    element.LookupParameter(parameterName).Set(0.8 / 304.8)
                elif maxItem >= 1250:
                    element.LookupParameter(parameterName).Set(0.9 / 304.8)
        except:
            pass
    if element.LookupParameter(parameterName).AsDouble() * 304.8 > 0:
        result.add(element.get_Parameter(BuiltInParameter.RBS_REFERENCE_FREESIZE).AsString() + ' δ = ' + str(element.LookupParameter(parameterName).AsDouble() * 304.8) + ' мм')			
    return result


def duct_depth_smoke(element):
    lookUpSize = element.get_Parameter(BuiltInParameter.RBS_REFERENCE_FREESIZE).AsString()
    digitSize = re.findall(r'\d+', lookUpSize)
    diamSize = re.search(r'ø\d+', lookUpSize)
    maxItem = maxItemFinder(digitSize)
    #to find a duct fitting adapter - square to round
    if diamSize:
        diamDigit = float(diamSize.group(0)[1:])
    if "ø" in lookUpSize and diamDigit == maxItem:
        if maxItem < 900:
            element.LookupParameter(parameterName).Set(0.8 / 304.8)
        elif maxItem >= 900 and maxItem < 1400:
            element.LookupParameter(parameterName).Set(1.0 / 304.8)
        elif maxItem >= 1400 and maxItem < 1800:
            element.LookupParameter(parameterName).Set(1.2 / 304.8)
        elif maxItem >= 1800:
            element.LookupParameter(parameterName).Set(1.4 / 304.8)	
    else:	
        if maxItem < 1250:
            element.LookupParameter(parameterName).Set(0.8 / 304.8)
        elif maxItem >= 1250:
            element.LookupParameter(parameterName).Set(0.9 / 304.8)
    if element.LookupParameter(parameterName).AsDouble() * 304.8 > 0:
        result.add(element.get_Parameter(BuiltInParameter.RBS_REFERENCE_FREESIZE).AsString() + ' δ = ' + str(element.LookupParameter(parameterName).AsDouble() * 304.8) + ' мм')
    return result


#region Параметры для логирования в Extensible Storage. Не менять по ходу изменений кода
extStorage_guid = "753380C4-DF00-40F8-9745-D53F328AC139"
extStorage_field_name = "Last_Run"
extStorage_name = "KPLN_DuctSize"
if __shiftclick__:
    try:
       obj = kpln_logs.create_obj(extStorage_guid, extStorage_field_name, extStorage_name)
       kpln_logs.read_log(obj)
    except:
       print("Логи запуска программы отсутствуют. Плагин в этом проекте ниразу не запускался")
    script.exit()
#endregion


# variables
doc = revit.doc
elements = []
result = set()
ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
ductFittings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
elements.extend(ducts)
elements.extend(ductFittings)


# parameters of duct, fittings and insulations
try:
    parameterDict = {p.Definition.Name: p.Definition.Name for p in ducts[0].Parameters
                    if p.IsShared and p.Definition.ParameterGroup.ToString() == 'PG_GEOMETRY'}
    for fittingParam in ductFittings[0].Parameters:
        if fittingParam.Definition.ParameterGroup.ToString() == 'PG_GEOMETRY':
            parameterDict[fittingParam.Definition.Name] = fittingParam.Definition.Name
except:
    print("В проекте нет воздуховодов!")
    script.exit()



# form
components = [ui.forms.flexform.Label("Выбери общий параметр для присвоения толщины"),
            ui.forms.flexform.Label("(должен быть присвоен воздуховодам"),
            ui.forms.flexform.Label("и соед. деталям):"),
            ui.forms.flexform.ComboBox("paramName", parameterDict),
            ui.forms.flexform.Label("Введи индивидуальную часть имени огнезащ. материала:"),
            ui.forms.flexform.TextBox("partOfInsulationName", Text="EI"),
            ui.forms.flexform.Label("Введи индивидуальную часть имени системы"),
            ui.forms.flexform.Label("(зависит от заданного сокращения для системы)"),
            ui.forms.flexform.Label("противодымных систем:"),
            ui.forms.flexform.TextBox("partOfSystemName", Text="Противодым."),
            ui.forms.flexform.Separator(),
            ui.forms.flexform.Button("Запуск")]
form = ui.forms.FlexForm("Толщина стали воздуховодов", components)
form.ShowDialog()



#region Запись логов
try:
    #main code
    parameterName = form.values["paramName"]
    partOfInsulationName = form.values["partOfInsulationName"]
    partOfSystemName = form.values["partOfSystemName"].lower()


    with db.Transaction('pyKPLN_Толщина воздуховодов'):
        try:
            for element in elements:
                nameSystem = element.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString().lower()
                if partOfSystemName in nameSystem:
                    duct_depth_smoke(element)
                else:
                    duct_depth_vent(element)
        
            #region Запись логов
            try:
                obj = kpln_logs.create_obj(extStorage_guid, extStorage_field_name, extStorage_name)
                kpln_logs.write_log(obj, "Запуск скрипта заполнения толщины воздуховодов")
            except Exception as ex:
                print("Лог не записался. Обратитесь в бим - отдел: " + ex.ToString())
            #endregion
        except:
            print('Параметр, в который будет внесены данные, должен иметь тип данных - "Длина" и быть присвоен категориям "Воздуховоды" и "Соединительные детали воздухводов"')
        

    for r in result:
        print(r)
except:
    ui.forms.Alert("Ошибка обработки данных. Обратись к разработчику!", title="Внимание", exit=True)
#endregion
