# -*- coding: utf-8 -*-

__title__ = "Толщина\nвоздуховодов"
__author__ = 'Tima Kutsko'
__doc__ = "Расстановка толщины воздуховодов согласно СП.60 и СП.7"


import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, Element, BuiltInCategory, BuiltInParameter
from pyrevit import script
from rpw import revit, db, ui
from rpw.ui.forms import Console
import re
from System import Guid


# guidName = Guid("e6e0f5cd-3e26-485b-9342-23882b20eb43")   #"КП_О_Наименование"
# guidShortName = Guid("f194bf60-b880-4217-b793-1e0c30dda5e9")   #"КП_О_Наименование краткое"
elements = []
result = set()


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




# variables
doc = revit.doc
ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
ductFittings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctFitting).WhereElementIsNotElementType().ToElements()
elements.extend(ducts)
elements.extend(ductFittings)




# parameters of duct, fittings and insulations
try:
    parameterDict = {p.Definition.Name: p.Definition.Name for p in ducts[0].Parameters
                    if p.IsShared and p.Definition.ParameterType.ToString() == 'Length'}
    for fittingParam in ductFittings[0].Parameters:
        if fittingParam.Definition.ParameterType.ToString() == 'Length':
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



# main script
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
        except:
            print('Параметр, в который будет внесены данные, должен иметь тип данных - "Длина" и быть присвоен категориям "Воздуховоды" и "Соединительные детали воздухводов"')

        """
        try:
            for element in elements:
                lookUpSize = element.get_Parameter(BuiltInParameter.RBS_REFERENCE_FREESIZE).AsString()
                digitSize = re.findall(r'\d+', lookUpSize)						
                diamSize = re.search(r'ø\d+', lookUpSize)					
                maxItem = maxItemFinder(digitSize)
                if diamSize: 															#to find a duct fitting adapter - square to round	
                    diamDigit = float(diamSize.group(0)[1:])				
                if "ø" in lookUpSize and diamDigit == maxItem:


                    if element.GetType().ToString()  == 'Autodesk.Revit.DB.Mechanical.Duct':
                        diameter = (element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).AsDouble()) * 304.8
                        name = 'ø' + str(format(diameter, '.0f')) + ' мм'
                    else:
                        name = None					


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


                    if element.GetType().ToString()  == 'Autodesk.Revit.DB.Mechanical.Duct':						
                        width = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble()
                        height = element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble()
                        if width > height:
                            rightWidth = width * 304.8
                            rightHeight = height * 304.8
                        elif width < height:
                            rightWidth = height * 304.8
                            rightHeight = width * 304.8
                        else:
                            rightWidth = width * 304.8
                            rightHeight = height * 304.8					
                        name = str(format(rightWidth, '.0f')) + 'x' + str(format(rightHeight, '.0f')) + ' мм'
                    else:
                        name = None					


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


                if name is not None:
                    nameSystem = element.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString()			
                    #
                    if partOfSystemName in nameSystem:
                        element.get_Parameter(guidName).Set('Воздуховод из черной стали, сварной' + ' δ = ' + str(element.LookupParameter(parameterName).AsDouble() * 304.8) + ' мм, ' + name)
                        #element.get_Parameter(guidName).Set('Воздуховод из черной стали, сварной δ = 1,2 мм, ' + name)
                        element.get_Parameter(guidShortName).Set('ГОСТ 19903-74*')
                    else:
                        element.get_Parameter(guidName).Set('Воздуховод из оцинкованной стали,' + ' δ = ' + str(element.LookupParameter(parameterName).AsDouble() * 304.8) + ' мм, ' + name)
                        element.get_Parameter(guidShortName).Set('ГОСТ 14918-80*')	
                    #		

        except:
            print('Параметр, в который будет внесены данные, должен иметь тип данных - "Длина" и быть присвоен категориям "Воздуховоды" и "Соединительные детали воздухводов"')
        """	
    for r in result:
        print(r)		
except:
    ui.forms.Alert("Ошибка обработки данных. Обратись к разработчику!", title="Внимание", exit=True)