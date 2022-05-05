# coding: utf-8

__title__ = "ЭлЦепи ЭОМ"
__author__ = 'Kapustin Roman'
__doc__ = ''''''

from rpw import *
from System import Guid
from Autodesk.Revit.DB import *
from pyrevit import script, forms
from rpw.ui.forms import*

doc = revit.doc
app = doc.Application
view = doc.ActiveView
systems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalCircuit).WhereElementIsNotElementType().ToElements()
sysDict = dict()
# for curSys in  systems:
#     pathPoints = curSys.GetCircuitPath()
#     panel = curSys.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM).AsString()
#     sysNum = curSys.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
#     sysFulName = panel + "_" + sysNum
#     sysDict[sysFulName] = pathPoints
# components = [Label('Система для анализа:'),
#                             ComboBox('combobox1', sysDict),
#                             Separator(),
#                             Button('Выбрать')]
# form = FlexForm('Система:', components)
# form.show()
# combobox = form.values.get('combobox1')
ui.forms.Alert('Выбери элемент у которого брать систему!', title="Внимение!")
while True:
    selectedReference = ui.Pick.pick_element(msg='Выбери элемент у которого брать систему"!')
    PickedEl = doc.GetElement(selectedReference.id)
    try:
        systems = PickedEl.MEPModel.ElectricalSystems
        for system in systems:
            pathPoints = system.GetCircuitPath()
            break
        break
    except:
        ui.forms.Alert('Выберан элемент без системы', title="Внимение!")
        continue
elemLines = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
lines = dict()
for line in elemLines:
    lines[line.LineStyle.Name] = line.LineStyle
components = [Label('Выбери тип линий:'),
                            ComboBox('combobox1', lines),
                            Separator(),
                            Button('Выбрать')]
form = FlexForm('Тип линий:', components)
form.show()
LineStyle = form.values.get('combobox1')
with db.Transaction("Построение цепи:"):
    for i in range(len(pathPoints)):
        lnStart =  XYZ(pathPoints[i].X,pathPoints[i].Y,0)
        try:
            lnEnd = XYZ(pathPoints[i+1].X,pathPoints[i+1].Y,0)
        except:
            break
        try:
            line = Line.CreateBound(lnStart, lnEnd)
            DetailCurve = doc.Create.  NewDetailCurve(view,line) 
            DetailCurve.LineStyle = LineStyle   
        except:
            continue
# line = Line.CreateBound(lnStart, lnEnd)
#                                 DetailCurve = doc.Create.  NewDetailCurve(view,line) 
#                                 DetailCurve.LineStyle = ActivStyle

# sysDict = dict()
# for element in els:
#     try:
#         Pref = element.get_Parameter(systemParam).AsString()
#     except:
#         continue
#     if Pref in sysDict:
#         sysDict[Pref].append(element)
#     else:
#         sysDict[Pref] = [element]
# components = [Label('Система для анализа:'),
#                             ComboBox('combobox1', sysDict),
#                             Separator(),
#                             Button('Выбрать')]
# form = FlexForm('Система:', components)
# form.show()
# combobox = form.values.get('combobox1')