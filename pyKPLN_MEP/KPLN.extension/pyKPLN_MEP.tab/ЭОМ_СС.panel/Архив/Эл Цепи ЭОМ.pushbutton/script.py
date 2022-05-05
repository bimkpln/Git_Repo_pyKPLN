# coding: utf-8

__title__ = "ЭлЦепи ЭОМ"
__author__ = 'Kapustin Roman'
__doc__ = ''''''

from rpw import *
from System import Guid
from Autodesk.Revit.DB import *
from pyrevit import script, forms
from rpw.ui.forms import*
from Autodesk.Revit.DB.Structure import StructuralType
doc = revit.doc
app = doc.Application
view = doc.ActiveView
WorkPlane = doc.ActiveView.GenLevel
systems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalCircuit).WhereElementIsNotElementType().ToElements()
PanelParam = Guid('d6ce883f-81d1-4d81-a4e7-fb897c10944c')
SysNumParam = Guid('85c1b5f7-8b83-4533-8fc4-c19f72de752b')
IdParam = Guid('2aa97f34-a1fc-489b-9371-ee666075e4ee')
typeCabelYGOParam = Guid('bd1cc219-f982-4da5-80aa-76145e6b814c')
allFam = FilteredElementCollector(doc).OfClass(Family).ToElements()
cabelFam = None
for fam in allFam:
    if "002_Кабель ЭОМ" in fam.Name:
        ids = fam.GetFamilySymbolIds()
        for i in ids:
            cabelFam =  doc.GetElement(i)
            break
        break
if not cabelFam:
    ui.forms.Alert('Семейство кабеля не подгружено в проет!', title="Внимение!")
    script.exit()
# sysDict = dict()
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
            PanelName = system.PanelName
            CircuitNumber = system.CircuitNumber
            pathPoints = system.GetCircuitPath()
            SysID = system.Id
            break
        break
    except:
        ui.forms.Alert('Выберан элемент без системы', title="Внимение!")
        continue
# elemLines = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
# lines = dict()
# for line in elemLines:
#     lines[line.LineStyle.Name] = line.LineStyle
# components = [Label('Выбери тип линий:'),
#                             ComboBox('combobox1', lines),
#                             Separator(),
#                             Button('Выбрать')]
# form = FlexForm('Тип линий:', components)
# form.show()
# LineStyle = form.values.get('combobox1')
commands= [CommandLink('Без типа прокладки', return_value='NoType'),
            CommandLink('Прокладка в гофротрубе', return_value='Gofr')]
dialog = TaskDialog('Выберите тип прокладки:',
               commands=commands,
                show_close=True)
dialog_out = dialog.show()



with db.Transaction("Построение цепи:"):
    # selectedReference = ui.Pick.pick_element(msg='Выбери элемент у которого брать систему"!')
    # PickedEl = doc.GetElement(selectedReference.id).Symbol
    for i in range(len(pathPoints)):
        lnStart =  XYZ(pathPoints[i].X,pathPoints[i].Y,0)
        try:
            lnEnd = XYZ(pathPoints[i+1].X,pathPoints[i+1].Y,0)
        except:
            break
        try:
            line = Line.CreateBound(lnStart, lnEnd)
            newobj = doc.Create.NewFamilyInstance(line,cabelFam,WorkPlane,StructuralType.NonStructural)   
            newobj.get_Parameter(PanelParam).Set(str(PanelName))
            newobj.get_Parameter(SysNumParam).Set(str(CircuitNumber))
            newobj.get_Parameter(IdParam).Set(str(SysID))
            if dialog_out == 'NoType':
                newobj.get_Parameter(typeCabelYGOParam).Set(0)
            if dialog_out == 'Gofr':
                newobj.get_Parameter(typeCabelYGOParam).Set(1)
        except:
            continue







# selectedReference = ui.Pick.pick_element(msg='Выбери элемент у которого брать систему"!')
#     PickedEl = doc.GetElement(selectedReference.id).Symbol
#     point =  XYZ(pathPoints[0].X,pathPoints[0].Y,0)
#     point2=  XYZ(pathPoints[1].X,pathPoints[1].Y,0)
#     line = Line.CreateBound(point, point2)
#     newobj = doc.Create.NewFamilyInstance(line,PickedEl,WorkPlane,StructuralType.NonStructural)
#     print(point)
#     # for i in range(len(pathPoints)):
#     #     lnStart =  XYZ(pathPoints[i].X,pathPoints[i].Y,0)
#     #     try:
#     #         lnEnd = XYZ(pathPoints[i+1].X,pathPoints[i+1].Y,0)
#     #     except:
#     #         break
#     #     try:
#     #         line = Line.CreateBound(lnStart, lnEnd)
#     #         DetailCurve = doc.Create.NewDetailCurve(view,line) 
#     #         DetailCurve.LineStyle = LineStyle   
#     #     except:
#     #         continue












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