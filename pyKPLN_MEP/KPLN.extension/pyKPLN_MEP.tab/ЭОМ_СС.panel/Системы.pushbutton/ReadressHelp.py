# coding: utf-8

__title__ = "Редактирование системы"
__author__ = 'Kapustin Roman'
__doc__ = ''''''

from rpw import *
from System import Guid
from libKPLN.MEP_Elements import allMEP_Elements
from System.Windows.Forms import Form, Button, ComboBox, Label
from System.Drawing import Point, Size
import webbrowser
import clr
import math
from Autodesk.Revit.DB import *
import re
from pyrevit import script, forms
from rpw.ui.forms import*
from Autodesk.Revit.ApplicationServices import *
from pyrevit.forms import WPFWindow
from System.Windows import Window
from pyrevit.framework import wpf
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.DB import MEPSystem
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
# clr.AddReference('ProtoGeometry')
# clr.AddReference("RevitServices")
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
flag = False
textToWrite = "System"
doc = revit.doc
app = doc.Application
uidoc = revit.uidoc
output = script.get_output()
view = doc.ActiveView
lnStart = XYZ(0,0,0)
output = script.get_output()
def ReadressHepl():
    elemCatTuple = [BuiltInCategory.OST_MechanicalEquipment,
                        BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                        BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_CableTrayFitting,
                        BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_LightingFixtures,
                        BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_DataDevices,
                        BuiltInCategory.OST_LightingDevices, BuiltInCategory.OST_NurseCallDevices,
                        BuiltInCategory.OST_SecurityDevices, BuiltInCategory.OST_FireAlarmDevices,
                        BuiltInCategory.OST_CommunicationDevices]
    ElCircuitCat = BuiltInCategory.OST_ElectricalCircuit

    sysParam = Guid("98b30303-a930-449c-b6c2-11604a0479cb")
    # КП_И_Адрес текущий
    sysParam = Guid("686ecb99-8191-425a-9cd1-65d3597c3f75")
    # КП_О_Позиция
    prefParam = Guid("ae8ff999-1f22-4ed7-ad33-61503d85f0f4")
    # RBZ_Количество занимаемых адресов
    AdressParam = Guid("1fd68ef8-37fe-4896-bf08-c54bf8751ff6")
    # КП_И_Адрес предыдущий
    ExElParam = Guid("07230eb2-e78d-4b00-9e54-3e4f55bbf21b")
    # КП_И_Префикс системы
    systemParam = Guid("bd9e5583-305b-495f-8702-de38d6e01b7e")
    # КП_И_Адрес устройства- 
    ElAdressParam = Guid("02beca6d-e93e-47aa-9cd7-cd5f2867145e")

    els = FilteredElementCollector(doc).OfCategory(ElCircuitCat).WhereElementIsNotElementType().ToElements()
    sysDict = []
    for element in els:
        try:
            Pref = element.get_Parameter(systemParam).AsString()
        except:
            continue
        if Pref not in sysDict:
            sysDict.append(Pref)
    components = [Label('Система для анализа:'),
                                ComboBox('combobox1', sysDict),
                                Separator(),
                                Button('Выбрать')]
    form = FlexForm('Система:', components)
    form.show()
    combobox = form.values.get('combobox1')
    while True:
        value = TextInput('Выберите стартовый номер', default="1")
        try:
            index = int(value)
            break
        except:
            ui.forms.Alert('Введено не число!', title="Внимение!")
    while True:
        selectedReference = ui.Pick.pick_element(msg='Выберите первый элемент цепи. По окончанию - нажми "Esc"!')
        element = doc.GetElement(selectedReference.id)
        if 'Марки'.lower() not in element.Category.Name.lower() and str(element.Category.Name.lower()) != 'оси' and str(element.Category.Name.lower()) != 'связанные файлы'  and str(element.Category.Name.lower()) != 'линии':
            break
    sysList = [1,1]
    for iteration in range(128):
        index += 1
        curSys = element.MEPModel.AssignedElectricalSystems
        if not curSys and iteration > 0:
            break
        sysList = []
        for i in curSys:
            sysList.append(i)
        sysList[0].get_Parameter(ElAdressParam).Set(index)
        sysList[0].get_Parameter(systemParam).Set(combobox)
        elements = sysList[0].Elements
        for el in elements:
            element = el
        