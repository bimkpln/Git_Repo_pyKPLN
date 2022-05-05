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
def Readress():
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
    # КП_И_Порядковый номер- 
    numParam = Guid("7a44592a-c698-4276-af5f-75be9b131b40")

    els = FilteredElementCollector(doc).OfCategory(ElCircuitCat).WhereElementIsNotElementType().ToElements()
    sysDict = dict()
    for element in els:
        try:
            Pref = element.get_Parameter(systemParam).AsString()
        except:
            continue
        if Pref in sysDict:
            sysDict[Pref].append(element)
        else:
            sysDict[Pref] = [element]
    components = [Label('Система для анализа:'),
                                ComboBox('combobox1', sysDict),
                                Separator(),
                                Button('Выбрать')]
    form = FlexForm('Система:', components)
    form.show()
    combobox = form.values.get('combobox1')
    newListSys = ['delete']*130

    for element in combobox:
        IndOFSys = element.get_Parameter(ElAdressParam).AsDouble()
        if IndOFSys == 999:
            continue
        else:
            try:
                index = str(IndOFSys).split('.')[0]
            except:
                index = str(IndOFSys).split(',')[0]
        newListSys[int(index)] = element
    while 'delete' in newListSys:
        newListSys.remove('delete')
    count = 0
    exCount = 0
    outList = list()
    for element in newListSys:
        IndOFSys = element.get_Parameter(ElAdressParam).AsDouble()
        ExAdress = element.get_Parameter(ExElParam).AsString()
        exEl = element.BaseEquipment
        curEls = element.Elements
        for el in curEls:
            curEl = el
        ExAdress = int(doc.GetElement(exEl.GetTypeId()).get_Parameter(AdressParam).AsDouble())
        Adress = int(doc.GetElement(curEl.GetTypeId()).get_Parameter(AdressParam).AsDouble())
        if count == 0 and ExAdress < 10:
            count += ExAdress+1
            exCount = 1
        elif count == 0 and ExAdress > 10:
            count += 1
            exCount = 0
        
        if IndOFSys% int(IndOFSys) == 0:
            outList.append([element,exCount,count])
            exCount = count
            count += Adress
        else:
            try:
                ExIndexParal = str(IndOFSys).split('.')[1]
            except:
                ExIndexParal = str(IndOFSys).split(',')[1]
            addZero = ExIndexParal+'000'
            ExIndexParal = int(addZero[:3])
            for findElement in newListSys: 
                IndOFSys = findElement.get_Parameter(ElAdressParam).AsDouble()
                if int(IndOFSys)== ExIndexParal:
                    CurIndexParal = newListSys.index(findElement)
            findedEl = outList[CurIndexParal][2]
            outList.append([element,findedEl,count])
            count += Adress
    num = 1
    flagZero = False
    for sys in outList:
        num +=1
        curSys = sys[0]
        exSys = curSys.get_Parameter(ElAdressParam).AsDouble()
        exIndex = sys[1]
        curIndex = sys[2]
        exElSys = curSys.BaseEquipment
        curElsSys = curSys.Elements
        for el in curElsSys:
            curElSys = el
        exPref = doc.GetElement(exElSys.GetTypeId()).get_Parameter(prefParam).AsString()
        curPref = doc.GetElement(curElSys.GetTypeId()).get_Parameter(prefParam).AsString()
        try:
            sysNum = (curSys.get_Parameter(systemParam).AsString()).split('_')[-1]
        except:
            ui.forms.Alert("Нет символа '_' !", title="Внимание")
            script.exit()
        if exIndex == 0:
            exMark = str(exPref)+sysNum
            exElSys.get_Parameter(sysParam).Set(exMark)
            flagZero = True
        else:
            exMark = str(exPref)+sysNum+'.'+ str (exIndex)
        curMark = str(curPref)+sysNum+'.'+ str (curIndex)
        curSys.get_Parameter(sysParam).Set(curMark)
        curSys.get_Parameter(ExElParam).Set(exMark)
        ExAdress = int(doc.GetElement(exElSys.GetTypeId()).get_Parameter(AdressParam).AsDouble())
        if exSys%int(exSys) == 0:
            if not flagZero:
                IndexSys = exIndex
            else:
                IndexSys = curIndex
        else:
            IndexSys = float(float(curIndex)+float(exIndex)/1000)
        curSys.get_Parameter(ElAdressParam).Set(IndexSys)
        curSys.get_Parameter(numParam).Set(num-1)
        curElSys.get_Parameter(sysParam).Set(curMark)
        curElSys.get_Parameter(ElAdressParam).Set(num)
        curElSys.get_Parameter(ExElParam).Set(num-1) 
    for element in combobox:
        IndOFSys = element.get_Parameter(ElAdressParam).AsDouble()
        if IndOFSys == 999:
            element.get_Parameter(ExElParam).Set(curMark)
        