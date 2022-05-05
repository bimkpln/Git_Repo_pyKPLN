# -*- coding: utf-8 -*-

__title__ = "Проверка арматуры\nвоздуховодов"
__author__ = 'Kapustin Roman'
__doc__ = ""

from Autodesk.Revit.DB import *
import System
from System import Enum, Guid
import re
from rpw import revit, db, ui
from pyrevit import script, forms
from rpw.ui.forms import*
class DuctAccessory():
    element = 0
    hostElement = 0
    BoundMaxCord = 0
    BoundMinCord = 0
    hostBoundMaxCord = 0
    hostBoundMinCord = 0
# main code
doc = revit.doc
output = script.get_output()
categoryDuctAccessory = BuiltInCategory.OST_DuctAccessory
# catARCHWall = BuiltInCategory.OST_Walls
catARCHWallRectOpening = BuiltInCategory.OST_MechanicalEquipment
linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
linkModelsAR = dict()
for i in linkModelInstances:
    if 'AR' in i.Name.upper() or 'АР' in i.Name.upper():
        linkModelsAR[i.Name] = i

if len(linkModelsAR) < 1:
    ui.forms.Alert("В проекте нет файла АР, обратитесь в BIM отдел!", title="Ошибка")
    script.exit()
# wallInArchPrj = []
components = [Label('Выбери модель АР для анализа:'),
                            ComboBox('combobox1', linkModelsAR),
                            Separator(),
                            Button('Выбрать')]
form = FlexForm('Модели АР:', components)
form.show()
CurLink = form.values.get('combobox1')
wallRectOpeningInArchPrj = []
DuctAccessoryInprj = []

for link in FilteredElementCollector(CurLink.GetLinkDocument()).OfCategory(catARCHWallRectOpening).WhereElementIsNotElementType().ToElements():
    if 'отверстие' in link.Symbol.Family.Name.lower():
        # print(link.Symbol.Family.Name)
        wallRectOpeningInArchPrj.append(link)   
try:
    value = float(TextInput('Допуск смещения:', default="20"))/304.8
except:
    value = float(20)/304.8
flagFind = False
project_base_point_Arch =FilteredElementCollector(CurLink.GetLinkDocument()).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()
project_base_point =FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()
for base_point_Arch in project_base_point_Arch:
    base_point_Arch_Position = base_point_Arch.Position
    base_point_Arch_Shared_Position = base_point_Arch.SharedPosition
    curArchPosition = XYZ(base_point_Arch_Position.X - base_point_Arch_Shared_Position.X, base_point_Arch_Position.Y - base_point_Arch_Shared_Position.Y, base_point_Arch_Position.Z - base_point_Arch_Shared_Position.Z)
for base_point in project_base_point:
    base_point_Position = base_point.Position
    base_point_Shared_Position = base_point.SharedPosition
    curPosition = XYZ(base_point_Position.X - base_point_Shared_Position.X, base_point_Position.Y - base_point_Shared_Position.Y, base_point_Position.Z - base_point_Shared_Position.Z)
deltaX = curArchPosition.X - curPosition.X
deltaY = curArchPosition.Y - curPosition.Y
deltaZ = curArchPosition.Z - curPosition.Z
DuctAccessoryList = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(categoryDuctAccessory).WhereElementIsNotElementType().ToElements()	
elements_in_outline = []
DuctAccessoryObjList = []
ErrorList = []
FirstFindFlag = False
for ductAccessoryEl in DuctAccessoryList:
    BoundDuctAccessory = ductAccessoryEl.get_BoundingBox(doc.ActiveView)
    DuctAccessoryCenter = XYZ((BoundDuctAccessory.Max.X + BoundDuctAccessory.Min.X)/2,(BoundDuctAccessory.Max.Y + BoundDuctAccessory.Min.Y)/2,(BoundDuctAccessory.Max.Z + BoundDuctAccessory.Min.Z)/2)
    for wallRectOpening in wallRectOpeningInArchPrj:
        if flagFind:
            flagFind = False
            break
        BoundingBox = wallRectOpening.Host.get_BoundingBox(doc.ActiveView)
        BoundingBoxMax = BoundingBox.Max
        BoundingBoxMin = BoundingBox.Min
        MinCordX = BoundingBoxMin.X - deltaX
        MinCordY = BoundingBoxMin.Y - deltaY
        MinCordZ = BoundingBoxMin.Z - deltaZ
        MaxCordX = BoundingBoxMax.X - deltaX
        MaxCordY = BoundingBoxMax.Y - deltaY
        MaxCordZ = BoundingBoxMax.Z - deltaZ
        CurCordX = DuctAccessoryCenter.X
        CurCordY = DuctAccessoryCenter.Y
        CurCordZ = DuctAccessoryCenter.Z
        
        if MinCordX <= CurCordX and MaxCordX > CurCordX:
            if MinCordY <= CurCordY and MaxCordY > CurCordY:
                if MinCordZ <= CurCordZ and MaxCordZ > CurCordZ:
                    flagFind = True
                    # print(BoundingBoxMin)
                    # print(BoundingBoxMax) 
                    # # print(wallRectOpening.Id)
                    # print('---------------')
                    # print(XYZ(MinCordX,MinCordY,MinCordZ))
                    # print(XYZ(MaxCordX,MaxCordY,MaxCordZ))
                    # print(DuctAccessoryCenter)
                    
                    ductAccessory = DuctAccessory()
                    orientation = wallRectOpening.Host.Orientation
                    flag = 'Стена под углом'
                    if orientation.X == 1 or orientation.X == -1:
                        flag = 'X'
                    elif orientation.Y == 1 or orientation.Y == -1:
                        flag = 'Y'
                    BoundingBox2 = ductAccessoryEl.get_BoundingBox(doc.ActiveView)
                    ductAccessory.hostElement = wallRectOpening
                    ductAccessory.hostBoundMinCord = BoundingBox.Min
                    ductAccessory.hostBoundMaxCord = BoundingBox.Max
                    ductAccessory.element = ductAccessoryEl
                    ductAccessory.BoundMinCord = BoundingBox2.Min
                    ductAccessory.BoundMaxCord = BoundingBox2.Max
                    DuctAccessoryObjList.append(ductAccessory)
                    if flag == 'X':
                        if round((ductAccessory.hostBoundMinCord.X - ductAccessory.BoundMinCord.X),5) != 0 or round((ductAccessory.hostBoundMaxCord.X - ductAccessory.BoundMaxCord.X),5):
                            if ductAccessory.hostBoundMinCord.X - ductAccessory.BoundMinCord.X > value/2 and ductAccessory.BoundMaxCord.X - ductAccessory.hostBoundMaxCord.X > value/2:
                                if not FirstFindFlag:
                                    print("СЕРЬЕЗНЫЕ ОШИБКИ:")
                                    FirstFindFlag = True
                                output.print_md('Элемент **{} с id {}** обе стороны выходят больше, чем на 10 мм <-||->'.format(ductAccessoryEl.Name,output.linkify(ductAccessoryEl.Id)))
                                continue
                            elif ductAccessory.hostBoundMinCord.X - ductAccessory.BoundMinCord.X > value or ductAccessory.BoundMaxCord.X - ductAccessory.hostBoundMaxCord.X > value:
                                if not FirstFindFlag:
                                    print("СЕРЬЕЗНЫЕ ОШИБКИ:")
                                    FirstFindFlag = True
                                output.print_md('Элемент **{} с id {}** одна сторона выходит больше, чем на 20 мм <-||->'.format(ductAccessoryEl.Name,output.linkify(ductAccessoryEl.Id)))
                                continue
                            elif ductAccessory.hostBoundMinCord.X - ductAccessory.BoundMinCord.X < -value/2 or  ductAccessory.BoundMaxCord.X - ductAccessory.hostBoundMaxCord.X > value/2:
                                ErrorList.append(ductAccessoryEl)
                                continue
                                
                    elif flag == 'Y':
                        if round((ductAccessory.hostBoundMinCord.Y - ductAccessory.BoundMinCord.Y),5) != 0 or round((ductAccessory.hostBoundMaxCord.Y - ductAccessory.BoundMaxCord.Y),5) != 0:
                            if ductAccessory.hostBoundMinCord.Y - ductAccessory.BoundMinCord.Y > value/2 and  ductAccessory.BoundMaxCord.Y - ductAccessory.hostBoundMaxCord.Y > value/2:
                                if not FirstFindFlag:
                                    print("СЕРЬЕЗНЫЕ ОШИБКИ:")
                                    FirstFindFlag = True
                                output.print_md('Элемент **{} с id {}** обе стороны выходят больше, чем на 10 мм <-||->'.format(ductAccessoryEl.Name,output.linkify(ductAccessoryEl.Id)))
                                continue
                            elif ductAccessory.hostBoundMinCord.Y - ductAccessory.BoundMinCord.Y > value or  ductAccessory.BoundMaxCord.Y - ductAccessory.hostBoundMaxCord.Y > value:
                                if not FirstFindFlag:
                                    print("СЕРЬЕЗНЫЕ ОШИБКИ:")
                                    FirstFindFlag = True
                                output.print_md('Элемент **{} с id {}** одна сторона выходит больше, чем на 20 мм <-||->'.format(ductAccessoryEl.Name,output.linkify(ductAccessoryEl.Id)))
                                continue
                            elif ductAccessory.hostBoundMinCord.Y - ductAccessory.BoundMinCord.Y < -value/2 or ductAccessory.BoundMaxCord.Y - ductAccessory.hostBoundMaxCord.Y > value/2:
                                ErrorList.append(ductAccessoryEl)
                                continue
                    elif  flag == 'Стена под углом':
                        ErrorList.append(ductAccessoryEl)
if len(ErrorList) > 0:
    print("ОШИБКИ ДЛЯ ПРОВЕРКИ:")
    for i in ErrorList:
        output.print_md('Элемент **{} с id {}** одна стороны заходит меньше, чем на 10 мм ||<- или ->||'.format(i.Name,output.linkify(i.Id)))
        