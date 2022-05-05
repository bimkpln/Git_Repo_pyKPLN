# -*- coding: utf-8 -*-

__title__ = "INGD_Секция"
__author__ = 'Kapustin Roman'
__doc__ =""
from Autodesk.Revit.DB import *
import System
from System import Enum, Guid
import re
from rpw import revit, db, ui
from pyrevit import script, forms
from rpw.ui.forms import*
# main code
doc = revit.doc
output = script.get_output()
# ABS_Участок
guid_section_param = Guid("97cb0bc4-eab8-40aa-b94e-7917ebad01ea")
# Имя оси 
guid_grid_name = Guid("1cacba6e-69a7-4fb5-a10b-613ea01c1188")
guid_grid_korp = Guid("a70049a2-734c-4e6b-8c9f-dacb183efd90")
# get all elements of current project
categoryListArch = [BuiltInCategory.OST_Walls, BuiltInCategory.OST_Ceilings, BuiltInCategory.OST_Rooms, 
                BuiltInCategory.OST_Windows, BuiltInCategory.OST_Floors, BuiltInCategory.OST_StairsRailing,
                BuiltInCategory.OST_Roofs, BuiltInCategory.OST_Doors, BuiltInCategory.OST_CurtainWallPanels, 
                BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_GenericModel,BuiltInCategory.OST_PlumbingFixtures]
categoryListEOS = [BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_FlexPipeCurves, BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_DataDevices,
                BuiltInCategory.OST_LightingDevices, BuiltInCategory.OST_NurseCallDevices,
                BuiltInCategory.OST_SecurityDevices, BuiltInCategory.OST_FireAlarmDevices,
                BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_CommunicationDevices,
                BuiltInCategory.OST_PipeInsulations,BuiltInCategory.OST_DuctInsulations]
categoryListKR = [BuiltInCategory.OST_StructuralFraming,BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_GenericModel,BuiltInCategory.OST_Floors,BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_StructuralFoundation,BuiltInCategory.OST_MechanicalEquipment,BuiltInCategory.OST_Windows]
catARCHgrid = BuiltInCategory.OST_Grids
catARCHlvl = BuiltInCategory.OST_Levels
commands= [CommandLink('ИОС', return_value='IOS'),CommandLink('АР', return_value='AR'),
                CommandLink('КР', return_value='KR')]
dialog = TaskDialog('Выберите раздел',
                commands=commands,
                show_close=True)
dialog_out = dialog.show()
if dialog_out == 'IOS':
    categoryList = categoryListEOS
elif dialog_out == 'KR':
    categoryList = categoryListKR
elif dialog_out == 'AR':
    categoryList = categoryListArch
else:
    script.exit()
# _________________________________ОБЪЯВЛЕНИЕ КЛАССОВ________________________________________________________________________

class Box:
    Korpus = None
    Sect = str()
    Level = str()
    ZCord = float()
    ZUpperCord = float()
    MinPointX= float()
    MinPointY= float()
    MaxPointX= float()
    MaxPointY= float()
    MinXYPoint = [MinPointX,MinPointY]
    MaxXYPoint = [MaxPointX,MaxPointY]
#________________________Функция проверки линий на пересечения __________________________________________________________________              
def slope(P1, P2):
    if (P2[0] - P1[0]) == 0:
        return 0
    else:
        return(P2[1] - P1[1]) / (P2[0] - P1[0])
def y_intercept(P1, slope):
    return P1[1] - slope * P1[0]
def line_intersect(m1, b1, m2, b2):
    if m1 == m2:
        # print ("These lines are parallel!!!")
        return None
    x = (b2 - b1) / (m1 - m2)
    y = m1 * x + b1
    return x,y

linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
linkModelsCord = [i for i in linkModelInstances if 'Разб'.upper() in i.Name.upper()]
if len(linkModelsCord) < 1:
    ui.forms.Alert("В проекте нет разбивочного файла, обратитесь в BIM отдел!", title="Ошибка")
    script.exit()
elif len(linkModelsCord) > 1:
    ui.forms.Alert("В проекте больше 1 разбивочного файла, обратитесь в BIM отдел!", title="Ошибка")
    script.exit()
else:
    linkModelsCord = linkModelsCord[0]
try:
    lvlsInRpoject = FilteredElementCollector(linkModelsCord.GetLinkDocument()).OfCategory(catARCHlvl).WhereElementIsNotElementType().ToElements()
except:
    ui.forms.Alert("В проекте нет разбивочного файла!", title="Ошибка")
    script.exit()
gridsInRpoject = FilteredElementCollector(linkModelsCord.GetLinkDocument()).OfCategory(catARCHgrid).WhereElementIsNotElementType().ToElements()
project_base_point_Arch =FilteredElementCollector(linkModelsCord.GetLinkDocument()).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
project_Shared_point_Arch = FilteredElementCollector(linkModelsCord.GetLinkDocument()).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()
for Shared_point_Arch in project_Shared_point_Arch:
    boundBoxSharedPointArch = XYZ(round(Shared_point_Arch.get_BoundingBox(doc.ActiveView).Max.X, 7),round(Shared_point_Arch.get_BoundingBox(doc.ActiveView).Max.Y, 7),round(Shared_point_Arch.get_BoundingBox(doc.ActiveView).Max.Z, 7))
try:
    for base_point_Arch in project_base_point_Arch:
        boundBoxPoint = XYZ(round(base_point_Arch.get_BoundingBox(doc.ActiveView).Max.X, 7),round(base_point_Arch.get_BoundingBox(doc.ActiveView).Max.Y, 7),round(base_point_Arch.get_BoundingBox(doc.ActiveView).Max.Z, 7))
except:
    boundBoxPoint = XYZ(round(project_base_point_Arch.get_BoundingBox(doc.ActiveView).Max.X, 7),round(project_base_point_Arch.get_BoundingBox(doc.ActiveView).Max.Y, 7),round(project_base_point_Arch.get_BoundingBox(doc.ActiveView).Max.Z, 7))

# 1._____________________формируется словарь из осей разбивочного в формате gridDICT[корпус][секция] = [mim,max] - где min,max - min,max координаты баундин бокса 
fourGridList = []
fourGridColisionList = []
SectBoxObjList = []
gridsKList = dict()
gridDICT = dict()

for gridKorp in gridsInRpoject:
    if gridKorp.Name.upper() == 'А' or gridKorp.Name.upper() == 'A':
            AGridCordArch = gridKorp.GetExtents().MinimumPoint
    if gridKorp.Name.upper() == '1':
        OneGridCordArch = gridKorp.GetExtents().MinimumPoint
    try:
        gridKVal = gridKorp.get_Parameter(guid_grid_korp).AsString()
        if gridKVal:
            if gridKVal not in gridsKList:
                gridsKList[gridKVal] = [gridKorp]
                gridDICT[gridKVal] = []
            else:
                gridsKList[gridKVal].append(gridKorp)
            
    except:
        continue
try:
    if len(gridDICT.keys()) == 0:
        gridsKList['no K'] = gridsInRpoject
except:
    a = 1

gridDICTHelp = dict()
for gridKListKey in gridsKList.keys():
    for gridKListVal in gridsKList[gridKListKey]:
        try:
            grid_name = gridKListVal.get_Parameter(guid_grid_name).AsString()
        except:
            ui.forms.Alert("Не настроен разбивочный файл, обратитесь в BIM отдел! (Нет параметра)!", title="Ошибка!")
            script.exit()
        if grid_name:
            if '-' in grid_name:
                grid_nameSplited = grid_name.split('-')
                for grid_nameSplitedVal in grid_nameSplited:
                    if grid_nameSplitedVal not in gridDICTHelp:
                        gridDICTHelp[grid_nameSplitedVal] = [gridKListVal]
                    else:
                        gridDICTHelp[grid_nameSplitedVal].append(gridKListVal)
            else:
                if grid_name not in gridDICTHelp:
                    gridDICTHelp[grid_name] = [gridKListVal]
                else:
                    gridDICTHelp[grid_name].append(gridKListVal)
    gridDICT[gridKListKey] = gridDICTHelp
    gridDICTHelp = dict()

# for i in gridDICT:
#     print(i)
#     for j in gridDICT[i]:
#         print(j)
#         for m in gridDICT[i][j]:
#             print(m.Name)
MaxMinList = []
for  i in gridDICT:
    # print(i)
    for j in gridDICT[i]:
        # print(j)
        for FirtOfFoutGrig in gridDICT[i][j]:
            for SecondOfFoutGrig in gridDICT[i][j]:
                A1 = [FirtOfFoutGrig.Curve.Tessellate()[0].X,FirtOfFoutGrig.Curve.Tessellate()[0].Y]
                A2 = [FirtOfFoutGrig.Curve.Tessellate()[1].X,FirtOfFoutGrig.Curve.Tessellate()[1].Y]
                B1 = [SecondOfFoutGrig.Curve.Tessellate()[0].X,SecondOfFoutGrig.Curve.Tessellate()[0].Y]
                B2 = [SecondOfFoutGrig.Curve.Tessellate()[1].X,SecondOfFoutGrig.Curve.Tessellate()[1].Y]
                if A1[0] < A2[0]: 
                    slope_A = slope(A1, A2)
                else:
                    slope_A = slope(A2, A1)
                if B1[0] < B2[0]: 
                    slope_B = slope(B1, B2)
                else:
                    slope_B = slope(B2, B1)
                
                if slope_A == 0:
                    continue
                if slope_B == 0:
                    continue
                y_int_A = y_intercept(A1, slope_A)
                y_int_B = y_intercept(B1, slope_B)
                lineIntersection = line_intersect(slope_A, y_int_A, slope_B, y_int_B)
                if lineIntersection:
                    if lineIntersection[0] > 1000 or lineIntersection[1] > 1000 or lineIntersection[0] < -1000 or lineIntersection[1] < -1000:
                        continue
                else:
                    continue
                if list(lineIntersection) not in fourGridColisionList:
                    fourGridColisionList.append(list(lineIntersection))
# ,key=lambda i: i[1]
        MaxMinList.append([min((fourGridColisionList)),max((fourGridColisionList))])
        # print([min(list(fourGridColisionList)),max(list(fourGridColisionList))])

        fourGridColisionList = []  
iter = 0
for  i in gridDICT:
    # print(i)
    for j in gridDICT[i]:
        # print(j)
        gridDICT[i][j] = MaxMinList[iter]
        iter+=1
        # print(gridDICT[i][j])
# for  i in gridDICT:
    
#     for j in gridDICT[i]:
#         print(j)
#         print(gridDICT[i][j])
DeltaX = 0
DeltaY = 0
DeltaZ = 0
gridsInEngRpoject = FilteredElementCollector(doc).OfCategory(catARCHgrid).WhereElementIsNotElementType().ToElements()
project_base_point_Enj =FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
project_Shared_point_Enj = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()
for Shared_point_Enj in project_Shared_point_Enj:
    boundBoxSharedPointEnj = XYZ(round(Shared_point_Enj.get_BoundingBox(doc.ActiveView).Max.X, 7),round(Shared_point_Enj.get_BoundingBox(doc.ActiveView).Max.Y, 7),round(Shared_point_Enj.get_BoundingBox(doc.ActiveView).Max.Z, 7))
for base_point_Enj in project_base_point_Enj:
    boundBoxPointEnj = XYZ(round(base_point_Enj.get_BoundingBox(doc.ActiveView).Max.X, 7),round(base_point_Enj.get_BoundingBox(doc.ActiveView).Max.Y, 7),round(base_point_Enj.get_BoundingBox(doc.ActiveView).Max.Z, 7))
# print(str(boundBoxSharedPointEnj))
# print(str(boundBoxSharedPointArch))
if str(boundBoxPointEnj) == str(boundBoxPoint):
    for gridInEngPrj in gridsInEngRpoject:
        DeltaY = boundBoxSharedPointArch.Y - boundBoxSharedPointEnj.Y
        DeltaX = boundBoxSharedPointArch.X - boundBoxSharedPointEnj.X
        DeltaZ = boundBoxSharedPointArch.Z - boundBoxSharedPointEnj.Z
else:
    ui.forms.Alert("Проблемы с координатами, обратитесь в BIM отдел!", title="Ошибка!")
    script.exit()


errorCount = 0
OutOfModelElList = []
OutOfModelAutoSectLvlFindFlag = False
errorEl = []
with db.Transaction("СИТИ_Заполнение пар-ров секции"):
    for curCat in categoryList:
        elementsList = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(curCat).WhereElementIsNotElementType().ToElements()
        for element in elementsList:
            try:
                Filter = element.get_Parameter(guid_section_param).AsString()
            except:
                ui.forms.Alert("У элемента нет параметров, обратитесь в BIM- отдел", title="Ошибка")
                print(element.Id)
                script.exit()
            try:
                superComponent = element.SuperComponent
            except:
                superComponent = None
            if superComponent is None:
                OutOfModelFlag = True
                try:
                    centerPoint = (element.GetSpatialElementCalculationPoint().X + DeltaX,element.GetSpatialElementCalculationPoint().Y + DeltaY)
                except:
                    try:
                        centerPoint = (((element.BoundingBox[doc.ActiveView].Max.X + element.BoundingBox[doc.ActiveView].Min.X) / 2)+DeltaX, ((element.BoundingBox[doc.ActiveView].Max.Y
                                    + element.BoundingBox[doc.ActiveView].Min.Y) / 2)+DeltaY)
                    except:
                        try:
                            centerPoint = (element.Location.Point.X + DeltaX,element.Location.Point.Y)
                        except:
                            output.print_md('Элемент **с id {}** левый! Проверьте по id!'.format(output.linkify(element.Id)))
                            errorEl.append(element)
                for  i in gridDICT:
                    for j in gridDICT[i]:
                        MinCordX = gridDICT[i][j][0][0]
                        MinCordY = gridDICT[i][j][0][1]
                        MaxCordX = gridDICT[i][j][1][0]
                        MaxCordY = gridDICT[i][j][1][1]
                        CurCordX = float(centerPoint[0])
                        CurCordY = float(centerPoint[1])
                        if MinCordX <= CurCordX and MaxCordX > CurCordX:
                            if MinCordY <= CurCordY and MaxCordY > CurCordY:                 
                                if len(j) == 1:
                                    out = '0' + j
                                else:
                                    out = j
                                try:
                                    element.get_Parameter(guid_section_param).Set(out)
                                except:
                                    ui.forms.Alert("У элемента нет параметров, обратитесь в BIM- отдел", title="Ошибка")
                                    script.exit()
                                OutOfModelFlag = False
                                SysFlag = False
                                break
                if OutOfModelFlag:
                    OutOfModelElList.append(element)
                    # output.print_md('Элемент **{} с id {}** вне модели! Проверьте вручную!'.format(element.Name,output.linkify(element.Id)))
                    errorCount +=1
                    OutOfModelFlag = True
    flagOut = False
    if errorCount > 0:
        # commands= [CommandLink('ДА', return_value='да'),CommandLink('НЕТ', return_value='нет')]   
        # dialog = TaskDialog('В проекте есть элементы вне осей.\nОпределить положение аналитически?', commands=commands)
        # dialog_out = dialog.show()
        flagOut = True
        OutOfModelAutoSectLvlFindFlag = True
        # if dialog_out == 'да':
        #     OutOfModelAutoSectLvlFindFlag = True
        # elif dialog_out == 'нет':
        #     OutOfModelAutoSectLvlFindFlag = False
        # else:
        #     script.exit()  
    if errorCount == 0:
        ui.forms.Alert("Завершено!", title="Завершено")
    else:
        # print('Количество ошибок - {} шт.'.format(errorCount))
        # if dialog_out == 'да':
        #     print('_________________________________________________________________________')
        #     print('Результаты аналитического поиска этажа и секции:')  
        #_______________________________________________________АНАЛИТИЧЕСКИЙ ПОИСК СЕКЦИИ И ЭТАЖА ДЛЯ ЭЛ ВНЕ МОДЕЛИ___________________________________________________________________ 
        if OutOfModelAutoSectLvlFindFlag:
            for element in OutOfModelElList:   
                ElFoungSectLvlFlag = False
                breakFlag = False
                for dopusk in [0.15,0.3,1,2,5,10]:
                    
                    if breakFlag:
                        # print(dopusk)
                        break
                    for delta in [[0,dopusk],[0,-dopusk],[dopusk,0],[-dopusk,0],[dopusk,dopusk],[-dopusk,-dopusk],[dopusk,-dopusk],[-dopusk,dopusk]]:
                        if breakFlag:
                        # print(dopusk)
                            break
                        try:
                            superComponent = element.SuperComponent
                        except:
                            superComponent = None
                        if superComponent is None and not ElFoungSectLvlFlag:
                            OutOfModelFlag = True
                            try:
                                centerPoint = (element.GetSpatialElementCalculationPoint().X + DeltaX + delta[0],element.GetSpatialElementCalculationPoint().Y + delta[1] + DeltaY)
                            except:
                                try:
                                    centerPoint = (((element.BoundingBox[doc.ActiveView].Max.X + element.BoundingBox[doc.ActiveView].Min.X) / 2)+DeltaX + delta[0], ((element.BoundingBox[doc.ActiveView].Max.Y
                                                + element.BoundingBox[doc.ActiveView].Min.Y) / 2)+DeltaY + delta[1])
                                except:
                                    centerPoint = (element.Location.Point.X + DeltaX + delta[0],element.Location.Point.Y + DeltaY + delta[1])

                            for  i in gridDICT:
                                for j in gridDICT[i]:
                                    MinCordX = float(gridDICT[i][j][0][0])
                                    MinCordY = float(gridDICT[i][j][0][1])
                                    MaxCordX = float(gridDICT[i][j][1][0])
                                    MaxCordY = float(gridDICT[i][j][1][1])
                                    CurCordX = float(centerPoint[0])
                                    CurCordY = float(centerPoint[1])
                                    if MinCordX <= CurCordX and MaxCordX > CurCordX:
                                        if MinCordY <= CurCordY and MaxCordY > CurCordY:                
                                            if len(j) == 1:
                                                out = '0' + j
                                            else:
                                                out = j
                                            try:
                                                element.get_Parameter(guid_section_param).Set(out)
                                            except:
                                                ui.forms.Alert("У элемента нет параметров, обратитесь в BIM- отдел", title="Ошибка")
                                                script.exit()
                                            OutOfModelFlag = False
                                            ElFoungSectLvlFlag = True
                                            errorCount -=1
                                            breakFlag = True
                                            break
                if OutOfModelFlag:
                    output.print_md('Элемент **{} с id {}** так и не нашел основу:( Проверьте вручную!'.format(element.Name,output.linkify(element.Id)))
                    OutOfModelFlag = True

if flagOut:
    # if dialog_out == 'да':
    if errorCount == 0 and len(errorEl)==0:
        ui.forms.Alert('Аналитически определено положение {} элементов.'.format(len(OutOfModelElList)), title="Завершено")
    else:
        for err in errorEl:
            output.print_md('Элемент **{} с id {}** Либо левый, либо имеет и систему и этаж!'.format(err.Name,output.linkify(err.Id)))
        print('Так и не решилось ошибок: - {} шт.'.format(errorCount+len(errorEl)))



